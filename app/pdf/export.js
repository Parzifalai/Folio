// Сборка итогового PDF: исходные страницы + всё, что нарисовано поверх.
// Объекты хранятся в координатах «как видно на экране» (пункты, ось Y вниз),
// а в PDF переводятся одной матрицей на страницу — так поворот страниц учитывается сам.

const FONT_FILES = {
  arial: ['arial', 'arialbd', 'ariali', 'arialbi'],
  times: ['times', 'timesbd', 'timesi', 'timesbi'],
  courier: ['cour', 'courbd', 'couri', 'courbi'],
};
const fontBytesCache = {};

function fontFile(o) {
  return FONT_FILES[o.font || 'arial'][(o.bold ? 1 : 0) + (o.italic ? 2 : 0)];
}

function hexRgb(hex) {
  const n = parseInt(hex.slice(1), 16);
  return PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
}

function dataUrlBytes(url) {
  const b = atob(url.slice(url.indexOf(',') + 1));
  const a = new Uint8Array(b.length);
  for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i);
  return a;
}

// Геометрия, общая для экрана и PDF.
const Geo = {
  pathD(pts) {
    const n = pts.length / 2;
    if (n === 1) return `M${pts[0]} ${pts[1]}L${pts[0] + 0.01} ${pts[1]}`;
    if (n === 2) return `M${pts[0]} ${pts[1]}L${pts[2]} ${pts[3]}`;
    let d = `M${pts[0].toFixed(2)} ${pts[1].toFixed(2)}`;
    for (let i = 1; i < n - 1; i++) {
      const x = pts[i * 2], y = pts[i * 2 + 1];
      const mx = (x + pts[i * 2 + 2]) / 2, my = (y + pts[i * 2 + 3]) / 2;
      d += `Q${x.toFixed(2)} ${y.toFixed(2)} ${mx.toFixed(2)} ${my.toFixed(2)}`;
    }
    d += `L${pts[pts.length - 2].toFixed(2)} ${pts[pts.length - 1].toFixed(2)}`;
    return d;
  },
  arrow(o) {
    const dx = o.x2 - o.x1, dy = o.y2 - o.y1;
    const len = Math.hypot(dx, dy) || 1;
    const ux = dx / len, uy = dy / len;
    const head = Math.min(len * 0.6, Math.max(9, o.width * 4.2));
    const half = head * 0.42;
    const bx = o.x2 - ux * head, by = o.y2 - uy * head;
    return {
      line: { x1: o.x1, y1: o.y1, x2: o.x2 - ux * head * 0.7, y2: o.y2 - uy * head * 0.7 },
      head: `M${o.x2} ${o.y2}L${bx - uy * half} ${by + ux * half}L${bx + uy * half} ${by - ux * half}Z`,
    };
  },
  // Положение строк текста: верх блока → базовая линия каждой строки.
  baselines(o, asc, desc) {
    const lh = o.size * 1.2;
    const half = (lh - (asc + desc) * o.size) / 2;
    return o.text.split('\n').map((line, i) => ({ line, base: o.y + i * lh + half + asc * o.size }));
  },
};

// Та же матрица, что строит pdf.js для страницы (масштаб 1): PDF → экран.
function viewMatrix(view, rotation) {
  const cx = (view[2] + view[0]) / 2, cy = (view[3] + view[1]) / 2;
  const [a, b, c, d] = { 0: [1, 0, 0, -1], 90: [0, 1, 1, 0], 180: [-1, 0, 0, 1], 270: [0, -1, -1, 0] }[rotation];
  const ox = a === 0 ? Math.abs(cy - view[1]) : Math.abs(cx - view[0]);
  const oy = a === 0 ? Math.abs(cx - view[0]) : Math.abs(cy - view[1]);
  const h = a === 0 ? view[2] - view[0] : view[3] - view[1];
  return { m: [a, b, c, d, ox - a * cx - c * cy, oy - b * cx - d * cy], height: h };
}

function toPdfPoint(m, x, y) {
  const det = m[0] * m[3] - m[1] * m[2];
  return [(x * m[3] - y * m[2] + m[2] * m[5] - m[4] * m[3]) / det, (-x * m[1] + y * m[0] + m[4] * m[1] - m[5] * m[0]) / det];
}

async function getFontBytes(name) {
  if (!fontBytesCache[name]) fontBytesCache[name] = await window.folio.font(name);
  return fontBytesCache[name];
}

async function buildPdf(ed, onlyIds) {
  const { PDFDocument, degrees, pushGraphicsState, popGraphicsState, concatTransformationMatrix,
    setLineJoin, LineJoinStyle, LineCapStyle, BlendMode } = PDFLib;

  const pages = onlyIds ? ed.pages.filter((p) => onlyIds.includes(p.id)) : ed.pages;
  const libDocs = {};
  const loadSrc = async (id) => {
    if (!libDocs[id]) {
      libDocs[id] = await PDFDocument.load(ed.source(id).bytes, { ignoreEncryption: true, updateMetadata: false });
    }
    return libDocs[id];
  };

  // Если страницы не переставляли и не удаляли — правим исходный файл на месте:
  // так сохраняются закладки, ссылки, поля форм и свойства документа.
  const s0 = ed.sources[0];
  const inPlace = !onlyIds && ed.sources.length === 1 && s0 && pages.length === s0.count &&
    pages.every((p, i) => p.src === s0.id && p.index === i);

  let out;
  const outPages = [];
  if (inPlace) {
    out = await loadSrc(s0.id);
    outPages.push(...out.getPages());
  } else {
    out = await PDFDocument.create();
    const bySrc = new Map();
    pages.forEach((p, i) => {
      if (p.src == null) return;
      if (!bySrc.has(p.src)) bySrc.set(p.src, []);
      bySrc.get(p.src).push(i);
    });
    const copied = [];
    for (const [sid, idxs] of bySrc) {
      const d = await loadSrc(sid);
      const cp = await out.copyPages(d, idxs.map((i) => pages[i].index));
      idxs.forEach((i, k) => (copied[i] = cp[k]));
    }
    pages.forEach((p, i) => {
      if (p.src != null) outPages.push(out.addPage(copied[i]));
      else outPages.push(out.addPage([p.view[2] - p.view[0], p.view[3] - p.view[1]]));
    });
  }
  out.registerFontkit(fontkit);

  const fonts = {};
  const images = {};
  const getFont = async (o) => {
    const f = fontFile(o);
    if (!fonts[f]) fonts[f] = await out.embedFont(await getFontBytes(f), { subset: true });
    return fonts[f];
  };
  const getImage = async (id) => {
    if (!images[id]) {
      const im = ed.imgs.get(id);
      const bytes = dataUrlBytes(im.url);
      images[id] = im.url.startsWith('data:image/jpeg') ? await out.embedJpg(bytes) : await out.embedPng(bytes);
    }
    return images[id];
  };

  for (let i = 0; i < pages.length; i++) {
    const p = pages[i];
    const pg = outPages[i];
    const total = (p.baseRot + p.rot) % 360;
    pg.setRotation(degrees(total));
    const objs = p.objs.filter((o) => !o.pristine); // нетронутое содержимое PDF уже в файле
    if (!objs.length) continue;

    const vm = viewMatrix(p.view, total);
    const H = vm.height;
    const map = (x, yUp) => toPdfPoint(vm.m, x, H - yUp);
    const o0 = map(0, 0), ox = map(1, 0), oy = map(0, 1);
    pg.pushOperators(
      pushGraphicsState(),
      concatTransformationMatrix(ox[0] - o0[0], ox[1] - o0[1], oy[0] - o0[0], oy[1] - o0[1], o0[0], o0[1]),
      setLineJoin(LineJoinStyle.Round),
    );

    const of = (t) => objs.filter((o) => t.includes(o.t));
    for (const o of of(['whiteout'])) {
      pg.drawRectangle({ x: o.x, y: H - o.y - o.h, width: o.w, height: o.h, color: hexRgb(o.color) });
    }
    for (const o of of(['marker'])) {
      pg.drawSvgPath(Geo.pathD(o.pts), {
        x: 0, y: H, borderColor: hexRgb(o.color), borderWidth: o.width,
        borderLineCap: LineCapStyle.Round, blendMode: BlendMode.Multiply,
      });
    }
    for (const o of of(['image'])) {
      pg.drawImage(await getImage(o.img), { x: o.x, y: H - o.y - o.h, width: o.w, height: o.h });
    }
    for (const o of of(['pen', 'rect', 'ellipse', 'line', 'arrow'])) {
      const c = hexRgb(o.color);
      if (o.t === 'pen') {
        pg.drawSvgPath(Geo.pathD(o.pts), { x: 0, y: H, borderColor: c, borderWidth: o.width, borderLineCap: LineCapStyle.Round });
      } else if (o.t === 'rect') {
        pg.drawRectangle({ x: o.x, y: H - o.y - o.h, width: o.w, height: o.h, borderColor: c, borderWidth: o.width });
      } else if (o.t === 'ellipse') {
        pg.drawEllipse({ x: o.x + o.w / 2, y: H - o.y - o.h / 2, xScale: o.w / 2, yScale: o.h / 2, borderColor: c, borderWidth: o.width });
      } else {
        const g = o.t === 'arrow' ? Geo.arrow(o) : { line: o };
        pg.drawLine({
          start: { x: g.line.x1, y: H - g.line.y1 }, end: { x: g.line.x2, y: H - g.line.y2 },
          thickness: o.width, color: c, lineCap: LineCapStyle.Round,
        });
        if (g.head) pg.drawSvgPath(g.head, { x: 0, y: H, color: c });
      }
    }
    for (const o of of(['text'])) {
      if (!o.text.trim()) continue;
      const font = await getFont(o);
      const fk = font.embedder && font.embedder.font;
      const asc = fk ? fk.ascent / fk.unitsPerEm : 0.9;
      const desc = fk ? Math.abs(fk.descent) / fk.unitsPerEm : 0.21;
      for (const { line, base } of Geo.baselines(o, asc, desc)) {
        if (line) pg.drawText(line, { x: o.x, y: H - base, size: o.size, font, color: hexRgb(o.color) });
      }
    }
    pg.pushOperators(popGraphicsState());
  }

  return out.save();
}
