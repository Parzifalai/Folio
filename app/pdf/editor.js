// Редактор PDF: просмотр, правка текста, рисование, страницы.
pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

// pdf.js дорисовывает страницу по кадрам экрана. Если Windows кадры не присылает
// (монитор уснул, окно перекрыто), продолжаем по таймеру, чтобы отрисовка не зависла.
(() => {
  const raf = window.requestAnimationFrame.bind(window);
  window.requestAnimationFrame = (cb) => {
    let done = false;
    const run = (t) => { if (!done) { done = true; cb(t); } };
    const id = raf(run);
    setTimeout(() => run(performance.now()), 100);
    return id;
  };
})();

const PT = 96 / 72; // экранных пикселей в пункте при масштабе 100%
const INK = ['#1d2125', '#2f5bea', '#d93025', '#188038', '#f29900', '#ffffff'];
const HL = ['#ffe14d', '#a8f07a', '#ff9ecf', '#8fd3ff'];
const COVER = ['#ffffff', '#1d2125'];
const FONT_CSS = { arial: 'Arial, sans-serif', times: '"Times New Roman", serif', courier: '"Courier New", monospace' };
const FONT_NAME = { arial: 'Arial', times: 'Times', courier: 'Courier' };
const FONT_METRICS = { arial: [0.905, 0.212], times: [0.891, 0.216], courier: [0.833, 0.3] };
const ZOOMS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5];
const A4 = [0, 0, 595.28, 841.89];

// Нажата ли сейчас кнопка мыши — чтобы после await не начать перетаскивание «вслепую».
let PTR_DOWN = false;
window.addEventListener('pointerdown', () => (PTR_DOWN = true), true);
window.addEventListener('pointerup', () => (PTR_DOWN = false), true);
window.addEventListener('pointercancel', () => (PTR_DOWN = false), true);

const Settings = (() => {
  const def = {
    pen: { color: INK[1], width: 2 },
    marker: { color: HL[0], width: 14 },
    shape: { kind: 'rect', color: INK[2], width: 2 },
    text: { font: 'arial', size: 12, color: INK[0], bold: false, italic: false },
    whiteout: { color: '#ffffff' },
  };
  let s = def;
  try {
    const saved = JSON.parse(localStorage.getItem('folio.tools') || 'null');
    if (saved) for (const k in def) s[k] = { ...def[k], ...(saved[k] || {}) };
  } catch {}
  return {
    get: (k) => s[k],
    set(k, patch) {
      Object.assign(s[k], patch);
      try { localStorage.setItem('folio.tools', JSON.stringify(s)); } catch {}
    },
  };
})();

let uid = 1;
const nid = () => uid++;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const el = (tag, cls, html) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
};
const SVGNS = 'http://www.w3.org/2000/svg';
const svgEl = () => {
  const s = document.createElementNS(SVGNS, 'svg');
  s.setAttribute('class', 'layer');
  return s;
};

const measureCtx = document.createElement('canvas').getContext('2d');
function measureText(o) {
  measureCtx.font = `${o.italic ? 'italic ' : ''}${o.bold ? 'bold ' : ''}${o.size}px ${FONT_CSS[o.font]}`;
  if ('fontKerning' in measureCtx) measureCtx.fontKerning = 'none';
  const lines = o.text.split('\n');
  let w = 0;
  for (const l of lines) w = Math.max(w, measureCtx.measureText(l).width);
  return { w: Math.max(w, o.size * 0.3), h: lines.length * o.size * 1.2 };
}

function bbox(o) {
  switch (o.t) {
    case 'text': {
      if (o.pristine && o.nbox) return o.nbox;
      const m = measureText(o);
      return { x: o.x, y: o.y, w: m.w, h: m.h };
    }
    case 'line':
    case 'arrow': {
      const x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2);
      return { x, y, w: Math.abs(o.x2 - o.x1), h: Math.abs(o.y2 - o.y1) };
    }
    case 'pen':
    case 'marker': {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      for (let i = 0; i < o.pts.length; i += 2) {
        x0 = Math.min(x0, o.pts[i]); x1 = Math.max(x1, o.pts[i]);
        y0 = Math.min(y0, o.pts[i + 1]); y1 = Math.max(y1, o.pts[i + 1]);
      }
      const r = o.width / 2;
      return { x: x0 - r, y: y0 - r, w: x1 - x0 + 2 * r, h: y1 - y0 + 2 * r };
    }
    default:
      return { x: o.x, y: o.y, w: o.w, h: o.h };
  }
}

function distSeg(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1, dy = y2 - y1;
  const l = dx * dx + dy * dy;
  let t = l ? ((px - x1) * dx + (py - y1) * dy) / l : 0;
  t = clamp(t, 0, 1);
  return Math.hypot(px - x1 - t * dx, py - y1 - t * dy);
}

function hits(o, x, y, tol) {
  if (o.t === 'pen' || o.t === 'marker') {
    const r = o.width / 2 + tol;
    const p = o.pts;
    if (p.length === 2) return Math.hypot(x - p[0], y - p[1]) <= r;
    for (let i = 0; i < p.length - 2; i += 2) if (distSeg(x, y, p[i], p[i + 1], p[i + 2], p[i + 3]) <= r) return true;
    return false;
  }
  if (o.t === 'line' || o.t === 'arrow') return distSeg(x, y, o.x1, o.y1, o.x2, o.y2) <= o.width / 2 + tol;
  if (o.t === 'rect' || o.t === 'ellipse') {
    // Пустые фигуры ловятся за контур, чтобы не мешать кликам внутри.
    const b = bbox(o);
    const inOuter = x >= b.x - tol && x <= b.x + b.w + tol && y >= b.y - tol && y <= b.y + b.h + tol;
    const m = o.width / 2 + tol;
    const inInner = x > b.x + m && x < b.x + b.w - m && y > b.y + m && y < b.y + b.h - m;
    if (o.t === 'rect') return inOuter && !inInner;
    const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
    const e = (rx, ry) => ((x - cx) / rx) ** 2 + ((y - cy) / ry) ** 2;
    return e(b.w / 2 + m, b.h / 2 + m) <= 1 && e(Math.max(1, b.w / 2 - m), Math.max(1, b.h / 2 - m)) >= 1;
  }
  const b = bbox(o);
  return x >= b.x - tol && x <= b.x + b.w + tol && y >= b.y - tol && y <= b.y + b.h + tol;
}

function translate(o, dx, dy) {
  if ('x' in o) { o.x += dx; o.y += dy; }
  if ('x1' in o) { o.x1 += dx; o.y1 += dy; o.x2 += dx; o.y2 += dy; }
  if (o.pts) for (let i = 0; i < o.pts.length; i += 2) { o.pts[i] += dx; o.pts[i + 1] += dy; }
}

// Растянуть объект из прямоугольника a в прямоугольник b (o0 — исходная копия).
function fitTo(o, o0, a, b) {
  const sx = a.w ? b.w / a.w : 1, sy = a.h ? b.h / a.h : 1;
  const X = (x) => b.x + (x - a.x) * sx, Y = (y) => b.y + (y - a.y) * sy;
  if (o.t === 'text') {
    o.size = clamp(Math.round(o0.size * sy * 2) / 2, 4, 400);
    o.x = X(o0.x); o.y = Y(o0.y);
  } else if (o.pts) {
    for (let i = 0; i < o.pts.length; i += 2) { o.pts[i] = X(o0.pts[i]); o.pts[i + 1] = Y(o0.pts[i + 1]); }
  } else if ('x1' in o) {
    o.x1 = X(o0.x1); o.y1 = Y(o0.y1); o.x2 = X(o0.x2); o.y2 = Y(o0.y2);
  } else {
    o.x = b.x; o.y = b.y; o.w = b.w; o.h = b.h;
  }
}

function hexOf(r, g, b) {
  return '#' + [r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('');
}

function readAsDataUrl(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(blob);
  });
}

// Картинку приводим к PNG/JPEG (их понимает PDF), учитывая поворот с камеры.
async function normalizeImage(bytes, name) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  const type = ext === 'png' ? 'image/png' : /jpe?g/.test(ext) ? 'image/jpeg' : 'image/' + ext;
  const bmp = await createImageBitmap(new Blob([bytes], { type }));
  const max = 3200;
  const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const w = Math.round(bmp.width * k), h = Math.round(bmp.height * k);
  if (ext === 'png' && k === 1) {
    return { url: await readAsDataUrl(new Blob([bytes], { type })), w, h };
  }
  const c = document.createElement('canvas');
  c.width = w; c.height = h;
  const g = c.getContext('2d');
  const jpeg = type === 'image/jpeg';
  if (jpeg) { g.fillStyle = '#fff'; g.fillRect(0, 0, w, h); }
  g.drawImage(bmp, 0, 0, w, h);
  return { url: c.toDataURL(jpeg ? 'image/jpeg' : 'image/png', 0.92), w, h };
}

class PdfEditor {
  constructor(host, { onChange }) {
    this.onChange = onChange;
    this.sources = [];
    this.pages = [];
    this.imgs = new Map();
    this.undoStack = [];
    this.redoStack = [];
    this.dirty = false;
    this.zoom = 1;
    this.tool = 'select';
    this.sel = null; // { pid, oid }
    this.editing = null;
    this.els = new Map();
    this.thumbEls = new Map();
    this.picked = new Set();
    this.cur = null;
    this.runCache = new Map();
    this.find = null;
    PdfEditor.all.add(this);

    this.root = el('div', 'doc');
    this.side = el('aside', 'side');
    this.thumbs = el('div', 'thumbs');
    this.sideGrip = el('div', 'side-grip');
    this.sideGrip.dataset.tip = 'Перетащите панель влево или вправо';
    this.side.append(this.sideGrip, this.thumbs);
    this.viewer = el('div', 'viewer t-select');
    this.viewer.tabIndex = -1;
    this.column = el('div', 'column');
    this.viewer.append(this.column);
    this.root.append(this.side, this.viewer);
    this.hud = el('div', 'hud');
    this.root.append(this.hud);
    host.append(this.root);

    this.buildDock();
    this.bindViewer();
    this.bindThumbs();
    this.bindSideGrip();

    this.ro = new ResizeObserver(() => {
      if (!this.viewer.clientWidth) return;
      if (this.fitPending) { this.fitPending = false; this.fitWidth(); this.viewer.scrollTop = 0; }
      else if (this.fitMode) this.fitWidth();
      this.checkVisible();
    });
    this.ro.observe(this.viewer);
    this.ro.observe(this.thumbs);
    this.thumbs.addEventListener('scroll', () => this.checkVisible(), { passive: true });
  }

  get S() { return this.zoom * PT; }
  source(id) { return this.sources.find((s) => s.id === id); }
  page(pid) { return this.pages.find((p) => p.id === +pid); }
  dims(p) {
    const vw = p.view[2] - p.view[0], vh = p.view[3] - p.view[1];
    return (p.baseRot + p.rot) % 180 ? { w: vh, h: vw } : { w: vw, h: vh };
  }
  total(p) { return (p.baseRot + p.rot) % 360; }

  // ---------- Загрузка ----------
  async addSource(bytes) {
    const pdf = await pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
    const src = { id: nid(), bytes, pdf, count: pdf.numPages, pg: new Map() };
    this.sources.push(src);
    const out = [];
    for (let i = 0; i < pdf.numPages; i++) {
      const pg = await this.pdfPage(src, i);
      out.push({ id: nid(), src: src.id, index: i, rot: 0, baseRot: pg.rotate % 360, view: pg.view.slice(), objs: [] });
    }
    return out;
  }

  pdfPage(src, i) {
    if (!src.pg.has(i)) src.pg.set(i, src.pdf.getPage(i + 1));
    return src.pg.get(i);
  }

  async imagePage(bytes, name) {
    const im = await normalizeImage(bytes, name);
    const imgId = nid();
    this.imgs.set(imgId, im);
    const land = im.w > im.h;
    const view = land ? [0, 0, A4[3], A4[2]] : A4.slice();
    const W = view[2], H = view[3];
    const k = Math.min(W / im.w, H / im.h);
    const w = im.w * k, h = im.h * k;
    return { id: nid(), src: null, index: 0, rot: 0, baseRot: 0, view,
      objs: [{ id: nid(), t: 'image', img: imgId, x: (W - w) / 2, y: (H - h) / 2, w, h }] };
  }

  async pagesFromFiles(files) {
    const out = [];
    for (const f of files) {
      if (/\.pdf$/i.test(f.name)) out.push(...(await this.addSource(f.bytes)));
      else out.push(await this.imagePage(f.bytes, f.name));
    }
    return out;
  }

  async open(files) {
    const loading = el('div', 'loading', 'Открываю…');
    this.viewer.append(loading);
    try {
      this.pages = await this.pagesFromFiles(files);
    } finally {
      loading.remove();
    }
    this.fitPending = true;
    this.rebuild();
    if (this.viewer.clientWidth) { this.fitPending = false; this.fitWidth(); }
    this.viewer.scrollTop = 0;
  }

  blank() {
    this.pages = [{ id: nid(), src: null, index: 0, rot: 0, baseRot: 0, view: A4.slice(), objs: [] }];
    this.fitPending = true;
    this.rebuild();
  }

  // ---------- История ----------
  // Снимок для отмены. Содержимое PDF, выбранное «на пробу», в историю не попадает.
  snapshot() {
    const s = structuredClone(this.pages);
    for (const p of s) p.objs = p.objs.filter((o) => !o.pristine);
    return s;
  }
  checkpoint(snap) {
    this.undoStack.push(snap || this.snapshot());
    if (this.undoStack.length > 300) this.undoStack.shift();
    this.redoStack = [];
    this.touch();
  }
  touch() {
    this.dirty = true;
    this.onChange && this.onChange();
  }
  undo() {
    this.endEdit();
    if (!this.undoStack.length) return;
    this.redoStack.push(this.snapshot());
    this.pages = this.undoStack.pop();
    this.afterRestore();
  }
  redo() {
    this.endEdit();
    if (!this.redoStack.length) return;
    this.undoStack.push(this.snapshot());
    this.pages = this.redoStack.pop();
    this.afterRestore();
  }
  afterRestore() {
    this.sel = null;
    this.rebuild();
    this.touch();
    this.updateCtx();
  }

  // ---------- Раскладка страниц ----------
  rebuild() {
    const keep = new Set(this.pages.map((p) => p.id));
    for (const [pid, rec] of this.els) {
      if (!keep.has(pid)) { rec.el.remove(); this.els.delete(pid); }
    }
    for (const p of this.pages) {
      let rec = this.els.get(p.id);
      if (!rec) rec = this.makePageEl(p);
      this.column.append(rec.el);
    }
    this.layout();
    this.rebuildThumbs();
    this.drawSel();
    this.updateHud(false);
  }

  makePageEl(p) {
    const e = el('div', 'page');
    e.dataset.pid = p.id;
    const canvas = document.createElement('canvas');
    const white = svgEl();
    const mult = svgEl();
    mult.classList.add('mult');
    const imgs = el('div', 'layer');
    const draw = svgEl();
    const text = el('div', 'layer');
    const marks = el('div', 'layer marks');
    marks.style.transform = 'none';
    e.append(canvas, white, mult, imgs, draw, text, marks);
    const rec = { el: e, canvas, white, mult, imgs, draw, text, marks, renderedKey: '', visible: false, task: null };
    this.els.set(p.id, rec);
    return rec;
  }

  layout() {
    const S = this.S;
    for (const p of this.pages) {
      const rec = this.els.get(p.id);
      const { w, h } = this.dims(p);
      rec.el.style.width = w * S + 'px';
      rec.el.style.height = h * S + 'px';
      for (const s of [rec.white, rec.mult, rec.draw]) {
        s.setAttribute('width', w * S);
        s.setAttribute('height', h * S);
        s.setAttribute('viewBox', `0 0 ${w} ${h}`);
      }
      for (const d of [rec.imgs, rec.text]) {
        d.style.width = w + 'px';
        d.style.height = h + 'px';
        d.style.transform = `scale(${S})`;
      }
      this.renderObjs(p);
    }
    this.drawSel();
    this.drawFind();
    this.checkVisible();
  }

  // Какие страницы рядом с экраном: их рисуем, далёкие — отпускаем.
  checkVisible() {
    clearTimeout(this.visFrame);
    this.visFrame = setTimeout(() => {
      const vr = this.viewer.getBoundingClientRect();
      const m = Math.max(600, vr.height);
      for (const [pid, rec] of this.els) {
        const r = rec.el.getBoundingClientRect();
        const vis = vr.height > 0 && r.bottom > vr.top - m && r.top < vr.bottom + m;
        rec.visible = vis;
        if (vis) this.queueRender(pid);
        else if (this.pages.length > 24 && rec.renderedKey) {
          rec.renderedKey = '';
          if (rec.task) { rec.task.cancel(); rec.task = null; }
          const c = document.createElement('canvas');
          rec.canvas.replaceWith(c);
          rec.canvas = c;
        }
      }
      if (this.root.classList.contains('side-open')) {
        const tr = this.thumbs.getBoundingClientRect();
        for (const [pid, t] of this.thumbEls) {
          const r = t.getBoundingClientRect();
          if (r.bottom > tr.top - 300 && r.top < tr.bottom + 300) this.renderThumb(pid);
        }
      }
    }, 16);
  }

  queueRender(pid) {
    clearTimeout(this.renderTimer);
    this.renderQueue = this.renderQueue || new Set();
    this.renderQueue.add(pid);
    this.renderTimer = setTimeout(() => {
      const q = [...this.renderQueue];
      this.renderQueue.clear();
      q.forEach((id) => this.renderPage(id));
    }, 60);
  }

  async renderPage(pid) {
    const p = this.page(pid);
    const rec = this.els.get(pid);
    if (!p || !rec || !rec.visible) return;
    const dpr = window.devicePixelRatio || 1;
    const key = `${this.S.toFixed(4)}|${this.total(p)}|${dpr}`;
    if (rec.renderedKey === key || (rec.task && rec.taskKey === key)) return;
    if (rec.task) { rec.task.cancel(); rec.task = null; }
    const { w, h } = this.dims(p);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.floor(w * this.S * dpr));
    c.height = Math.max(1, Math.floor(h * this.S * dpr));
    if (p.src == null) {
      const g = c.getContext('2d');
      g.fillStyle = '#fff';
      g.fillRect(0, 0, c.width, c.height);
    } else {
      const src = this.source(p.src);
      const pg = await this.pdfPage(src, p.index);
      const vp = pg.getViewport({ scale: this.S * dpr, rotation: this.total(p) });
      const task = pg.render({ canvasContext: c.getContext('2d'), viewport: vp, annotationMode: pdfjsLib.AnnotationMode.ENABLE_FORMS });
      rec.task = task;
      rec.taskKey = key;
      try {
        await task.promise;
      } catch (err) {
        if (err && err.name === 'RenderingCancelledException') return;
        console.warn(err);
      }
      if (rec.task !== task) return;
      rec.task = null;
    }
    rec.canvas.replaceWith(c);
    rec.canvas = c;
    rec.renderedKey = key;
    // Заранее разбираем текст и картинки страницы — чтобы выделение срабатывало сразу.
    if (p.src != null) setTimeout(() => { this.runs(p); this.images(p); }, 50);
  }

  // ---------- Объекты на странице ----------
  renderObjs(p) {
    const rec = this.els.get(p.id);
    if (!rec) return;
    const { w, h } = this.dims(p);
    let white = '', mult = '', draw = '';
    const editingEl = this.editing && this.editing.pid === p.id ? this.editing.el : null;
    rec.imgs.textContent = '';
    for (const node of [...rec.text.children]) if (node !== editingEl) node.remove();
    for (const o of p.objs) {
      if (o.pristine) continue; // не тронутое содержимое PDF видно и так
      if (o.t === 'whiteout') white += `<rect x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" fill="${o.color}"/>`;
      else if (o.t === 'marker') mult += `<path d="${Geo.pathD(o.pts)}" fill="none" stroke="${o.color}" stroke-width="${o.width}" stroke-linecap="round" stroke-linejoin="round"/>`;
      else if (o.t === 'pen') draw += `<path d="${Geo.pathD(o.pts)}" fill="none" stroke="${o.color}" stroke-width="${o.width}" stroke-linecap="round" stroke-linejoin="round"/>`;
      else if (o.t === 'rect') draw += `<rect x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" fill="none" stroke="${o.color}" stroke-width="${o.width}" stroke-linejoin="round"/>`;
      else if (o.t === 'ellipse') draw += `<ellipse cx="${o.x + o.w / 2}" cy="${o.y + o.h / 2}" rx="${o.w / 2}" ry="${o.h / 2}" fill="none" stroke="${o.color}" stroke-width="${o.width}"/>`;
      else if (o.t === 'line' || o.t === 'arrow') {
        const g = o.t === 'arrow' ? Geo.arrow(o) : { line: o };
        draw += `<line x1="${g.line.x1}" y1="${g.line.y1}" x2="${g.line.x2}" y2="${g.line.y2}" stroke="${o.color}" stroke-width="${o.width}" stroke-linecap="round"/>`;
        if (g.head) draw += `<path d="${g.head}" fill="${o.color}"/>`;
      } else if (o.t === 'image') {
        const im = el('img', 'obj-img');
        im.src = this.imgs.get(o.img).url;
        im.draggable = false;
        Object.assign(im.style, { left: o.x + 'px', top: o.y + 'px', width: o.w + 'px', height: o.h + 'px' });
        rec.imgs.append(im);
      } else if (o.t === 'text') {
        if (editingEl && this.editing.oid === o.id) { this.styleText(editingEl, o); continue; }
        const t = el('div', 'obj-text');
        t.textContent = o.text;
        this.styleText(t, o);
        rec.text.append(t);
      }
    }
    rec.white.innerHTML = white;
    rec.mult.innerHTML = mult;
    rec.draw.innerHTML = draw;
    void w; void h;
  }

  styleText(t, o) {
    Object.assign(t.style, {
      left: o.x + 'px', top: o.y + 'px', fontSize: o.size + 'px', fontFamily: FONT_CSS[o.font],
      fontWeight: o.bold ? '700' : '400', fontStyle: o.italic ? 'italic' : 'normal', color: o.color,
    });
  }

  refresh(p) {
    this.renderObjs(p);
    this.drawSel();
  }

  findObj(pid, oid) {
    const p = this.page(pid);
    return p ? p.objs.find((o) => o.id === oid) : null;
  }

  removeObj(p, o) {
    p.objs = p.objs.filter((x) => x !== o && !(o.cover && x.id === o.cover));
  }

  hitTest(p, x, y) {
    const tol = 4 / this.S;
    const order = ['text', 'pen', 'rect', 'ellipse', 'line', 'arrow', 'image', 'marker', 'whiteout'];
    for (const t of order) {
      for (let i = p.objs.length - 1; i >= 0; i--) {
        const o = p.objs[i];
        if (o.t === t && !(o.pristine && o.t === 'whiteout') && hits(o, x, y, tol)) return o;
      }
    }
    return null;
  }

  // ---------- Выделение ----------
  // sel = { pid, ids } — один или несколько объектов на одной странице.
  select(pid, ids) {
    const next = ids && ids.length ? { pid: +pid, ids: [...new Set(ids)] } : null;
    this.dropPristine(next);
    this.sel = next;
    this.drawSel();
    this.updateCtx();
  }

  selectedObjs() {
    if (!this.sel) return null;
    const p = this.page(this.sel.pid);
    const objs = p ? this.sel.ids.map((id) => p.objs.find((o) => o.id === id)).filter(Boolean) : [];
    if (!objs.length) { this.sel = null; return null; }
    return { p, objs };
  }

  // Ровно один выбранный объект — для правки текста и концов линий.
  selected() {
    const s = this.selectedObjs();
    return s && s.objs.length === 1 ? { p: s.p, o: s.objs[0] } : null;
  }

  isSelected(p, o) { return !!this.sel && this.sel.pid === p.id && this.sel.ids.includes(o.id); }

  toggleSel(p, objs) {
    const cur = this.sel && this.sel.pid === p.id ? [...this.sel.ids] : [];
    for (const o of objs) {
      const i = cur.indexOf(o.id);
      if (i >= 0) cur.splice(i, 1); else cur.push(o.id);
    }
    this.select(p.id, cur);
  }

  clearSel() {
    this.endEdit();
    this.dropPristine(null);
    this.sel = null;
  }

  unionBox(objs) {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const o of objs) {
      const b = bbox(o);
      x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y);
      x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h);
    }
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  // Содержимое PDF, выбранное «на пробу»: пока его не двигали и не меняли, файл не трогаем.
  dropPristine(keep) {
    const keepIds = new Set(keep ? keep.ids : []);
    for (const p of this.pages) {
      if (!p.objs.some((o) => o.pristine)) continue;
      const keepCovers = new Set(p.objs.filter((o) => keepIds.has(o.id) && o.cover).map((o) => o.cover));
      const n = p.objs.length;
      p.objs = p.objs.filter((o) => !o.pristine || keepIds.has(o.id) || keepCovers.has(o.id));
      if (p.objs.length !== n) this.renderObjs(p);
    }
  }

  // Объект тронули — теперь это настоящая правка, вместе с замазкой под ним.
  unPristine(p, objs) {
    for (const o of objs) {
      delete o.pristine;
      if (o.cover) {
        const c = p.objs.find((x) => x.id === o.cover);
        if (c) delete c.pristine;
      }
    }
  }

  drawSel() {
    for (const rec of this.els.values()) rec.marks.querySelectorAll('.sel,.sel-item,.handle').forEach((n) => n.remove());
    const s = this.selectedObjs();
    if (!s) return;
    const rec = this.els.get(s.p.id);
    if (!rec) return;
    const S = this.S;
    const box = (b, cls) => {
      const d = el('div', cls);
      Object.assign(d.style, { left: b.x * S + 'px', top: b.y * S + 'px', width: b.w * S + 'px', height: b.h * S + 'px' });
      rec.marks.append(d);
    };
    if (s.objs.length > 1) for (const o of s.objs) box(bbox(o), 'sel-item');
    const b = this.unionBox(s.objs);
    box(b, 'sel' + (this.editing ? ' editing' : ''));
    if (this.editing) return;
    const one = s.objs.length === 1 ? s.objs[0] : null;
    const hs = one && (one.t === 'line' || one.t === 'arrow')
      ? [[one.x1, one.y1, 'p1'], [one.x2, one.y2, 'p2']]
      : [[b.x, b.y, 'nw'], [b.x + b.w, b.y, 'ne'], [b.x, b.y + b.h, 'sw'], [b.x + b.w, b.y + b.h, 'se']];
    for (const [x, y, k] of hs) {
      const hd = el('div', 'handle');
      hd.dataset.h = k;
      hd.style.left = x * S + 'px';
      hd.style.top = y * S + 'px';
      hd.style.cursor = k === 'nw' || k === 'se' ? 'nwse-resize' : k === 'ne' || k === 'sw' ? 'nesw-resize' : 'move';
      rec.marks.append(hd);
    }
  }

  // ---------- Содержимое самого PDF: картинки ----------
  async images(p) {
    if (p.src == null) return [];
    const key = `img:${p.src}:${p.index}:${this.total(p)}`;
    if (this.runCache.has(key)) return this.runCache.get(key);
    const prom = (async () => {
      const pg = await this.pdfPage(this.source(p.src), p.index);
      const vp = pg.getViewport({ scale: 1, rotation: this.total(p) });
      const ops = await pg.getOperatorList();
      const O = pdfjsLib.OPS, U = pdfjsLib.Util;
      const { w: W, h: H } = this.dims(p);
      const I = [1, 0, 0, 1, 0, 0];
      let ctm = I;
      const stack = [];
      const out = [];
      for (let i = 0; i < ops.fnArray.length; i++) {
        const fn = ops.fnArray[i], a = ops.argsArray[i];
        if (fn === O.save) stack.push(ctm);
        else if (fn === O.restore) ctm = stack.pop() || I;
        else if (fn === O.transform) ctm = U.transform(ctm, a);
        else if (fn === O.paintFormXObjectBegin) { stack.push(ctm); if (a && a[0]) ctm = U.transform(ctm, a[0]); }
        else if (fn === O.paintFormXObjectEnd) ctm = stack.pop() || I;
        else if (fn === O.beginGroup) { stack.push(ctm); if (a && a[0] && a[0].matrix) ctm = U.transform(ctm, a[0].matrix); }
        else if (fn === O.endGroup) ctm = stack.pop() || I;
        else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject || fn === O.paintImageMaskXObject) {
          const m = U.transform(vp.transform, ctm);
          const xs = [m[4], m[0] + m[4], m[2] + m[4], m[0] + m[2] + m[4]];
          const ys = [m[5], m[1] + m[5], m[3] + m[5], m[1] + m[3] + m[5]];
          const x = Math.max(0, Math.min(...xs)), y = Math.max(0, Math.min(...ys));
          const w = Math.min(W, Math.max(...xs)) - x, h = Math.min(H, Math.max(...ys)) - y;
          if (w < 6 || h < 6 || w * h > W * H * 0.85) continue; // фон-скан целиком не берём
          out.push({ key: `${key}:${i}`, kind: 'image', x, y, w, h });
        }
      }
      return out;
    })().catch(() => []);
    this.runCache.set(key, prom);
    return prom;
  }

  usedNative(p) { return new Set(p.objs.filter((o) => o.run).map((o) => o.run)); }

  async nativeAt(p, x, y, kinds = ['text', 'image']) {
    const used = this.usedNative(p);
    if (kinds.includes('text')) {
      const r = (await this.runs(p)).find((r) => !used.has(r.key) && x >= r.x - 1 && x <= r.x + r.w + 1 && y >= r.y - 1 && y <= r.y + r.h + 1);
      if (r) return { ...r, kind: 'text' };
    }
    if (kinds.includes('image')) {
      const ims = await this.images(p);
      for (let i = ims.length - 1; i >= 0; i--) {
        const m = ims[i];
        if (!used.has(m.key) && x >= m.x && x <= m.x + m.w && y >= m.y && y <= m.y + m.h) return m;
      }
    }
    return null;
  }

  async nativesIn(p, b) {
    const used = this.usedNative(p);
    const inside = (r) => {
      const cx = r.x + r.w / 2, cy = r.y + r.h / 2;
      return !used.has(r.key) && cx >= b.x && cx <= b.x + b.w && cy >= b.y && cy <= b.y + b.h;
    };
    const ims = (await this.images(p)).filter(inside);
    const rs = (await this.runs(p)).filter(inside).map((r) => ({ ...r, kind: 'text' }));
    return [...ims, ...rs];
  }

  // Строки и картинки PDF превращаем в объекты «на пробу» (под ними — замазка цветом фона).
  convertNative(p, list) {
    const out = [];
    for (const n of list) {
      if (n.kind === 'image') {
        const crop = this.cropCanvas(p, n);
        if (!crop) continue;
        const imgId = nid();
        this.imgs.set(imgId, crop);
        const cover = { id: nid(), t: 'whiteout', x: n.x, y: n.y, w: n.w, h: n.h, color: this.sampleRing(p, n), run: n.key, pristine: true };
        const o = { id: nid(), t: 'image', img: imgId, x: n.x, y: n.y, w: n.w, h: n.h, cover: cover.id, pristine: true };
        p.objs.push(cover, o);
        out.push(o);
      } else {
        const { bg, fg } = this.sampleColors(p, n);
        const [asc, desc] = FONT_METRICS[n.font];
        const size = Math.round(n.size * 2) / 2;
        const cover = { id: nid(), t: 'whiteout', x: n.x - 0.8, y: n.y - 0.4, w: n.w + 1.6, h: n.h + 0.8, color: bg, run: n.key, pristine: true };
        const top = n.base - (1.2 * size - (asc + desc) * size) / 2 - asc * size;
        const o = { id: nid(), t: 'text', x: n.x, y: top, text: n.str, font: n.font, size, color: fg, bold: n.bold, italic: n.italic,
          cover: cover.id, orig: n.str, pristine: true, nbox: { x: n.x, y: n.y, w: n.w, h: n.h } };
        p.objs.push(cover, o);
        out.push(o);
      }
    }
    return out;
  }

  pageCanvas(p) {
    const rec = this.els.get(p.id);
    return rec && rec.renderedKey && rec.canvas.width ? rec.canvas : null;
  }

  cropCanvas(p, r) {
    const c = this.pageCanvas(p);
    if (!c) return null;
    const k = c.width / this.dims(p).w;
    const sx = Math.max(0, Math.floor(r.x * k)), sy = Math.max(0, Math.floor(r.y * k));
    const sw = Math.min(c.width - sx, Math.ceil(r.w * k)), sh = Math.min(c.height - sy, Math.ceil(r.h * k));
    if (sw < 2 || sh < 2) return null;
    const out = document.createElement('canvas');
    out.width = sw; out.height = sh;
    out.getContext('2d').drawImage(c, sx, sy, sw, sh, 0, 0, sw, sh);
    return { url: out.toDataURL(sw * sh > 700000 ? 'image/jpeg' : 'image/png', 0.92), w: sw, h: sh };
  }

  // Цвет фона вокруг картинки — им закрываем её старое место.
  sampleRing(p, r) {
    const c = this.pageCanvas(p);
    if (!c) return '#ffffff';
    const k = c.width / this.dims(p).w;
    const pad = Math.max(2, Math.round(3 * k));
    const ix0 = Math.floor(r.x * k), iy0 = Math.floor(r.y * k), ix1 = Math.ceil((r.x + r.w) * k), iy1 = Math.ceil((r.y + r.h) * k);
    const x0 = Math.max(0, ix0 - pad), y0 = Math.max(0, iy0 - pad), x1 = Math.min(c.width, ix1 + pad), y1 = Math.min(c.height, iy1 + pad);
    let d;
    try { d = c.getContext('2d').getImageData(x0, y0, x1 - x0, y1 - y0).data; } catch { return '#ffffff'; }
    const W = x1 - x0, hist = new Map();
    for (let y = y0; y < y1; y += 2) {
      for (let x = x0; x < x1; x += 2) {
        if (x >= ix0 && x < ix1 && y >= iy0 && y < iy1) continue;
        const i = ((y - y0) * W + (x - x0)) * 4;
        const q = ((d[i] >> 3) << 10) | ((d[i + 1] >> 3) << 5) | (d[i + 2] >> 3);
        hist.set(q, (hist.get(q) || 0) + 1);
      }
    }
    let bq = 0x7fff, bn = -1;
    for (const [q, n] of hist) if (n > bn) { bn = n; bq = q; }
    const ch = (v) => Math.min(255, v * 8 + 4);
    const r8 = ch((bq >> 10) & 31), g8 = ch((bq >> 5) & 31), b8 = ch(bq & 31);
    return r8 > 246 && g8 > 246 && b8 > 246 ? '#ffffff' : hexOf(r8, g8, b8);
  }

  // ---------- Текст из PDF (правка существующего) ----------
  async runs(p) {
    if (p.src == null) return [];
    const key = `${p.src}:${p.index}:${this.total(p)}`;
    if (this.runCache.has(key)) return this.runCache.get(key);
    const prom = (async () => {
      const src = this.source(p.src);
      const pg = await this.pdfPage(src, p.index);
      const vp = pg.getViewport({ scale: 1, rotation: this.total(p) });
      const tc = await pg.getTextContent();
      const items = [];
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) {
          if (it.str && items.length) items[items.length - 1].space = true;
          continue;
        }
        const m = pdfjsLib.Util.transform(vp.transform, it.transform);
        if (Math.abs(m[1]) > 0.01 * Math.abs(m[0]) || m[0] <= 0) continue; // только горизонтальный текст
        const size = Math.hypot(m[2], m[3]);
        if (size < 1) continue;
        const st = tc.styles[it.fontName] || {};
        let fname = '';
        try { if (pg.commonObjs.has(it.fontName)) fname = pg.commonObjs.get(it.fontName).name || ''; } catch {}
        const fam = st.fontFamily || '';
        const font = /courier|mono|consol/i.test(fname + fam) ? 'courier'
          : /times|serif|roman|georgia|cambria|garamond|minion|book/i.test(fname) || (fam === 'serif' && !/sans/i.test(fname)) ? 'times' : 'arial';
        items.push({
          str: it.str, x: m[4], base: m[5], size, w: it.width,
          font, bold: /bold|black|heavy|semibold|demi/i.test(fname), italic: /italic|oblique/i.test(fname),
          asc: st.ascent || 0.8, desc: st.descent || -0.2,
        });
      }
      // Склеиваем кусочки одной строки в одну «фразу».
      const runs = [];
      for (const it of items) {
        const last = runs[runs.length - 1];
        if (last && Math.abs(last.base - it.base) < last.size * 0.25 && Math.abs(last.size - it.size) < last.size * 0.15 &&
            it.x >= last.x + last.w - last.size * 0.3 && it.x - (last.x + last.w) < last.size * 1.1) {
          const gap = it.x - (last.x + last.w);
          const needSpace = (last.space || gap > last.size * 0.18) && !/\s$/.test(last.str) && !/^\s/.test(it.str);
          last.str += (needSpace ? ' ' : '') + it.str;
          last.w = it.x + it.w - last.x;
          last.space = it.space;
          last.bold = last.bold && it.bold;
        } else {
          runs.push({ ...it });
        }
      }
      return runs.map((r, i) => {
        const top = r.base - Math.max(r.asc, 0.75) * r.size;
        const bottom = r.base - Math.min(r.desc, -0.18) * r.size;
        return { key: `${key}:${i}`, str: r.str.trim(), x: r.x, y: top, w: r.w, h: bottom - top, base: r.base, size: r.size, font: r.font, bold: r.bold, italic: r.italic };
      });
    })();
    this.runCache.set(key, prom);
    return prom;
  }

  async runAt(p, x, y) {
    const rs = await this.runs(p);
    const used = new Set(p.objs.filter((o) => o.run).map((o) => o.run));
    return rs.find((r) => !used.has(r.key) && x >= r.x - 1 && x <= r.x + r.w + 1 && y >= r.y - 1 && y <= r.y + r.h + 1) || null;
  }

  // Цвета фона и текста под фразой — берём с отрисованной страницы.
  sampleColors(p, r) {
    const rec = this.els.get(p.id);
    const c = rec && rec.canvas;
    const { w } = this.dims(p);
    if (!c || !c.width || !rec.renderedKey) return { bg: '#ffffff', fg: '#000000' };
    const k = c.width / w;
    const x = Math.floor(r.x * k), y = Math.floor(r.y * k);
    const W = Math.max(1, Math.ceil(r.w * k)), H = Math.max(1, Math.ceil(r.h * k));
    let data;
    try { data = c.getContext('2d', { willReadFrequently: true }).getImageData(x, y, W, H).data; } catch { return { bg: '#ffffff', fg: '#000000' }; }
    const hist = new Map();
    for (let i = 0; i < data.length; i += 4) {
      const q = ((data[i] >> 3) << 10) | ((data[i + 1] >> 3) << 5) | (data[i + 2] >> 3);
      hist.set(q, (hist.get(q) || 0) + 1);
    }
    let bq = 0, bn = -1;
    for (const [q, n] of hist) if (n > bn) { bn = n; bq = q; }
    const br = ((bq >> 10) & 31) * 8 + 4, bgg = ((bq >> 5) & 31) * 8 + 4, bb = (bq & 31) * 8 + 4;
    let far = -1, fr = 0, fg = 0, fb = 0;
    for (let i = 0; i < data.length; i += 4) {
      const d = Math.abs(data[i] - br) + Math.abs(data[i + 1] - bgg) + Math.abs(data[i + 2] - bb);
      if (d > far) { far = d; fr = data[i]; fg = data[i + 1]; fb = data[i + 2]; }
    }
    const bgHex = bn / (data.length / 4) > 0.35 ? hexOf(Math.min(255, br), Math.min(255, bgg), Math.min(255, bb)) : '#ffffff';
    const bgFix = /^#f[89a-f]f[89a-f]f[89a-f]$/i.test(bgHex) ? '#ffffff' : bgHex;
    return { bg: bgFix, fg: far > 120 ? hexOf(fr, fg, fb) : '#000000' };
  }

  replaceRun(p, r, clientX, clientY) {
    const snap = this.snapshot();
    const { bg, fg } = this.sampleColors(p, r);
    const [asc, desc] = FONT_METRICS[r.font];
    const size = Math.round(r.size * 2) / 2;
    const cover = { id: nid(), t: 'whiteout', x: r.x - 0.8, y: r.y - 0.4, w: r.w + 1.6, h: r.h + 0.8, color: bg, run: r.key };
    const top = r.base - (1.2 * size - (asc + desc) * size) / 2 - asc * size;
    const t = { id: nid(), t: 'text', x: r.x, y: top, text: r.str, font: r.font, size, color: fg, bold: r.bold, italic: r.italic, cover: cover.id, orig: r.str };
    p.objs.push(cover, t);
    this.startEdit(p, t, { snap, isNew: true, clientX, clientY });
  }

  // ---------- Редактирование текста ----------
  startEdit(p, o, { snap, isNew = false, clientX, clientY } = {}) {
    this.endEdit();
    const base = snap || this.snapshot();
    const wasPristine = !!o.pristine;
    if (wasPristine) this.unPristine(p, [o]);
    this.dropPristine({ ids: [o.id] });
    this.sel = { pid: p.id, ids: [o.id] };
    const rec = this.els.get(p.id);
    const t = el('div', 'obj-text editing');
    t.textContent = o.text;
    t.contentEditable = 'plaintext-only';
    t.spellcheck = false;
    this.editing = { pid: p.id, oid: o.id, el: t, snap: base, isNew, wasPristine, before: o.text };
    rec.text.append(t);
    this.renderObjs(p);
    t.focus();
    const range = document.createRange();
    let placed = false;
    if (clientX != null && document.caretRangeFromPoint) {
      const r = document.caretRangeFromPoint(clientX, clientY);
      if (r && t.contains(r.startContainer)) { range.setStart(r.startContainer, r.startOffset); placed = true; }
    }
    if (!placed) { range.selectNodeContents(t); range.collapse(false); }
    range.collapse(true);
    const s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
    t.addEventListener('input', () => {
      o.text = t.innerText.replace(/\r/g, '');
      this.drawSel();
    });
    t.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); this.endEdit(); return; }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'b' || e.key === 'и' || e.key === 'i' || e.key === 'ш')) {
        e.preventDefault();
        this.setProp(/[bи]/.test(e.key) ? 'bold' : 'italic', null);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && /^[sоo]$/i.test(e.key)) return; // сохранение/открытие пусть дойдут
      e.stopPropagation();
    });
    t.addEventListener('blur', () => setTimeout(() => {
      if (this.editing && this.editing.el === t && document.activeElement !== t) this.endEdit();
    }, 0));
    this.drawSel();
    this.updateCtx();
  }

  endEdit() {
    const ed = this.editing;
    if (!ed) return;
    this.editing = null;
    const p = this.page(ed.pid);
    const o = p && p.objs.find((x) => x.id === ed.oid);
    ed.el.remove();
    if (!o) return;
    o.text = o.text.replace(/\n+$/, '');
    const cover = o.cover ? p.objs.find((x) => x.id === o.cover) : null;
    const unselect = () => { if (this.sel && this.sel.ids.includes(o.id)) this.sel = null; };
    let changed = true;
    if (!o.text.trim()) {
      if (ed.isNew) { this.removeObj(p, o); changed = false; }
      else p.objs = p.objs.filter((x) => x !== o); // текст стёрли — замазка остаётся
      unselect();
    } else if (ed.isNew && o.orig != null && o.text === o.orig && !ed.propsChanged) {
      this.removeObj(p, o); // ничего не поменяли — возвращаем оригинал как был
      unselect();
      changed = false;
    } else if (!ed.isNew && o.text === ed.before && !ed.propsChanged) {
      changed = false;
      if (ed.wasPristine) { o.pristine = true; if (cover) cover.pristine = true; }
    }
    if (changed) this.checkpoint(ed.snap);
    this.renderObjs(p);
    this.drawSel();
    this.updateCtx();
  }

  // ---------- Мышь и перо ----------
  toPt(e, pageEl) {
    const r = pageEl.getBoundingClientRect();
    return { x: (e.clientX - r.left) / this.S, y: (e.clientY - r.top) / this.S };
  }

  bindViewer() {
    const v = this.viewer;
    v.addEventListener('pointerdown', (e) => this.onDown(e));
    // Не даём просмотрщику забирать фокус у редактируемого текста.
    v.addEventListener('mousedown', (e) => { if (!e.target.closest('.obj-text.editing')) e.preventDefault(); });
    v.addEventListener('pointermove', (e) => this.onHover(e));
    v.addEventListener('dblclick', (e) => {
      const pe = e.target.closest('.page');
      if (!pe) return;
      const p = this.page(pe.dataset.pid);
      const { x, y } = this.toPt(e, pe);
      const o = this.hitTest(p, x, y);
      if (o && o.t === 'text') this.startEdit(p, o, { clientX: e.clientX, clientY: e.clientY });
    });
    v.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * Math.exp(-e.deltaY * 0.0022), e);
    }, { passive: false });
    v.addEventListener('scroll', () => { this.updateHud(true); this.checkVisible(); }, { passive: true });
  }

  async onHover(e) {
    if (this.drag) return;
    const pe = e.target.closest('.page');
    const clear = () => {
      this.column.querySelectorAll('.hover-run').forEach((n) => n.remove());
      this.column.querySelectorAll('.over-obj,.over-run').forEach((n) => n.classList.remove('over-obj', 'over-run'));
    };
    const tool = this.tool;
    if (!pe || this.editing || (tool !== 'select' && tool !== 'text')) { clear(); return; }
    const p = this.page(pe.dataset.pid);
    const { x, y } = this.toPt(e, pe);
    const o = this.hitTest(p, x, y);
    const token = (this.hoverToken = (this.hoverToken || 0) + 1);
    let n = null;
    if (!o || (tool === 'text' && o.t !== 'text')) n = await this.nativeAt(p, x, y, tool === 'text' ? ['text'] : ['text', 'image']);
    if (token !== this.hoverToken) return;
    clear();
    if (o && (tool === 'select' || o.t === 'text')) { if (tool === 'select') pe.classList.add('over-obj'); return; }
    if (!n) return;
    const S = this.S;
    const h = el('div', 'hover-run' + (tool === 'select' ? ' soft' : ''));
    Object.assign(h.style, { left: (n.x - 1) * S + 'px', top: n.y * S + 'px', width: (n.w + 2) * S + 'px', height: n.h * S + 'px' });
    this.els.get(p.id).marks.append(h);
    pe.classList.add(tool === 'select' ? 'over-obj' : 'over-run');
  }

  async onDown(e) {
    if (e.button !== 0) return;
    for (const n of this.column.querySelectorAll('.hover-run')) n.remove();
    const handle = e.target.closest('.handle');
    const pe = e.target.closest('.page');
    if (this.editing) {
      if (e.target.closest('.obj-text.editing')) return;
      this.endEdit();
      if (this.tool === 'text') return;
    }
    if (!pe) { if (!e.target.closest('.dock,.ctx')) this.select(null); return; }
    const p = this.page(pe.dataset.pid);
    const pt = this.toPt(e, pe);
    this.cur = p.id;

    if (handle) return this.beginResize(e, pe, p, handle.dataset.h);

    const tool = this.tool;
    const add = e.shiftKey || e.ctrlKey;
    if (tool === 'text') {
      const o = this.hitTest(p, pt.x, pt.y);
      if (o && o.t === 'text') {
        this.select(p.id, [o.id]);
        return this.beginMove(e, pe, p, true);
      }
      const r = await this.runAt(p, pt.x, pt.y);
      if (r) return this.replaceRun(p, r, e.clientX, e.clientY);
      const st = Settings.get('text');
      const t = { id: nid(), t: 'text', x: pt.x, y: pt.y - st.size * 0.6, text: '', font: st.font, size: st.size, color: st.color, bold: st.bold, italic: st.italic };
      const snap = this.snapshot();
      p.objs.push(t);
      return this.startEdit(p, t, { snap, isNew: true });
    }
    if (tool === 'select') {
      const o = this.hitTest(p, pt.x, pt.y);
      if (o) {
        if (add) return this.toggleSel(p, [o]);
        if (!this.isSelected(p, o)) this.select(p.id, [o.id]);
        return this.beginMove(e, pe, p);
      }
      const n = await this.nativeAt(p, pt.x, pt.y);
      if (n) {
        const objs = this.convertNative(p, [n]);
        if (objs.length) {
          if (add) return this.toggleSel(p, objs);
          this.select(p.id, objs.map((x) => x.id));
          if (PTR_DOWN) this.beginMove(e, pe, p);
          return;
        }
      }
      if (!add) this.select(null);
      if (PTR_DOWN) this.beginMarquee(e, pe, p, add);
      return;
    }
    if (tool === 'pen' || tool === 'marker') return this.beginStroke(e, pe, p, pt, tool);
    if (tool === 'shape' || tool === 'whiteout') return this.beginShape(e, pe, p, pt, tool);
  }

  track(e, move, up) {
    const target = e.target;
    try { target.setPointerCapture(e.pointerId); } catch {}
    const onMove = (ev) => move(ev);
    const onUp = (ev) => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      this.drag = null;
      up && up(ev);
    };
    this.drag = true;
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  }

  // Двигаем всё выделенное разом.
  beginMove(e, pe, p, editOnClick) {
    const s = this.selectedObjs();
    if (!s) return;
    const start = this.toPt(e, pe);
    const snap = this.snapshot();
    let orig = null;
    let moved = false;
    this.track(e, (ev) => {
      const q = this.toPt(ev, pe);
      const dx = q.x - start.x, dy = q.y - start.y;
      if (!moved && Math.hypot(dx, dy) * this.S < 3) return;
      if (!moved) {
        moved = true;
        this.checkpoint(snap);
        this.unPristine(p, s.objs);
        orig = s.objs.map((o) => structuredClone(o));
      }
      s.objs.forEach((o, i) => { Object.assign(o, structuredClone(orig[i])); translate(o, dx, dy); });
      this.refresh(p);
    }, (ev) => {
      if (!moved && editOnClick && s.objs.length === 1) this.startEdit(p, s.objs[0], { clientX: ev.clientX, clientY: ev.clientY });
    });
  }

  beginResize(e, pe, p, h) {
    const s = this.selectedObjs();
    if (!s) return;
    const snap = this.snapshot();
    const one = s.objs.length === 1 ? s.objs[0] : null;
    const keepRatio = s.objs.some((o) => o.t === 'image' || o.t === 'text');
    let moved = false, orig = null, b0 = null;
    this.track(e, (ev) => {
      const q = this.toPt(ev, pe);
      if (!moved) {
        moved = true;
        this.checkpoint(snap);
        this.unPristine(p, s.objs);
        orig = s.objs.map((o) => structuredClone(o));
        b0 = this.unionBox(s.objs);
      }
      if (h === 'p1' || h === 'p2') {
        if (h === 'p1') { one.x1 = q.x; one.y1 = q.y; } else { one.x2 = q.x; one.y2 = q.y; }
        this.refresh(p);
        return;
      }
      const ax = h.includes('w') ? b0.x + b0.w : b0.x;
      const ay = h.includes('n') ? b0.y + b0.h : b0.y;
      let w = Math.max(4, Math.abs(q.x - ax)), hh = Math.max(4, Math.abs(q.y - ay));
      if (keepRatio || ev.shiftKey) {
        const k = Math.max(w / (b0.w || 1), hh / (b0.h || 1));
        w = b0.w * k; hh = b0.h * k;
      }
      const nb = { x: h.includes('w') ? ax - w : ax, y: h.includes('n') ? ay - hh : ay, w, h: hh };
      s.objs.forEach((o, i) => { Object.assign(o, structuredClone(orig[i])); fitTo(o, orig[i], b0, nb); });
      if (one && one.t === 'text') {
        const nb2 = bbox(one);
        one.x = h.includes('w') ? ax - nb2.w : ax;
        one.y = h.includes('n') ? ay - nb2.h : ay;
      }
      this.refresh(p);
    });
  }

  // Рамка выделения: всё, что задела рамка, — в группу.
  beginMarquee(e, pe, p, add) {
    const start = this.toPt(e, pe);
    const rec = this.els.get(p.id);
    const box = el('div', 'marquee');
    rec.marks.append(box);
    const prev = add && this.sel && this.sel.pid === p.id ? [...this.sel.ids] : [];
    let b = null;
    this.track(e, (ev) => {
      const q = this.toPt(ev, pe);
      const { w, h } = this.dims(p);
      const qx = clamp(q.x, 0, w), qy = clamp(q.y, 0, h);
      b = { x: Math.min(start.x, qx), y: Math.min(start.y, qy), w: Math.abs(qx - start.x), h: Math.abs(qy - start.y) };
      const S = this.S;
      Object.assign(box.style, { left: b.x * S + 'px', top: b.y * S + 'px', width: b.w * S + 'px', height: b.h * S + 'px' });
    }, async () => {
      box.remove();
      if (!b || (b.w * this.S < 4 && b.h * this.S < 4)) return;
      const touches = (o) => {
        const ob = bbox(o);
        return ob.x < b.x + b.w && ob.x + ob.w > b.x && ob.y < b.y + b.h && ob.y + ob.h > b.y;
      };
      const objs = p.objs.filter((o) => !o.pristine && !(o.t === 'whiteout' && o.run) && touches(o));
      const natives = this.convertNative(p, await this.nativesIn(p, b));
      this.select(p.id, [...prev, ...objs.map((o) => o.id), ...natives.map((o) => o.id)]);
    });
  }

  beginStroke(e, pe, p, pt, tool) {
    const st = Settings.get(tool);
    const o = { id: nid(), t: tool, pts: [pt.x, pt.y], color: st.color, width: st.width };
    this.checkpoint();
    p.objs.push(o);
    this.renderObjs(p);
    const minD = 0.6 / this.zoom;
    this.track(e, (ev) => {
      const evs = ev.getCoalescedEvents ? ev.getCoalescedEvents() : [ev];
      for (const ce of evs.length ? evs : [ev]) {
        const q = this.toPt(ce, pe);
        if (ev.shiftKey) { o.pts = [o.pts[0], o.pts[1], q.x, q.y]; continue; }
        const n = o.pts.length;
        if (Math.hypot(q.x - o.pts[n - 2], q.y - o.pts[n - 1]) >= minD) o.pts.push(q.x, q.y);
      }
      this.renderObjs(p);
    }, () => {
      o.pts = o.pts.map((v) => Math.round(v * 100) / 100);
      this.renderObjs(p);
    });
  }

  beginShape(e, pe, p, pt, tool) {
    const st = Settings.get(tool);
    const kind = tool === 'whiteout' ? 'whiteout' : st.kind;
    const isLine = kind === 'line' || kind === 'arrow';
    const o = isLine
      ? { id: nid(), t: kind, x1: pt.x, y1: pt.y, x2: pt.x, y2: pt.y, color: st.color, width: st.width }
      : { id: nid(), t: kind, x: pt.x, y: pt.y, w: 0, h: 0, color: st.color, ...(kind === 'whiteout' ? {} : { width: st.width }) };
    const snap = this.snapshot();
    p.objs.push(o);
    let moved = false;
    this.track(e, (ev) => {
      const q = this.toPt(ev, pe);
      moved = true;
      if (isLine) {
        let dx = q.x - pt.x, dy = q.y - pt.y;
        if (ev.shiftKey) {
          const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
          const l = Math.hypot(dx, dy);
          dx = Math.cos(a) * l; dy = Math.sin(a) * l;
        }
        o.x2 = pt.x + dx; o.y2 = pt.y + dy;
      } else {
        let w = q.x - pt.x, h = q.y - pt.y;
        if (ev.shiftKey) { const m = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * m; h = Math.sign(h || 1) * m; }
        o.x = Math.min(pt.x, pt.x + w); o.y = Math.min(pt.y, pt.y + h); o.w = Math.abs(w); o.h = Math.abs(h);
      }
      this.renderObjs(p);
    }, () => {
      const b = bbox(o);
      if (!moved || Math.max(b.w, b.h) < 3) {
        if (kind === 'whiteout') { this.removeObj(p, o); this.renderObjs(p); return; }
        // Просто клик — ставим фигуру стандартного размера.
        if (isLine) { o.x2 = pt.x + 80; o.y2 = pt.y; } else { o.w = 90; o.h = kind === 'ellipse' ? 90 : 60; }
      }
      this.checkpoint(snap);
      this.renderObjs(p);
      this.setTool('select');
      this.select(p.id, [o.id]);
    });
  }

  // ---------- Инструменты и свойства ----------
  buildDock() {
    const dock = el('div', 'dock');
    const tools = [
      ['pages', 'Страницы · Ctrl+B'],
      null,
      ['select', 'Выбор · V'],
      ['text', 'Текст · T — щёлкните по тексту, чтобы исправить, или по пустому месту, чтобы добавить'],
      ['pen', 'Перо · P'],
      ['marker', 'Маркер · H'],
      ['shape', 'Фигуры · R'],
      ['whiteout', 'Замазать · E'],
      ['image', 'Картинка · I'],
      ['sign', 'Подпись · S'],
      null,
      ['hide', 'Скрыть панель · Ctrl+.'],
    ];
    this.toolBtns = {};
    for (const t of tools) {
      if (!t) { dock.append(el('div', 'sep')); continue; }
      const b = el('button', 'tool', ICONS[t[0]] + '<span class="dot"></span>');
      b.dataset.tip = t[1];
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.addEventListener('click', () => this.onToolClick(t[0]));
      dock.append(b);
      this.toolBtns[t[0]] = b;
    }
    this.dock = dock;
    this.ctx = el('div', 'ctx');
    this.ctx.hidden = true;
    this.ctx.addEventListener('pointerdown', (e) => e.preventDefault());
    const handle = el('button', 'dock-handle');
    handle.dataset.tip = 'Показать панель · Ctrl+.';
    handle.addEventListener('click', () => window.setDockHidden && window.setDockHidden(false));
    this.root.append(dock, this.ctx, handle);
    this.dockMover = makeDockMovable(this.root, dock, this.ctx, 'pdf');
    this.setTool('select');
  }

  onToolClick(t) {
    if (t === 'pages') return this.toggleSide();
    if (t === 'hide') return window.setDockHidden && window.setDockHidden(true);
    if (t === 'image') return this.pickImage();
    if (t === 'sign') return this.openSignature();
    this.setTool(this.tool === t && t !== 'select' ? 'select' : t);
  }

  setTool(t) {
    this.endEdit();
    this.tool = t;
    this.viewer.className = 'viewer t-' + t;
    for (const [k, b] of Object.entries(this.toolBtns)) {
      b.classList.toggle('on', k === t || (k === 'pages' && this.root.classList.contains('side-open')));
      const st = Settings.get(k);
      if (st && st.color) b.style.setProperty('--c', st.color);
    }
    if (t !== 'select' && t !== 'text') this.select(null);
    else this.updateCtx();
    for (const n of this.column.querySelectorAll('.hover-run')) n.remove();
  }

  // Что показывать над пеналом: свойства выделенного или текущего инструмента.
  updateCtx() {
    const c = this.ctx;
    c.textContent = '';
    const s = this.selectedObjs();
    const shapeT = ['rect', 'ellipse', 'line', 'arrow'];
    const kindOf = (t) => (shapeT.includes(t) ? 'shape' : t);
    let kind = null, props = null, target = null;
    if (s) {
      target = 'obj';
      if (new Set(s.objs.map((o) => kindOf(o.t))).size === 1) { kind = kindOf(s.objs[0].t); props = s.objs[0]; }
    } else if (['pen', 'marker', 'shape', 'text', 'whiteout'].includes(this.tool)) {
      kind = this.tool; props = Settings.get(this.tool); target = 'tool';
    }
    if (!target) { c.hidden = true; return; }
    const add = (n) => c.append(n);
    const sep = () => add(el('div', 'sep'));
    const btn = (icon, tip, on, fn, cls = '') => {
      const b = el('button', 'cbtn ' + cls + (on ? ' on' : ''), icon);
      b.dataset.tip = tip;
      b.addEventListener('click', fn);
      add(b);
      return b;
    };
    const colors = (list) => {
      for (const col of list) {
        const b = el('button', 'sw' + (props.color === col ? ' on' : ''), '<i></i>');
        b.style.setProperty('--c', col);
        b.addEventListener('click', () => this.setProp('color', col));
        add(b);
      }
    };
    const size = (key, step, min, max, tip) => {
      const w = el('div', 'size');
      w.dataset.tip = tip;
      const minus = el('button', '', '−');
      const out = el('output', '', String(props[key]));
      const plus = el('button', '', '+');
      minus.addEventListener('click', () => this.setProp(key, clamp(+(props[key] - step).toFixed(1), min, max)));
      plus.addEventListener('click', () => this.setProp(key, clamp(+(props[key] + step).toFixed(1), min, max)));
      w.append(minus, out, plus);
      add(w);
    };

    if (kind === 'text') {
      const f = el('button', 'fontname', FONT_NAME[props.font]);
      f.dataset.tip = 'Шрифт';
      f.style.fontFamily = FONT_CSS[props.font];
      f.addEventListener('click', () => {
        const order = ['arial', 'times', 'courier'];
        this.setProp('font', order[(order.indexOf(props.font) + 1) % order.length]);
      });
      add(f);
      size('size', props.size >= 24 ? 2 : 1, 4, 200, 'Размер');
      btn(ICONS.bold, 'Жирный · Ctrl+B', props.bold, () => this.setProp('bold', null));
      btn(ICONS.italic, 'Курсив · Ctrl+I', props.italic, () => this.setProp('italic', null));
      sep();
      colors(INK);
    } else if (kind === 'shape') {
      if (target === 'tool') {
        for (const k of shapeT) {
          btn(ICONS[k], { rect: 'Прямоугольник', ellipse: 'Овал', line: 'Линия', arrow: 'Стрелка' }[k], props.kind === k, () => this.setProp('kind', k));
        }
        sep();
      }
      colors(INK);
      sep();
      size('width', 0.5, 0.5, 20, 'Толщина');
    } else if (kind === 'pen') {
      colors(INK); sep(); size('width', 0.5, 0.5, 20, 'Толщина');
    } else if (kind === 'marker') {
      colors(HL); sep(); size('width', 2, 4, 40, 'Толщина');
    } else if (kind === 'whiteout') {
      colors(COVER);
    }
    if (target === 'obj') {
      if (c.childNodes.length) sep();
      btn(ICONS.copy, 'Дублировать · Ctrl+D', false, () => this.duplicateSel());
      btn(ICONS.trash, 'Удалить · Delete', false, () => this.deleteSel(), 'danger');
    }
    c.hidden = !c.childNodes.length;
    if (this.dockMover) this.dockMover.placeCtx();
  }

  setProp(key, val) {
    const s = this.selectedObjs();
    const sk = (t) => (['rect', 'ellipse', 'line', 'arrow'].includes(t) ? 'shape' : t);
    if (s) {
      if (key === 'kind') return;
      if (val === null) val = !s.objs[0][key];
      const snap = this.snapshot();
      for (const o of s.objs) if (key in o) o[key] = val;
      if (this.editing) this.editing.propsChanged = true;
      else { this.checkpoint(snap); this.unPristine(s.p, s.objs); }
      if (Settings.get(sk(s.objs[0].t))) Settings.set(sk(s.objs[0].t), { [key]: val });
      this.refresh(s.p);
      if (this.editing) this.editing.el.focus();
    } else {
      const st = Settings.get(this.tool);
      if (!st) return;
      if (val === null) val = !st[key];
      Settings.set(this.tool, { [key]: val });
    }
    const b = this.toolBtns[s ? sk(s.objs[0].t) : this.tool];
    if (b && key === 'color') b.style.setProperty('--c', val);
    this.updateCtx();
  }

  deleteSel() {
    if (this.editing) this.endEdit();
    const s = this.selectedObjs();
    if (!s) return;
    this.checkpoint();
    const ids = new Set(s.objs.map((o) => o.id));
    s.p.objs = s.p.objs.filter((o) => !ids.has(o.id));
    // Замазка под удалённым содержимым PDF остаётся — иначе оно проступит обратно.
    for (const o of s.objs) {
      const c = o.cover && s.p.objs.find((x) => x.id === o.cover);
      if (c) delete c.pristine;
    }
    this.sel = null;
    this.refresh(s.p);
    this.updateCtx();
  }

  cleanCopy(o) {
    const c = structuredClone(o);
    c.id = nid();
    delete c.cover; delete c.run; delete c.orig; delete c.pristine; delete c.nbox;
    return c;
  }

  duplicateSel() {
    this.endEdit();
    const s = this.selectedObjs();
    if (!s) return;
    this.checkpoint();
    const copies = s.objs.map((o) => { const c = this.cleanCopy(o); translate(c, 12, 12); return c; });
    s.p.objs.push(...copies);
    this.refresh(s.p);
    this.select(s.p.id, copies.map((c) => c.id));
  }

  nudge(dx, dy) {
    const s = this.selectedObjs();
    if (!s) return;
    this.checkpoint();
    this.unPristine(s.p, s.objs);
    for (const o of s.objs) translate(o, dx, dy);
    this.refresh(s.p);
  }

  // ---------- Картинки и подпись ----------
  currentPage() {
    if (this.cur && this.page(this.cur)) return this.page(this.cur);
    return this.pages[0];
  }

  async insertImage(bytes, name, widthShare = 0.5) {
    const p = this.currentPage();
    if (!p) return;
    const im = await normalizeImage(bytes, name);
    const id = nid();
    this.imgs.set(id, im);
    this.placeImage(p, id, im, widthShare);
  }

  placeImage(p, imgId, im, widthShare) {
    const { w, h } = this.dims(p);
    const k = Math.min((w * widthShare) / im.w, (h * 0.6) / im.h);
    const ow = im.w * k, oh = im.h * k;
    // Ставим в центр видимой части страницы.
    const rec = this.els.get(p.id);
    const vr = this.viewer.getBoundingClientRect(), pr = rec.el.getBoundingClientRect();
    const midY = clamp(((vr.top + vr.bottom) / 2 - pr.top) / this.S, oh / 2, h - oh / 2);
    this.checkpoint();
    const o = { id: nid(), t: 'image', img: imgId, x: (w - ow) / 2, y: midY - oh / 2, w: ow, h: oh };
    p.objs.push(o);
    this.setTool('select');
    this.renderObjs(p);
    this.select(p.id, [o.id]);
  }

  async pickImage() {
    const paths = await window.folio.openDialog({ filters: [{ name: T('Картинки'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] }] });
    for (const path of paths.slice(0, 1)) {
      const bytes = await window.folio.readFile(path);
      await this.insertImage(bytes, path);
    }
  }

  openSignature() {
    const saved = (() => { try { return JSON.parse(localStorage.getItem('folio.signs') || '[]'); } catch { return []; } })();
    const bg = el('div', 'modal-bg');
    const m = el('div', 'modal');
    bg.append(m);
    const list = el('div', 'sig-saved');
    const close = () => { bg.remove(); document.removeEventListener('keydown', onKey, true); };
    const use = async (url) => {
      close();
      const bmp = await createImageBitmap(await (await fetch(url)).blob());
      const id = nid();
      this.imgs.set(id, { url, w: bmp.width, h: bmp.height });
      this.placeImage(this.currentPage(), id, { w: bmp.width, h: bmp.height }, 0.28);
    };
    const drawList = () => {
      list.textContent = '';
      saved.forEach((url, i) => {
        const b = el('button');
        b.dataset.tip = 'Вставить эту подпись';
        const img = el('img');
        img.src = url;
        const x = el('span', 'x', '✕');
        x.addEventListener('click', (e) => {
          e.stopPropagation();
          saved.splice(i, 1);
          try { localStorage.setItem('folio.signs', JSON.stringify(saved)); } catch {}
          drawList();
        });
        b.append(img, x);
        b.addEventListener('click', () => use(url));
        list.append(b);
      });
      list.hidden = !saved.length;
    };
    drawList();

    const pad = el('div', 'sig-pad');
    const line = el('div', 'sig-line');
    const cv = document.createElement('canvas');
    const dpr = window.devicePixelRatio || 1;
    cv.width = 520 * dpr; cv.height = 200 * dpr;
    cv.style.width = '520px'; cv.style.height = '200px';
    pad.append(line, cv);
    const g = cv.getContext('2d');
    g.scale(dpr, dpr);
    g.lineCap = 'round'; g.lineJoin = 'round';
    let color = '#1d2125', inked = false;
    let last = null;
    cv.addEventListener('pointerdown', (e) => {
      cv.setPointerCapture(e.pointerId);
      last = { x: e.offsetX, y: e.offsetY, w: 2.6 };
    });
    cv.addEventListener('pointermove', (e) => {
      if (!last) return;
      for (const ce of e.getCoalescedEvents ? e.getCoalescedEvents() : [e]) {
        const x = ce.offsetX, y = ce.offsetY;
        const d = Math.hypot(x - last.x, y - last.y);
        const w = clamp(3.2 - d * 0.12, 1.3, 3.2) * 0.4 + last.w * 0.6;
        g.strokeStyle = color; g.lineWidth = w;
        g.beginPath(); g.moveTo(last.x, last.y); g.lineTo(x, y); g.stroke();
        last = { x, y, w };
        inked = true;
      }
    });
    cv.addEventListener('pointerup', () => (last = null));

    const row = el('div', 'modal-row');
    for (const col of ['#1d2125', '#1f3fae']) {
      const sw = el('button', 'sw' + (col === color ? ' on' : ''), '<i></i>');
      sw.style.setProperty('--c', col);
      sw.addEventListener('click', () => {
        color = col;
        row.querySelectorAll('.sw').forEach((n) => n.classList.toggle('on', n === sw));
      });
      row.append(sw);
    }
    row.append(el('div', 'grow'));
    const clear = el('button', 'btn', 'Очистить');
    clear.addEventListener('click', () => { g.clearRect(0, 0, 520, 200); inked = false; });
    const cancel = el('button', 'btn', 'Отмена');
    cancel.addEventListener('click', close);
    const ok = el('button', 'btn primary', 'Вставить');
    ok.addEventListener('click', () => {
      if (!inked) return close();
      const url = trimCanvas(cv);
      saved.unshift(url);
      saved.splice(6);
      try { localStorage.setItem('folio.signs', JSON.stringify(saved)); } catch {}
      use(url);
    });
    row.append(clear, cancel, ok);
    m.append(list, pad, row);
    const onKey = (e) => {
      if (e.key === 'Escape') { e.stopPropagation(); close(); }
      if (e.key === 'Enter') { e.stopPropagation(); ok.click(); }
    };
    document.addEventListener('keydown', onKey, true);
    bg.addEventListener('pointerdown', (e) => { if (e.target === bg) close(); });
    document.body.append(bg);
  }

  // ---------- Масштаб ----------
  fitWidth() {
    const maxW = Math.max(...this.pages.map((p) => this.dims(p).w), 100);
    const avail = this.viewer.clientWidth - 84;
    this.setZoom(clamp(avail / (maxW * PT), 0.25, 1.2), null, true);
    this.fitMode = true;
  }

  setZoom(z, anchor, quiet) {
    z = clamp(z, 0.25, 5);
    this.fitMode = false;
    if (Math.abs(z - this.zoom) < 1e-4) return;
    // Запоминаем точку под курсором (или центр), чтобы после масштаба она осталась на месте.
    const vr = this.viewer.getBoundingClientRect();
    const ax = anchor ? anchor.clientX : vr.left + vr.width / 2;
    const ay = anchor ? anchor.clientY : vr.top + vr.height / 2;
    let ref = null;
    for (const [pid, rec] of this.els) {
      const r = rec.el.getBoundingClientRect();
      if (ay >= r.top - 9 && ay <= r.bottom + 9) {
        ref = { pid, fx: (ax - r.left) / r.width, fy: (ay - r.top) / r.height };
        break;
      }
    }
    this.zoom = z;
    this.layout();
    if (ref) {
      const r = this.els.get(ref.pid).el.getBoundingClientRect();
      this.viewer.scrollLeft += r.left + ref.fx * r.width - ax;
      this.viewer.scrollTop += r.top + ref.fy * r.height - ay;
    }
    if (!quiet) this.flashHud(Math.round(z * 100) + '%');
  }

  zoomStep(dir) {
    const i = ZOOMS.findIndex((v) => v > this.zoom + 1e-3);
    const next = dir > 0 ? ZOOMS[i < 0 ? ZOOMS.length - 1 : i] : [...ZOOMS].reverse().find((v) => v < this.zoom - 1e-3) || ZOOMS[0];
    this.setZoom(next);
  }

  flashHud(text) {
    this.hud.textContent = text;
    this.hud.classList.add('show');
    clearTimeout(this.hudTimer);
    this.hudTimer = setTimeout(() => this.hud.classList.remove('show'), 900);
  }

  updateHud(show) {
    const vr = this.viewer.getBoundingClientRect();
    const mid = vr.top + vr.height * 0.4;
    let best = null, bd = Infinity;
    for (const p of this.pages) {
      const r = this.els.get(p.id).el.getBoundingClientRect();
      const d = r.top <= mid && r.bottom >= mid ? 0 : Math.min(Math.abs(r.top - mid), Math.abs(r.bottom - mid));
      if (d < bd) { bd = d; best = p; }
      if (d === 0) break;
    }
    if (!best) return;
    if (best.id !== this.cur) {
      this.cur = best.id;
      for (const [pid, t] of this.thumbEls) t.classList.toggle('cur', pid === best.id);
    }
    if (show && this.pages.length > 1) this.flashHud(`${this.pages.indexOf(best) + 1} / ${this.pages.length}`);
  }

  scrollToPage(pid) {
    const rec = this.els.get(+pid);
    if (!rec) return;
    this.viewer.scrollTo({ top: rec.el.offsetTop - 20, behavior: 'smooth' });
    this.cur = +pid;
  }

  // ---------- Панель страниц ----------
  toggleSide(force) {
    const open = force != null ? force : !this.root.classList.contains('side-open');
    this.root.classList.toggle('side-open', open);
    this.toolBtns.pages.classList.toggle('on', open);
    if (open) {
      const t = this.thumbEls.get(this.cur);
      if (t) t.scrollIntoView({ block: 'center' });
      setTimeout(() => this.checkVisible(), 200);
    }
  }

  rebuildThumbs() {
    const keepKeys = new Set();
    const frag = document.createDocumentFragment();
    const gap = (i) => {
      const g = el('div', 'gap');
      g.dataset.at = i;
      const b = el('button', '', ICONS.plus);
      b.dataset.tip = 'Вставить пустую страницу';
      g.append(b);
      return g;
    };
    frag.append(gap(0));
    this.pages.forEach((p, i) => {
      let t = this.thumbEls.get(p.id);
      const { w, h } = this.dims(p);
      const k = 134 / Math.max(w, h * 0.78);
      if (!t) {
        t = el('div', 'thumb');
        t.dataset.pid = p.id;
        t.innerHTML = `<div class="t-paper"></div><div class="t-num"></div><div class="t-acts"><button data-act="rotate">${ICONS.rotate}</button><button data-act="delete">${ICONS.trash}</button></div>`;
        t.querySelector('[data-act=rotate]').dataset.tip = 'Повернуть';
        t.querySelector('[data-act=delete]').dataset.tip = 'Удалить страницу';
        this.thumbEls.set(p.id, t);
      }
      const paper = t.firstChild;
      paper.style.width = Math.round(w * k) + 'px';
      paper.style.height = Math.round(h * k) + 'px';
      t.querySelector('.t-num').textContent = i + 1;
      t.classList.toggle('picked', this.picked.has(p.id));
      t.classList.toggle('cur', p.id === this.cur);
      keepKeys.add(p.id);
      frag.append(t, gap(i + 1));
      if (t.dataset.key && t.dataset.key !== this.thumbKey(p)) this.renderThumb(p.id);
    });
    for (const [pid, t] of this.thumbEls) if (!keepKeys.has(pid)) { t.remove(); this.thumbEls.delete(pid); }
    const addBtn = el('button', 'add-pages', ICONS.plus);
    addBtn.dataset.tip = 'Добавить страницы из файла';
    addBtn.addEventListener('click', () => this.addPagesDialog());
    this.thumbs.textContent = '';
    this.thumbs.append(frag, addBtn);
    this.checkVisible();
  }

  thumbKey(p) { return `${this.total(p)}|${p.objs.filter((o) => !o.pristine).length}`; }

  async renderThumb(pid) {
    const p = this.page(pid);
    const t = this.thumbEls.get(+pid);
    if (!p || !t) return;
    const key = this.thumbKey(p);
    if (t.dataset.key === key) return;
    t.dataset.key = key;
    const paper = t.firstChild;
    const { w, h } = this.dims(p);
    const k = (parseFloat(paper.style.width) / w) * 2;
    const c = document.createElement('canvas');
    c.width = Math.ceil(w * k); c.height = Math.ceil(h * k);
    const g = c.getContext('2d');
    g.fillStyle = '#fff';
    g.fillRect(0, 0, c.width, c.height);
    if (p.src != null) {
      const pg = await this.pdfPage(this.source(p.src), p.index);
      try {
        await pg.render({ canvasContext: g, viewport: pg.getViewport({ scale: k, rotation: this.total(p) }) }).promise;
      } catch {}
    }
    // Картинки и замазку рисуем грубо, чтобы миниатюра была узнаваемой.
    for (const o of p.objs) {
      if (o.pristine) continue;
      if (o.t === 'image') {
        const im = new Image();
        im.src = this.imgs.get(o.img).url;
        try { await im.decode(); g.drawImage(im, o.x * k, o.y * k, o.w * k, o.h * k); } catch {}
      } else if (o.t === 'whiteout') {
        g.fillStyle = o.color; g.fillRect(o.x * k, o.y * k, o.w * k, o.h * k);
      }
    }
    paper.querySelector('canvas')?.remove();
    paper.append(c);
  }

  // Панель страниц слева или справа (FLIP-анимация, чтобы переезд был плавным).
  setSideRight(v, animate) {
    if (this.root.classList.contains('side-right') === !!v) return;
    const parts = [this.side, this.viewer];
    const before = parts.map((n) => n.getBoundingClientRect().left);
    this.root.classList.toggle('side-right', !!v);
    if (!animate || !this.root.classList.contains('side-open')) return;
    parts.forEach((n, i) => {
      const dx = before[i] - n.getBoundingClientRect().left;
      if (Math.abs(dx) > 1) n.animate([{ transform: `translateX(${dx}px)` }, { transform: 'translateX(0)' }], { duration: 340, easing: 'cubic-bezier(.3, .7, .2, 1)' });
    });
  }

  bindSideGrip() {
    const g = this.sideGrip;
    g.addEventListener('dblclick', () => {
      const right = !this.root.classList.contains('side-right');
      this.setSideRight(right, true);
      window.folioSidePref && window.folioSidePref(this, right);
    });
    g.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      g.setPointerCapture(e.pointerId);
      const sx = e.clientX;
      this.side.classList.add('dragging');
      document.body.classList.add('grabbing');
      const move = (ev) => { this.side.style.transform = `translateX(${ev.clientX - sx}px)`; };
      const up = (ev) => {
        g.removeEventListener('pointermove', move);
        g.removeEventListener('pointerup', up);
        g.removeEventListener('pointercancel', up);
        document.body.classList.remove('grabbing');
        const r = this.root.getBoundingClientRect();
        const right = ev.clientX > r.left + r.width / 2;
        // Переезжаем с того места, где панель отпустили.
        const from = this.side.getBoundingClientRect().left;
        this.side.style.transform = '';
        this.side.classList.remove('dragging');
        const vFrom = this.viewer.getBoundingClientRect().left;
        this.root.classList.toggle('side-right', right);
        const dx = from - this.side.getBoundingClientRect().left;
        const vx = vFrom - this.viewer.getBoundingClientRect().left;
        const ease = { duration: 360, easing: 'cubic-bezier(.3, .7, .2, 1)' };
        if (Math.abs(dx) > 1) this.side.animate([{ transform: `translateX(${dx}px)` }, { transform: 'translateX(0)' }], ease);
        if (Math.abs(vx) > 1) this.viewer.animate([{ transform: `translateX(${vx}px)` }, { transform: 'translateX(0)' }], ease);
        window.folioSidePref && window.folioSidePref(this, right);
      };
      g.addEventListener('pointermove', move);
      g.addEventListener('pointerup', up);
      g.addEventListener('pointercancel', up);
    });
  }

  isShown() { return !this.root.hidden && !!this.root.offsetParent; }

  bindThumbs() {
    const th = this.thumbs;
    th.addEventListener('click', (e) => {
      const g = e.target.closest('.gap button');
      if (g) { this.insertBlankAt(+g.parentNode.dataset.at); return; }
      const act = e.target.closest('[data-act]');
      const t = e.target.closest('.thumb');
      if (!t) return;
      const pid = +t.dataset.pid;
      if (act) {
        e.stopPropagation();
        if (act.dataset.act === 'rotate') this.rotatePages([pid], 90);
        else this.deletePages([pid]);
        return;
      }
      if (this.thumbDragged) return;
      if (e.ctrlKey || e.metaKey) {
        this.picked.has(pid) ? this.picked.delete(pid) : this.picked.add(pid);
      } else if (e.shiftKey && this.cur) {
        const a = this.pages.findIndex((p) => p.id === this.cur), b = this.pages.findIndex((p) => p.id === pid);
        for (let i = Math.min(a, b); i <= Math.max(a, b); i++) this.picked.add(this.pages[i].id);
      } else {
        this.picked.clear();
        this.scrollToPage(pid);
      }
      this.rebuildThumbs();
    });
    th.addEventListener('contextmenu', (e) => {
      const t = e.target.closest('.thumb');
      if (!t) return;
      e.preventDefault();
      const pid = +t.dataset.pid;
      const ids = this.picked.has(pid) ? [...this.picked] : [pid];
      const many = ids.length > 1;
      const i = this.pages.findIndex((p) => p.id === pid);
      showMenu(e.clientX, e.clientY, [
        ['Повернуть вправо', () => this.rotatePages(ids, 90)],
        ['Повернуть влево', () => this.rotatePages(ids, 270)],
        ['Дублировать', () => this.duplicatePages(ids)],
        ['Вставить пустую страницу', () => this.insertBlankAt(i + 1)],
        null,
        [many ? `Сохранить ${ids.length} стр. отдельным PDF` : 'Сохранить страницу отдельным PDF', () => this.extract(ids)],
        null,
        [many ? `Удалить ${ids.length} стр.` : 'Удалить страницу', () => this.deletePages(ids), 'Delete'],
      ]);
    });
    // Перетаскивание миниатюр: порядок внутри документа или копия в соседний документ.
    th.addEventListener('pointerdown', (e) => {
      const t = e.target.closest('.thumb');
      if (!t || e.button !== 0 || e.target.closest('[data-act]')) return;
      const pid = +t.dataset.pid;
      const sx = e.clientX, sy = e.clientY;
      const ids = () => (this.picked.has(pid) ? this.pages.filter((p) => this.picked.has(p.id)).map((p) => p.id) : [pid]);
      let line = null, target = null, started = false, ghost = null;
      this.thumbDragged = false;
      t.classList.add('pressed');
      const clearMarks = () => {
        line?.remove();
        line = null;
        document.querySelectorAll('.viewer.drop-into').forEach((n) => n.classList.remove('drop-into'));
      };
      const move = (ev) => {
        if (!started && Math.hypot(ev.clientX - sx, ev.clientY - sy) < 6) return;
        if (!started) {
          started = true;
          this.thumbDragged = true;
          t.classList.remove('pressed');
          ghost = dragGhost(t, ids().length, sx, sy);
          for (const id of ids()) this.thumbEls.get(id)?.classList.add('dragging');
          for (const ed of PdfEditor.all) if (ed !== this && ed.isShown()) ed.toggleSide(true);
        }
        ghost.move(ev.clientX, ev.clientY);
        clearMarks();
        target = null;
        ghost.label('');
        const hit = document.elementFromPoint(ev.clientX, ev.clientY);
        const ed = hit && [...PdfEditor.all].find((x) => x.isShown() && x.root.contains(hit));
        if (!ed) return;
        if (hit.closest('.thumbs')) {
          const list = [...ed.thumbEls.values()];
          let at = list.length;
          for (let i = 0; i < list.length; i++) {
            const r = list[i].getBoundingClientRect();
            if (ev.clientY < r.top + r.height / 2) { at = i; break; }
          }
          line = el('div', 'drop-line');
          ed.thumbs.insertBefore(line, list[at] || ed.thumbs.querySelector('.add-pages'));
          const tr = ed.thumbs.getBoundingClientRect();
          if (ev.clientY < tr.top + 40) ed.thumbs.scrollTop -= 12;
          if (ev.clientY > tr.bottom - 40) ed.thumbs.scrollTop += 12;
          target = { ed, at };
          ghost.label(ed === this ? 'Переместить' : 'Копировать сюда', ed !== this);
        } else if (ed !== this && hit.closest('.viewer')) {
          ed.viewer.classList.add('drop-into');
          const pe = hit.closest('.page');
          let at = ed.pages.length;
          if (pe) {
            const i = ed.pages.findIndex((p) => p.id === +pe.dataset.pid);
            const r = pe.getBoundingClientRect();
            at = ev.clientY < r.top + r.height / 2 ? i : i + 1;
          }
          target = { ed, at };
          ghost.label('Копировать сюда', true);
        }
      };
      const up = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        t.classList.remove('pressed');
        if (ghost) target ? ghost.drop() : ghost.cancel();
        clearMarks();
        this.thumbs.querySelectorAll('.dragging').forEach((n) => n.classList.remove('dragging'));
        if (started && target) {
          if (target.ed === this) this.movePages(ids(), target.at);
          else target.ed.importPages(this, ids(), target.at);
        }
        if (started) setTimeout(() => (this.thumbDragged = false), 0);
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
    });
  }

  // Копируем страницы из другого документа вместе со всем, что на них нарисовано.
  async importPages(from, ids, at) {
    this.clearSel();
    const order = from.pages.filter((p) => ids.includes(p.id));
    const srcMap = new Map();
    const out = [];
    for (const p0 of order) {
      const p = structuredClone(p0);
      p.objs = p.objs.filter((o) => !o.pristine);
      p.id = nid();
      if (p.src != null) {
        if (!srcMap.has(p.src)) {
          const fs = from.source(p.src);
          let mine = this.sources.find((s) => s.bytes === fs.bytes);
          if (!mine) {
            const pdf = await pdfjsLib.getDocument({ data: fs.bytes.slice(), isEvalSupported: false }).promise;
            mine = { id: nid(), bytes: fs.bytes, pdf, count: fs.count, pg: new Map() };
            this.sources.push(mine);
          }
          srcMap.set(p.src, mine.id);
        }
        const ns = srcMap.get(p.src);
        for (const o of p.objs) if (o.run) o.run = o.run.replace(/^(img:)?\d+:/, (_m, pre) => `${pre || ''}${ns}:`);
        p.src = ns;
      }
      const idMap = new Map();
      for (const o of p.objs) { const n = nid(); idMap.set(o.id, n); o.id = n; }
      for (const o of p.objs) {
        if (o.cover) o.cover = idMap.get(o.cover);
        if (o.t === 'image') { const n = nid(); this.imgs.set(n, from.imgs.get(o.img)); o.img = n; }
      }
      out.push(p);
    }
    if (!out.length) return;
    this.checkpoint();
    this.pages.splice(clamp(at, 0, this.pages.length), 0, ...out);
    this.rebuild();
    this.scrollToPage(out[0].id);
    toast(out.length === 1 ? 'Страница скопирована' : `Скопировано страниц: ${out.length}`);
  }

  movePages(ids, at) {
    const set = new Set(ids);
    const before = this.pages.map((p) => p.id).join();
    const moving = this.pages.filter((p) => set.has(p.id));
    const target = this.pages[at];
    const rest = this.pages.filter((p) => !set.has(p.id));
    let idx = target ? rest.indexOf(target) : rest.length;
    if (idx < 0) idx = rest.length;
    const next = [...rest.slice(0, idx), ...moving, ...rest.slice(idx)];
    if (next.map((p) => p.id).join() === before) return;
    this.clearSel();
    this.checkpoint();
    this.pages = next;
    this.rebuild();
  }

  rotatePages(ids, deg) {
    this.clearSel();
    this.checkpoint();
    for (const id of ids) {
      const p = this.page(id);
      for (let k = 0; k < deg / 90; k++) {
        const { h } = this.dims(p);
        // Поворот на 90° по часовой: (x, y) → (h − y, x)
        for (const o of p.objs) {
          if (o.pts) for (let i = 0; i < o.pts.length; i += 2) { const x = o.pts[i]; o.pts[i] = h - o.pts[i + 1]; o.pts[i + 1] = x; }
          else if ('x1' in o) { [o.x1, o.y1] = [h - o.y1, o.x1]; [o.x2, o.y2] = [h - o.y2, o.x2]; }
          else if (o.t === 'text' || o.t === 'image') {
            const b = bbox(o);
            const cx = h - (b.y + b.h / 2), cy = b.x + b.w / 2;
            o.x = cx - b.w / 2; o.y = cy - b.h / 2;
          } else { [o.x, o.y, o.w, o.h] = [h - o.y - o.h, o.x, o.h, o.w]; }
        }
        p.rot = (p.rot + 90) % 360;
      }
    }
    this.sel = null;
    this.rebuild();
    this.updateCtx();
  }

  deletePages(ids) {
    if (ids.length >= this.pages.length) { toast('Нельзя удалить все страницы'); return; }
    this.clearSel();
    this.checkpoint();
    const set = new Set(ids);
    this.pages = this.pages.filter((p) => !set.has(p.id));
    for (const id of ids) this.picked.delete(id);
    this.sel = null;
    this.rebuild();
    this.updateCtx();
  }

  duplicatePages(ids) {
    this.clearSel();
    this.checkpoint();
    const out = [];
    for (const p of this.pages) {
      out.push(p);
      if (ids.includes(p.id)) {
        const c = structuredClone(p);
        c.id = nid();
        c.objs.forEach((o) => (o.id = nid()));
        out.push(c);
      }
    }
    this.pages = out;
    this.rebuild();
  }

  insertBlankAt(i) {
    this.clearSel();
    const ref = this.pages[i - 1] || this.pages[i];
    const { w, h } = ref ? this.dims(ref) : { w: A4[2], h: A4[3] };
    this.checkpoint();
    const p = { id: nid(), src: null, index: 0, rot: 0, baseRot: 0, view: [0, 0, w, h], objs: [] };
    this.pages.splice(clamp(i, 0, this.pages.length), 0, p);
    this.rebuild();
    this.scrollToPage(p.id);
  }

  async addPagesDialog() {
    const paths = await window.folio.openDialog();
    if (!paths.length) return;
    const files = [];
    for (const path of paths) files.push({ name: path, bytes: await window.folio.readFile(path) });
    await this.addFiles(files);
  }

  async addFiles(files, at) {
    const ps = await this.pagesFromFiles(files);
    if (!ps.length) return;
    this.checkpoint();
    const idx = at == null ? this.pages.length : at;
    this.pages.splice(idx, 0, ...ps);
    this.rebuild();
    this.scrollToPage(ps[0].id);
    toast(ps.length === 1 ? 'Добавлена 1 страница' : `Добавлено страниц: ${ps.length}`);
  }

  async extract(ids) {
    const order = this.pages.filter((p) => ids.includes(p.id)).map((p) => p.id);
    const base = (this.path || T('Документ.pdf')).replace(/\.[^.\\/]+$/, '');
    const target = await window.folio.saveDialog(`${base} — стр. ${order.map((id) => this.pages.findIndex((p) => p.id === id) + 1).join(',')}.pdf`);
    if (!target) return;
    const bytes = await buildPdf(this, order);
    await window.folio.writeFile(target, bytes);
    toast('Сохранено отдельным файлом');
  }

  // ---------- Поиск ----------
  openFind() {
    if (this.findBar) { this.findBar.querySelector('input').select(); return; }
    const bar = el('div', 'findbar');
    bar.innerHTML = `${ICONS.search}<input type="text" placeholder="Найти в документе"><span class="count"></span>`;
    const input = bar.querySelector('input');
    const count = bar.querySelector('.count');
    const prev = el('button', 'cbtn', ICONS.up);
    const next = el('button', 'cbtn', ICONS.down);
    const x = el('button', 'cbtn', ICONS.close);
    prev.dataset.tip = 'Предыдущее · Shift+Enter';
    next.dataset.tip = 'Следующее · Enter';
    bar.append(prev, next, x);
    this.root.append(bar);
    this.findBar = bar;
    const run = async () => {
      const q = input.value.trim().toLowerCase();
      const hitsList = [];
      if (q) {
        for (const p of this.pages) {
          for (const r of await this.runs(p)) {
            const s = r.str.toLowerCase();
            let i = s.indexOf(q);
            while (i >= 0) {
              hitsList.push({ pid: p.id, x: r.x + (r.w * i) / s.length, y: r.y, w: (r.w * q.length) / s.length, h: r.h });
              i = s.indexOf(q, i + q.length);
            }
          }
        }
      }
      if (input.value.trim().toLowerCase() !== q) return;
      this.find = { q, hits: hitsList, i: 0 };
      count.textContent = q ? (hitsList.length ? `1/${hitsList.length}` : '0') : '';
      this.drawFind(true);
    };
    let timer;
    input.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(run, 160); });
    const go = (d) => {
      if (!this.find || !this.find.hits.length) return;
      this.find.i = (this.find.i + d + this.find.hits.length) % this.find.hits.length;
      count.textContent = `${this.find.i + 1}/${this.find.hits.length}`;
      this.drawFind(true);
    };
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); go(e.shiftKey ? -1 : 1); }
      if (e.key === 'Escape') this.closeFind();
    });
    prev.addEventListener('click', () => go(-1));
    next.addEventListener('click', () => go(1));
    x.addEventListener('click', () => this.closeFind());
    input.focus();
  }

  closeFind() {
    this.findBar?.remove();
    this.findBar = null;
    this.find = null;
    this.drawFind();
  }

  drawFind(scroll) {
    for (const rec of this.els.values()) rec.marks.querySelectorAll('.find-hit').forEach((n) => n.remove());
    if (!this.find) return;
    const S = this.S;
    this.find.hits.forEach((h, i) => {
      const rec = this.els.get(h.pid);
      if (!rec) return;
      const d = el('div', 'find-hit' + (i === this.find.i ? ' cur' : ''));
      Object.assign(d.style, { left: h.x * S + 'px', top: h.y * S + 'px', width: h.w * S + 'px', height: h.h * S + 'px' });
      rec.marks.append(d);
      if (scroll && i === this.find.i) {
        const vr = this.viewer.getBoundingClientRect(), dr = d.getBoundingClientRect();
        if (dr.top < vr.top + 60 || dr.bottom > vr.bottom - 100) this.viewer.scrollTop += dr.top - vr.top - vr.height / 3;
      }
    });
  }

  // ---------- Буфер обмена ----------
  copy() {
    const s = this.selectedObjs();
    if (!s) return false;
    const imgs = new Map();
    for (const o of s.objs) if (o.t === 'image') imgs.set(o.img, this.imgs.get(o.img));
    PdfEditor.clip = { objs: s.objs.map((o) => this.cleanCopy(o)), imgs };
    return true;
  }

  async paste(e) {
    const items = e.clipboardData ? [...e.clipboardData.items] : [];
    const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
    const p = this.currentPage();
    if (!p) return;
    if (PdfEditor.clip && text === CLIP_MARK) {
      this.checkpoint();
      const ids = [];
      for (const src of PdfEditor.clip.objs) {
        const o = structuredClone(src);
        o.id = nid();
        if (o.t === 'image') { const id = nid(); this.imgs.set(id, PdfEditor.clip.imgs.get(o.img)); o.img = id; }
        translate(o, 12, 12);
        p.objs.push(o);
        ids.push(o.id);
      }
      this.refresh(p);
      this.select(p.id, ids);
      return;
    }
    const imgItem = items.find((i) => i.type.startsWith('image/'));
    if (imgItem) {
      const f = imgItem.getAsFile();
      await this.insertImage(new Uint8Array(await f.arrayBuffer()), 'paste.' + (f.type.split('/')[1] || 'png'));
      return;
    }
    if (text && text.trim()) {
      const st = Settings.get('text');
      const { w } = this.dims(p);
      const o = { id: nid(), t: 'text', x: w * 0.1, y: 72, text: text.replace(/\r/g, '').trimEnd(), font: st.font, size: st.size, color: st.color, bold: st.bold, italic: st.italic };
      this.checkpoint();
      p.objs.push(o);
      this.refresh(p);
      this.select(p.id, [o.id]);
    }
  }

  // ---------- Клавиатура ----------
  onKey(e) {
    if (this.editing) return false;
    const k = e.key.toLowerCase();
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && (k === 'z' || k === 'я')) { e.shiftKey ? this.redo() : this.undo(); return true; }
    if (ctrl && (k === 'y' || k === 'н')) { this.redo(); return true; }
    if (ctrl && (k === 'f' || k === 'а')) { this.openFind(); return true; }
    if (ctrl && (k === 'b' || k === 'и')) { this.toggleSide(); return true; }
    if (ctrl && (k === 'd' || k === 'в')) { this.duplicateSel(); return true; }
    if (ctrl && (k === 'a' || k === 'ф')) {
      const p = this.currentPage();
      if (p) { this.setTool('select'); this.select(p.id, p.objs.filter((o) => !o.pristine && !(o.t === 'whiteout' && o.run)).map((o) => o.id)); }
      return true;
    }
    if (ctrl && (k === '=' || k === '+')) { this.zoomStep(1); return true; }
    if (ctrl && k === '-') { this.zoomStep(-1); return true; }
    if (ctrl && k === '0') { this.fitWidth(); this.flashHud(Math.round(this.zoom * 100) + '%'); return true; }
    if (ctrl) return false;
    if (e.key === 'Delete' || e.key === 'Backspace') { this.deleteSel(); return true; }
    if (e.key === 'Escape') {
      if (this.findBar) this.closeFind();
      else if (this.sel) this.select(null);
      else if (this.tool !== 'select') this.setTool('select');
      else return false;
      return true;
    }
    if (e.key.startsWith('Arrow') && this.sel) {
      const d = e.shiftKey ? 10 : 1;
      this.nudge(e.key === 'ArrowLeft' ? -d : e.key === 'ArrowRight' ? d : 0, e.key === 'ArrowUp' ? -d : e.key === 'ArrowDown' ? d : 0);
      return true;
    }
    if (e.key === 'Enter' && this.sel) {
      const s = this.selected();
      if (s && s.o.t === 'text') { this.startEdit(s.p, s.o); return true; }
    }
    const map = { v: 'select', м: 'select', t: 'text', е: 'text', p: 'pen', з: 'pen', h: 'marker', р: 'marker', r: 'shape', к: 'shape', e: 'whiteout', у: 'whiteout' };
    if (map[k]) { this.setTool(map[k]); return true; }
    if (k === 'i' || k === 'ш') { this.pickImage(); return true; }
    if (k === 's' || k === 'ы') { this.openSignature(); return true; }
    if (e.key === 'PageDown' || e.key === 'PageUp') {
      const i = this.pages.findIndex((p) => p.id === this.cur);
      const n = this.pages[clamp(i + (e.key === 'PageDown' ? 1 : -1), 0, this.pages.length - 1)];
      if (n) this.scrollToPage(n.id);
      return true;
    }
    return false;
  }

  // ---------- Сохранение ----------
  async save(asNew) {
    this.endEdit();
    let target = this.path && /\.pdf$/i.test(this.path) && !asNew ? this.path : null;
    if (!target) {
      const base = (this.path || T('Документ')).replace(/\.[^.\\/]+$/, '');
      target = await window.folio.saveDialog(base + '.pdf');
      if (!target) return false;
    }
    const bytes = await buildPdf(this);
    await window.folio.writeFile(target, bytes);
    this.path = target;
    this.dirty = false;
    this.onChange && this.onChange();
    return true;
  }

  destroy() {
    PdfEditor.all.delete(this);
    this.ro.disconnect();
    for (const s of this.sources) s.pdf.destroy();
    this.root.remove();
  }
}

// «Призрак» страницы, которую несут мышью: летит за курсором с инерцией и наклоном.
function dragGhost(t, count, sx, sy) {
  const paper = t.querySelector('.t-paper');
  const pr = paper.getBoundingClientRect();
  const g = el('div', 'drag-ghost');
  const src = paper.querySelector('canvas');
  const card = el('div', 'dg-card');
  card.style.width = pr.width + 'px';
  card.style.height = pr.height + 'px';
  if (src) {
    const c = document.createElement('canvas');
    c.width = src.width;
    c.height = src.height;
    c.getContext('2d').drawImage(src, 0, 0);
    card.append(c);
  }
  if (count > 1) {
    card.classList.add('stack');
    card.append(el('span', 'dg-count', String(count)));
  }
  const lab = el('div', 'dg-label');
  g.append(card, lab);
  document.body.append(g);
  document.body.classList.add('grabbing');
  // Точка, за которую «взяли» страницу, остаётся под курсором.
  const ox = sx - pr.left, oy = sy - pr.top;
  let x = pr.left, y = pr.top, tx = x, ty = y, rot = 0, raf = 0, alive = true;
  const frame = () => {
    const vx = tx - x;
    x += vx * 0.3;
    y += (ty - y) * 0.3;
    rot += (clamp(vx * 0.35, -9, 9) - rot) * 0.2;
    g.style.transform = `translate(${x}px, ${y}px) rotate(${rot}deg)`;
    if (alive) raf = requestAnimationFrame(frame);
  };
  raf = requestAnimationFrame(frame);
  requestAnimationFrame(() => g.classList.add('lifted'));
  const finish = (cls, toX, toY) => {
    alive = false;
    cancelAnimationFrame(raf);
    document.body.classList.remove('grabbing');
    g.classList.add(cls);
    if (toX != null) g.style.transform = `translate(${toX}px, ${toY}px) rotate(0deg)`;
    setTimeout(() => g.remove(), 320);
  };
  return {
    move(cx, cy) { tx = cx - ox; ty = cy - oy; },
    label(text, copy) { lab.textContent = text; lab.hidden = !text; g.classList.toggle('copy', !!copy); },
    drop() { finish('dropped', x, y); },
    cancel() { finish('back', pr.left, pr.top); },
  };
}

PdfEditor.all = new Set();
PdfEditor.clip = null;
// Метка в системном буфере: «там лежат объекты Folio, а не текст».
const CLIP_MARK = '\u2063folio-objects\u2063';

function trimCanvas(cv) {
  const g = cv.getContext('2d');
  const { width: w, height: h } = cv;
  const d = g.getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = 0, y1 = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (d[(y * w + x) * 4 + 3] > 8) {
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
      }
    }
  }
  const pad = 6;
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(w, x1 + pad); y1 = Math.min(h, y1 + pad);
  const out = document.createElement('canvas');
  out.width = x1 - x0; out.height = y1 - y0;
  out.getContext('2d').drawImage(cv, x0, y0, out.width, out.height, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}
