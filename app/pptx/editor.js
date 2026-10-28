// Презентации PowerPoint (.pptx): слайды лентой, миниатюры сбоку, правка текста прямо на слайде,
// новые слайды из макетов, копия, удаление, порядок. Нетронутые слайды сохраняются как были.

const PNS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
};
const PP_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const EMU = 9525; // EMU в пикселе (96 dpi)
const PP_GAP = 26;
const PP_PRST = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', darkBlue: '00008B', darkGray: 'A9A9A9', lightGray: 'D3D3D3', navy: '000080', silver: 'C0C0C0' };

const kids = (e, name, ns = PNS.a) => (e ? [...e.children].filter((k) => k.localName === name && (!ns || k.namespaceURI === ns)) : []);
const kid = (e, name, ns) => kids(e, name, ns)[0] || null;
const kidAny = (e, name) => (e ? [...e.children].find((k) => k.localName === name) || null : null);
const attrN = (e, n, d) => (e && e.hasAttribute(n) ? +e.getAttribute(n) : d);

// ---------- Цвета ----------
function hex2rgb(h) { return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)); }
function rgb2hsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}
function hsl2rgb([h, s, l]) {
  if (!s) return [l, l, l].map((v) => Math.round(v * 255));
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)].map((v) => Math.round(clamp(v, 0, 1) * 255));
}

class PptxEditor {
  constructor(host, { onChange }) {
    this.kind = 'pptx';
    this.onChange = onChange;
    this.dirty = false;
    this.zoom = 1;
    this.undoStack = [];
    this.media = new Map();

    this.root = el('div', 'doc pp');
    this.side = el('aside', 'side');
    this.sideGrip = el('div', 'side-grip');
    this.sideGrip.dataset.tip = 'Перетащите панель влево или вправо';
    this.thumbs = el('div', 'thumbs pp-thumbs');
    this.side.append(this.sideGrip, this.thumbs);
    this.viewer = el('div', 'viewer pp-viewer');
    this.viewer.tabIndex = -1;
    this.col = el('div', 'pp-col');
    this.viewer.append(this.col);
    this.hud = el('div', 'hud');
    this.root.append(this.side, this.viewer, this.hud);
    host.append(this.root);
    this.root.classList.add('side-open');
    this.thumbEls = [];

    this.buildDock();
    this.bindSideGrip();
    this.bindThumbs();
    this.viewer.addEventListener('scroll', () => this.trackCurrent(), { passive: true });
    this.viewer.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    this.ro = new ResizeObserver(() => { if (this.fitMode !== false && this.slides) this.fitWidth(); });
    this.ro.observe(this.viewer);
    this.onSel = () => this.syncTools();
    document.addEventListener('selectionchange', this.onSel);
  }

  // ---------- Загрузка ----------
  async open(files) {
    const loading = el('div', 'loading', 'Открываю…');
    this.root.append(loading);
    try {
      await this.load(files[0].bytes);
      await this.renderAll();
    } finally {
      loading.remove();
    }
  }

  async blank() {
    await this.load(await makeBlankPptx());
    await this.renderAll();
  }

  async readXml(path) {
    const f = this.zip.file(path);
    if (!f) return null;
    const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    if (x.getElementsByTagName('parsererror').length) throw new Error('Повреждённый файл: ' + path);
    return x;
  }
  resolve(target, base) {
    if (target.startsWith('/')) return target.slice(1);
    const out = [];
    for (const s of (base + target).split('/')) {
      if (s === '..') out.pop();
      else if (s && s !== '.') out.push(s);
    }
    return out.join('/');
  }
  relsPathOf(path) { const i = path.lastIndexOf('/'); return path.slice(0, i + 1) + '_rels/' + path.slice(i + 1) + '.rels'; }
  async part(path) {
    if (this.parts.has(path)) return this.parts.get(path);
    const p = { path, base: path.replace(/[^/]*$/, ''), xml: await this.readXml(path) };
    p.relsXml = await this.readXml(this.relsPathOf(path));
    p.rels = {};
    for (const r of p.relsXml ? p.relsXml.documentElement.children : []) {
      p.rels[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: r.getAttribute('Target') || '', ext: r.getAttribute('TargetMode') === 'External', el: r };
    }
    this.parts.set(path, p);
    return p;
  }
  relOf(p, kind) {
    const e = Object.entries(p.rels).find(([, r]) => r.type.endsWith('/' + kind));
    return e ? { id: e[0], path: this.resolve(e[1].target, p.base), ...e[1] } : null;
  }

  async load(bytes) {
    this.zip = await JSZip.loadAsync(bytes);
    this.parts = new Map();
    const rr = await this.readXml('_rels/.rels');
    let presPath = 'ppt/presentation.xml';
    for (const r of rr ? rr.documentElement.children : []) if ((r.getAttribute('Type') || '').endsWith('/officeDocument')) presPath = this.resolve(r.getAttribute('Target'), '');
    this.pres = await this.part(presPath);
    this.ct = await this.readXml('[Content_Types].xml');
    const P = this.pres.xml.documentElement;
    const sz = kid(P, 'sldSz', PNS.p);
    this.W = Math.round(attrN(sz, 'cx', 12192000) / EMU);
    this.H = Math.round(attrN(sz, 'cy', 6858000) / EMU);
    this.defStyle = kid(P, 'defaultTextStyle', PNS.p);
    this.slides = [];
    for (const s of kids(kid(P, 'sldIdLst', PNS.p), 'sldId', PNS.p)) {
      const rid = s.getAttributeNS(PNS.r, 'id') || s.getAttribute('r:id');
      const rel = this.pres.rels[rid];
      if (!rel) continue;
      const sl = await this.part(this.resolve(rel.target, this.pres.base));
      sl.id = +s.getAttribute('id');
      sl.rid = rid;
      sl.sldIdEl = s;
      await this.linkSlide(sl);
      this.slides.push(sl);
    }
    this.deleted = [];
    this.fitMode = true;
  }

  // Макет, образец, тема — для слайда.
  async linkSlide(sl) {
    const lr = this.relOf(sl, 'slideLayout');
    sl.layout = lr ? await this.part(lr.path) : null;
    if (sl.layout && !sl.layout.master) {
      const mr = this.relOf(sl.layout, 'slideMaster');
      sl.layout.master = mr ? await this.part(mr.path) : null;
    }
    const m = sl.layout && sl.layout.master;
    if (m && !m.theme) {
      const tr = this.relOf(m, 'theme');
      m.theme = tr ? this.parseTheme(await this.readXml(tr.path)) : this.parseTheme(null);
      const cm = kid(m.xml.documentElement, 'clrMap', PNS.p);
      m.clrMap = {};
      if (cm) for (const a of cm.attributes) m.clrMap[a.name] = a.value;
      const tx = kid(m.xml.documentElement, 'txStyles', PNS.p);
      m.txStyles = { title: kid(tx, 'titleStyle', PNS.p), body: kid(tx, 'bodyStyle', PNS.p), other: kid(tx, 'otherStyle', PNS.p) };
    }
  }

  parseTheme(x) {
    const t = { colors: {}, major: 'Calibri Light', minor: 'Calibri', fills: [], lines: [], bgFills: [] };
    if (!x) return t;
    const cs = x.getElementsByTagNameNS(PNS.a, 'clrScheme')[0];
    for (const c of cs ? cs.children : []) {
      const v = c.firstElementChild;
      if (v) t.colors[c.localName] = v.getAttribute('lastClr') || v.getAttribute('val') || '000000';
    }
    const fs = x.getElementsByTagNameNS(PNS.a, 'fontScheme')[0];
    if (fs) {
      const mj = kid(kid(fs, 'majorFont'), 'latin'), mn = kid(kid(fs, 'minorFont'), 'latin');
      if (mj) t.major = mj.getAttribute('typeface');
      if (mn) t.minor = mn.getAttribute('typeface');
    }
    const fm = x.getElementsByTagNameNS(PNS.a, 'fmtScheme')[0];
    if (fm) {
      t.fills = [...(kid(fm, 'fillStyleLst') || { children: [] }).children];
      t.lines = [...(kid(fm, 'lnStyleLst') || { children: [] }).children];
      t.bgFills = [...(kid(fm, 'bgFillStyleLst') || { children: [] }).children];
    }
    return t;
  }

  // ---------- Цвет и заливка ----------
  // el — элемент цвета (srgbClr, schemeClr…); ctx — { m: образец, ph: цвет-заместитель }
  color(c, ctx) {
    if (!c) return null;
    let hex = null;
    const n = c.localName;
    const theme = ctx.m ? ctx.m.theme : { colors: {} };
    if (n === 'srgbClr') hex = c.getAttribute('val');
    else if (n === 'sysClr') hex = c.getAttribute('lastClr') || '000000';
    else if (n === 'prstClr') hex = PP_PRST[c.getAttribute('val')] || '000000';
    else if (n === 'scrgbClr') hex = ['r', 'g', 'b'].map((k) => Math.round(clamp(attrN(c, k, 0) / 100000, 0, 1) * 255).toString(16).padStart(2, '0')).join('');
    else if (n === 'schemeClr') {
      let v = c.getAttribute('val');
      if (v === 'phClr') return ctx.ph || '#000000';
      const map = ctx.clrMap || (ctx.m && ctx.m.clrMap) || {};
      if (map[v]) v = map[v];
      else if (v === 'bg1') v = 'lt1'; else if (v === 'tx1') v = 'dk1'; else if (v === 'bg2') v = 'lt2'; else if (v === 'tx2') v = 'dk2';
      hex = theme.colors[v] || '000000';
    } else if (n === 'hslClr') {
      hex = hsl2rgb([attrN(c, 'hue', 0) / 21600000, attrN(c, 'sat', 0) / 100000, attrN(c, 'lum', 0) / 100000]).map((v) => v.toString(16).padStart(2, '0')).join('');
    }
    if (!hex) return null;
    let rgb = hex2rgb(hex), alpha = 1;
    for (const m of c.children) {
      const v = attrN(m, 'val', 0) / 100000;
      switch (m.localName) {
        case 'lumMod': { const h = rgb2hsl(rgb); h[2] *= v; rgb = hsl2rgb(h); break; }
        case 'lumOff': { const h = rgb2hsl(rgb); h[2] = clamp(h[2] + v, 0, 1); rgb = hsl2rgb(h); break; }
        case 'satMod': { const h = rgb2hsl(rgb); h[1] = clamp(h[1] * v, 0, 1); rgb = hsl2rgb(h); break; }
        case 'tint': rgb = rgb.map((x) => Math.round(255 - (255 - x) * v)); break;
        case 'shade': rgb = rgb.map((x) => Math.round(x * v)); break;
        case 'alpha': alpha = v; break;
      }
    }
    return alpha < 1 ? `rgba(${rgb.join(',')},${alpha.toFixed(3)})` : '#' + rgb.map((x) => x.toString(16).padStart(2, '0')).join('');
  }
  colorIn(e, ctx) {
    if (!e) return null;
    const c = [...e.children].find((k) => /Clr$/.test(k.localName));
    return this.color(c, ctx);
  }

  // Заливка → CSS background (или null — нет заливки, undefined — не задана).
  async fill(spPr, ctx) {
    if (!spPr) return undefined;
    for (const f of spPr.children) {
      switch (f.localName) {
        case 'noFill': return null;
        case 'solidFill': return this.colorIn(f, ctx);
        case 'gradFill': {
          const gs = kids(kid(f, 'gsLst'), 'gs').map((g) => `${this.colorIn(g, ctx)} ${attrN(g, 'pos', 0) / 1000}%`);
          if (!gs.length) return null;
          const lin = kid(f, 'lin');
          if (kid(f, 'path')) return `radial-gradient(circle, ${gs.join(', ')})`;
          const ang = lin ? attrN(lin, 'ang', 0) / 60000 : 90;
          return `linear-gradient(${(ang + 90) % 360}deg, ${gs.join(', ')})`;
        }
        case 'blipFill': {
          const url = await this.blipUrl(f, ctx);
          return url ? `center / cover no-repeat url("${url}")` : null;
        }
        case 'pattFill': return this.colorIn(kid(f, 'fgClr'), ctx);
        case 'grpFill': return ctx.grpFill;
      }
    }
    return undefined;
  }

  async blipUrl(blipFill, ctx) {
    const blip = kid(blipFill, 'blip');
    if (!blip) return null;
    const rid = blip.getAttributeNS(PNS.r, 'embed') || blip.getAttribute('r:embed');
    const part = ctx.part;
    const rel = part && part.rels[rid];
    if (!rel || rel.ext) return null;
    const path = this.resolve(rel.target, part.base);
    if (!this.media.has(path)) {
      const f = this.zip.file(path);
      if (!f) return null;
      const ext = path.split('.').pop().toLowerCase();
      const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/emf', wmf: 'image/wmf', tif: 'image/tiff', tiff: 'image/tiff' }[ext] || 'application/octet-stream';
      this.media.set(path, f.async('uint8array').then((b) => URL.createObjectURL(new Blob([b], { type }))));
    }
    return this.media.get(path);
  }

  lineStyle(spPr, style, ctx) {
    const ln = kid(spPr, 'ln');
    let color, w = 0, dash = null;
    if (ln) {
      if (kid(ln, 'noFill')) return null;
      w = attrN(ln, 'w', 9525) / EMU;
      const sf = kid(ln, 'solidFill');
      if (sf) color = this.colorIn(sf, ctx);
      const gf = kid(ln, 'gradFill');
      if (!color && gf) color = this.colorIn(kid(kid(gf, 'gsLst'), 'gs'), ctx);
      const pd = kid(ln, 'prstDash');
      if (pd) dash = pd.getAttribute('val');
    }
    if (!color && style) {
      const ref = kid(style, 'lnRef');
      const idx = attrN(ref, 'idx', 0);
      if (idx > 0) {
        color = this.colorIn(ref, ctx);
        const th = ctx.m && ctx.m.theme.lines[idx - 1];
        if (!w) w = th ? attrN(th, 'w', 9525) / EMU : 1;
      }
    }
    if (!color) return null;
    return { color, w: Math.max(w || 1, 0.75), dash, head: ln && kid(ln, 'headEnd'), tail: ln && kid(ln, 'tailEnd') };
  }

  // ---------- Наследование заполнителей ----------
  phOf(sp) {
    const nv = kidAny(sp, 'nvSpPr') || kidAny(sp, 'nvPicPr') || kidAny(sp, 'nvGraphicFramePr');
    const ph = nv && kidAny(kidAny(nv, 'nvPr'), 'ph');
    if (!ph) return null;
    return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') };
  }
  findPh(part, ph) {
    if (!part || !ph) return null;
    const tree = part.xml.getElementsByTagNameNS(PNS.p, 'spTree')[0];
    const all = tree ? [...tree.getElementsByTagNameNS(PNS.p, 'sp')] : [];
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
    let hit = ph.idx != null ? all.find((s) => { const q = this.phOf(s); return q && q.idx === ph.idx; }) : null;
    if (!hit) hit = all.find((s) => { const q = this.phOf(s); return q && q.type === ph.type; });
    if (!hit) hit = all.find((s) => { const q = this.phOf(s); return q && norm(q.type) === norm(ph.type); });
    return hit || null;
  }

  // ---------- Свойства текста ----------
  parseRPr(r, ctx) {
    if (!r) return {};
    const o = {};
    if (r.hasAttribute('sz')) o.sz = +r.getAttribute('sz') / 100;
    if (r.hasAttribute('b')) o.b = r.getAttribute('b') === '1' || r.getAttribute('b') === 'true';
    if (r.hasAttribute('i')) o.i = r.getAttribute('i') === '1' || r.getAttribute('i') === 'true';
    if (r.hasAttribute('u')) o.u = r.getAttribute('u') !== 'none';
    if (r.hasAttribute('strike')) o.strike = r.getAttribute('strike') !== 'noStrike';
    if (r.hasAttribute('baseline')) o.baseline = +r.getAttribute('baseline') / 1000;
    if (r.hasAttribute('cap')) o.cap = r.getAttribute('cap');
    if (r.hasAttribute('spc')) o.spc = +r.getAttribute('spc') / 100;
    const sf = kid(r, 'solidFill');
    if (sf) o.color = this.colorIn(sf, ctx);
    const gf = kid(r, 'gradFill');
    if (gf && !o.color) o.color = this.colorIn(kid(kid(gf, 'gsLst'), 'gs'), ctx);
    const lat = kid(r, 'latin') || kid(r, 'ea') || kid(r, 'cs');
    if (lat) o.font = lat.getAttribute('typeface');
    const hl = kid(r, 'highlight');
    if (hl) o.hl = this.colorIn(hl, ctx);
    return o;
  }
  parsePPr(p, ctx) {
    if (!p) return {};
    const o = {};
    if (p.hasAttribute('algn')) o.algn = p.getAttribute('algn');
    if (p.hasAttribute('marL')) o.marL = +p.getAttribute('marL') / EMU;
    if (p.hasAttribute('indent')) o.indent = +p.getAttribute('indent') / EMU;
    const sp = (n) => { const e = kid(p, n); if (!e) return undefined; const pct = kid(e, 'spcPct'), pts = kid(e, 'spcPts'); return pct ? { pct: attrN(pct, 'val', 100000) / 100000 } : pts ? { pt: attrN(pts, 'val', 0) / 100 } : undefined; };
    const ls = sp('lnSpc'), sb = sp('spcBef'), sa = sp('spcAft');
    if (ls) o.lnSpc = ls;
    if (sb) o.spcBef = sb;
    if (sa) o.spcAft = sa;
    if (kid(p, 'buNone')) o.bu = { none: true };
    const bc = kid(p, 'buChar'), ba = kid(p, 'buAutoNum'), bb = kid(p, 'buBlip');
    if (bc) o.bu = { char: bc.getAttribute('char') };
    if (ba) o.bu = { auto: ba.getAttribute('type') || 'arabicPeriod', start: attrN(ba, 'startAt', 1) };
    if (bb) o.bu = { char: '•' };
    const bf = kid(p, 'buFont');
    if (bf) o.buFont = bf.getAttribute('typeface');
    const bcl = kid(p, 'buClr');
    if (bcl) o.buColor = this.colorIn(bcl, ctx);
    const bsz = kid(p, 'buSzPct');
    if (bsz) o.buSz = attrN(bsz, 'val', 100000) / 100000;
    const d = kid(p, 'defRPr');
    if (d) o.def = this.parseRPr(d, ctx);
    return o;
  }
  lvlFrom(list, lvl, ctx) {
    if (!list) return null;
    const e = kid(list, `lvl${lvl + 1}pPr`) || (lvl === 0 ? null : null);
    return e ? this.parsePPr(e, ctx) : null;
  }
  merge(...objs) {
    const out = {};
    for (const o of objs) {
      if (!o) continue;
      for (const [k, v] of Object.entries(o)) {
        if (v === undefined) continue;
        if (k === 'def') out.def = { ...(out.def || {}), ...v };
        else out[k] = v;
      }
    }
    return out;
  }

  // ---------- Отрисовка ----------
  async renderAll() {
    this.col.textContent = '';
    this.slideEls = [];
    for (let i = 0; i < this.slides.length; i++) {
      const el2 = await this.renderSlide(this.slides[i], i);
      this.col.append(el2);
      this.slideEls.push(el2);
    }
    this.applyZoom();
    if (this.fitMode) requestAnimationFrame(() => this.fitWidth());
    this.buildThumbs();
    this.trackCurrent();
  }

  async renderSlide(sl, index) {
    const wrap = el('div', 'pp-page');
    const box = el('div', 'pp-slide');
    box.style.width = this.W + 'px';
    box.style.height = this.H + 'px';
    wrap.append(box);
    wrap.slide = sl;
    sl.el = box;
    const m = sl.layout && sl.layout.master;
    const ctx = (part) => ({ m, part, clrMap: this.clrMapFor(sl, m) });
    // Фон: слайд → макет → образец
    let bg = null;
    for (const part of [sl, sl.layout, m]) {
      if (!part) continue;
      const b = part.xml.getElementsByTagNameNS(PNS.p, 'bg')[0];
      if (!b) continue;
      const bgPr = kid(b, 'bgPr', PNS.p);
      if (bgPr) { bg = await this.fill(bgPr, ctx(part)); if (bg !== undefined) break; }
      const ref = kid(b, 'bgRef', PNS.p);
      if (ref) {
        const idx = attrN(ref, 'idx', 0);
        const ph = this.colorIn(ref, ctx(part));
        const th = m && (idx >= 1001 ? m.theme.bgFills[idx - 1001] : m.theme.fills[idx - 1]);
        bg = th ? await this.fill({ children: [th] }, { ...ctx(part), ph }) : ph;
        if (bg === undefined) bg = ph;
        break;
      }
    }
    box.style.background = bg || '#ffffff';
    const showMaster = (part) => !part || part.xml.documentElement.getAttribute('showMasterSp') !== '0';
    if (showMaster(sl)) {
      if (m && showMaster(sl.layout)) await this.renderTree(m, box, ctx(m), sl, false, index);
      if (sl.layout) await this.renderTree(sl.layout, box, ctx(sl.layout), sl, false, index);
    }
    await this.renderTree(sl, box, ctx(sl), sl, true, index);
    if (sl.xml.documentElement.getAttribute('show') === '0') wrap.classList.add('pp-hidden');
    return wrap;
  }

  clrMapFor(sl, m) {
    let map = m ? { ...m.clrMap } : {};
    for (const part of [sl.layout, sl]) {
      if (!part) continue;
      const ov = kid(part.xml.documentElement, 'clrMapOvr', PNS.p);
      const o = ov && kidAny(ov, 'overrideClrMapping');
      if (o) { map = {}; for (const a of o.attributes) map[a.name] = a.value; }
    }
    return map;
  }

  async renderTree(part, box, ctx, sl, own, index) {
    const tree = part.xml.getElementsByTagNameNS(PNS.p, 'spTree')[0];
    if (!tree) return;
    const ident = (x, y, w, h) => ({ x, y, w, h });
    for (const n of tree.children) await this.renderNode(n, box, ctx, sl, own, index, ident, part);
  }

  xfrmOf(n, sl, own) {
    const spPr = kidAny(n, 'spPr') || kidAny(n, 'grpSpPr');
    let x = kid(spPr, 'xfrm') || kidAny(n, 'xfrm');
    if (!x && own !== null) {
      const ph = this.phOf(n);
      if (ph) {
        for (const part of [sl.layout, sl.layout && sl.layout.master]) {
          const s = this.findPh(part, ph);
          const sx = s && kid(kidAny(s, 'spPr'), 'xfrm');
          if (sx) { x = sx; break; }
        }
      }
    }
    if (!x) return null;
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    return {
      x: attrN(off, 'x', 0) / EMU, y: attrN(off, 'y', 0) / EMU, w: attrN(ext, 'cx', 0) / EMU, h: attrN(ext, 'cy', 0) / EMU,
      rot: attrN(x, 'rot', 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1',
      chOff: kid(x, 'chOff'), chExt: kid(x, 'chExt'),
    };
  }

  async renderNode(n, box, ctx, sl, own, index, T, part) {
    const name = n.localName;
    if (name === 'AlternateContent') {
      const ch = kidAny(n, 'Choice') || kidAny(n, 'Fallback');
      const fb = kidAny(n, 'Fallback');
      for (const k of (fb || ch) ? (fb || ch).children : []) await this.renderNode(k, box, ctx, sl, own, index, T, part);
      return;
    }
    if (!['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'].includes(name)) return;
    // Заполнители макета и образца на слайде не рисуем — их заменяют заполнители слайда
    if (!own && this.phOf(n)) return;
    const xf = this.xfrmOf(n, sl, own ? true : null);
    if (!xf && name !== 'grpSp') return;
    const r = T(xf.x, xf.y, xf.w, xf.h);
    if (name === 'grpSp') {
      if (!xf) return;
      const cx = attrN(xf.chOff, 'x', xf.x * EMU) / EMU, cy = attrN(xf.chOff, 'y', xf.y * EMU) / EMU;
      const cw = attrN(xf.chExt, 'cx', xf.w * EMU) / EMU || 1, ch = attrN(xf.chExt, 'cy', xf.h * EMU) / EMU || 1;
      const kx = r.w / cw, ky = r.h / ch;
      const T2 = (x, y, w, h) => ({ x: r.x + (x - cx) * kx, y: r.y + (y - cy) * ky, w: w * kx, h: h * ky });
      const g = el('div', 'pp-grp');
      Object.assign(g.style, { left: 0, top: 0, width: '100%', height: '100%' });
      if (xf.rot || xf.flipH || xf.flipV) { g.style.transformOrigin = `${r.x + r.w / 2}px ${r.y + r.h / 2}px`; g.style.transform = this.tf(xf); }
      box.append(g);
      const grpFill = await this.fill(kidAny(n, 'grpSpPr'), ctx);
      for (const k of n.children) await this.renderNode(k, g, { ...ctx, grpFill }, sl, own, index, T2, part);
      return;
    }
    const shape = el('div', 'pp-sp');
    Object.assign(shape.style, { left: r.x + 'px', top: r.y + 'px', width: Math.max(r.w, 0) + 'px', height: Math.max(r.h, 0) + 'px' });
    if (xf.rot || xf.flipH || xf.flipV) shape.style.transform = this.tf(xf);
    box.append(shape);
    if (name === 'pic') { await this.renderPic(n, shape, ctx, r); return; }
    if (name === 'graphicFrame') { await this.renderFrame(n, shape, ctx, sl, own, index, r); return; }
    // sp / cxnSp
    const spPr = kidAny(n, 'spPr');
    const style = kidAny(n, 'style');
    let fill = await this.fill(spPr, ctx);
    if (fill === undefined && style) {
      const ref = kid(style, 'fillRef');
      const idx = attrN(ref, 'idx', 0);
      if (idx > 0) {
        const phc = this.colorIn(ref, ctx);
        const th = ctx.m && ctx.m.theme.fills[idx - 1];
        fill = th ? await this.fill({ children: [th] }, { ...ctx, ph: phc }) : phc;
        if (fill === undefined) fill = phc;
      }
    }
    // Заполнитель наследует заливку/линию из макета, если своей нет
    if (fill === undefined && own) {
      const ph = this.phOf(n);
      if (ph) for (const p2 of [sl.layout, sl.layout && sl.layout.master]) {
        const s = this.findPh(p2, ph);
        if (!s) continue;
        const f2 = await this.fill(kidAny(s, 'spPr'), { ...ctx, part: p2 });
        if (f2 !== undefined) { fill = f2; break; }
      }
    }
    const line = this.lineStyle(spPr, style, ctx);
    const geom = kid(spPr, 'prstGeom');
    const prst = geom ? geom.getAttribute('prst') : kid(spPr, 'custGeom') ? 'cust' : 'rect';
    const isLine = name === 'cxnSp' || /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst);
    if (isLine || (prst !== 'rect' && prst !== 'roundRect' && prst !== 'ellipse') || (fill && /gradient|url/.test(fill) && prst !== 'rect')) {
      shape.append(this.svgShape(prst, r.w, r.h, fill, line, isLine, kid(spPr, 'custGeom'), geom));
    } else {
      if (fill) shape.style.background = fill;
      if (line) shape.style.outline = `${line.w}px ${line.dash && line.dash !== 'solid' ? 'dashed' : 'solid'} ${line.color}`, shape.style.outlineOffset = -line.w / 2 + 'px';
      if (prst === 'ellipse') shape.style.borderRadius = '50%';
      if (prst === 'roundRect') {
        const av = geom && [...(kid(geom, 'avLst') || { children: [] }).children].find((g) => g.getAttribute('name') === 'adj');
        const adj = av ? +av.getAttribute('fmla').replace(/\D/g, '') / 100000 : 0.16667;
        shape.style.borderRadius = Math.min(r.w, r.h) * adj + 'px';
      }
    }
    // Тень — лёгкая, если задана
    const eff = kid(spPr, 'effectLst');
    if (eff && kid(eff, 'outerShdw') && !isLine) shape.style.filter = 'drop-shadow(0 2px 4px rgba(0,0,0,.25))';
    const tx = kidAny(n, 'txBody');
    if (tx) this.renderText(tx, shape, n, ctx, sl, own, index, r);
  }

  tf(xf) {
    const t = [];
    if (xf.rot) t.push(`rotate(${xf.rot}deg)`);
    if (xf.flipH) t.push('scaleX(-1)');
    if (xf.flipV) t.push('scaleY(-1)');
    return t.join(' ');
  }

  svgShape(prst, w, h, fill, line, isLine, cust, geom) {
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('class', 'pp-svg');
    svg.setAttribute('width', Math.max(w, 1));
    svg.setAttribute('height', Math.max(h, 1));
    svg.setAttribute('viewBox', `0 0 ${Math.max(w, 1)} ${Math.max(h, 1)}`);
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.style.overflow = 'visible';
    let d = '';
    const adj = (name, def) => {
      const g = geom && [...(kid(geom, 'avLst') || { children: [] }).children].find((x) => x.getAttribute('name') === name);
      return g ? +g.getAttribute('fmla').replace(/[^\d-]/g, '') / 100000 : def;
    };
    const m = Math.min(w, h);
    const P = (pts) => 'M' + pts.map((p) => p.join(' ')).join(' L') + ' Z';
    if (isLine) d = prst.startsWith('bentConnector') ? `M0 0 L${w / 2} 0 L${w / 2} ${h} L${w} ${h}` : `M0 0 L${w} ${h}`;
    else if (cust) {
      for (const path of cust.getElementsByTagNameNS(PNS.a, 'path')) {
        const pw = attrN(path, 'w', w * EMU) || 1, ph = attrN(path, 'h', h * EMU) || 1;
        const sx = w / pw, sy = h / ph;
        const pt = (e) => { const p = kid(e, 'pt') || e; return `${attrN(p, 'x', 0) * sx} ${attrN(p, 'y', 0) * sy}`; };
        for (const c of path.children) {
          if (c.localName === 'moveTo') d += 'M' + pt(c) + ' ';
          else if (c.localName === 'lnTo') d += 'L' + pt(c) + ' ';
          else if (c.localName === 'cubicBezTo') d += 'C' + kids(c, 'pt').map((p) => `${attrN(p, 'x', 0) * sx} ${attrN(p, 'y', 0) * sy}`).join(' ') + ' ';
          else if (c.localName === 'quadBezTo') d += 'Q' + kids(c, 'pt').map((p) => `${attrN(p, 'x', 0) * sx} ${attrN(p, 'y', 0) * sy}`).join(' ') + ' ';
          else if (c.localName === 'close') d += 'Z ';
        }
      }
    } else {
      switch (prst) {
        case 'ellipse': d = `M0 ${h / 2} A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2} A${w / 2} ${h / 2} 0 1 0 0 ${h / 2} Z`; break;
        case 'roundRect': { const r = m * adj('adj', 0.16667); d = `M${r} 0 H${w - r} Q${w} 0 ${w} ${r} V${h - r} Q${w} ${h} ${w - r} ${h} H${r} Q0 ${h} 0 ${h - r} V${r} Q0 0 ${r} 0 Z`; break; }
        case 'triangle': { const a = w * adj('adj', 0.5); d = P([[a, 0], [w, h], [0, h]]); break; }
        case 'rtTriangle': d = P([[0, 0], [w, h], [0, h]]); break;
        case 'diamond': d = P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]); break;
        case 'parallelogram': { const a = m * adj('adj', 0.25); d = P([[a, 0], [w, 0], [w - a, h], [0, h]]); break; }
        case 'trapezoid': { const a = m * adj('adj', 0.25); d = P([[0, h], [a, 0], [w - a, 0], [w, h]]); break; }
        case 'pentagon': d = P([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]); break;
        case 'hexagon': { const a = m * adj('adj', 0.25); d = P([[a, 0], [w - a, 0], [w, h / 2], [w - a, h], [a, h], [0, h / 2]]); break; }
        case 'octagon': { const a = m * adj('adj', 0.29289); d = P([[a, 0], [w - a, 0], [w, a], [w, h - a], [w - a, h], [a, h], [0, h - a], [0, a]]); break; }
        case 'homePlate': { const a = m * adj('adj', 0.5); d = P([[0, 0], [w - a, 0], [w, h / 2], [w - a, h], [0, h]]); break; }
        case 'chevron': { const a = m * adj('adj', 0.5); d = P([[0, 0], [w - a, 0], [w, h / 2], [w - a, h], [0, h], [a, h / 2]]); break; }
        case 'rightArrow': { const a = m * adj('adj2', 0.5), t = h * adj('adj1', 0.5); d = P([[0, (h - t) / 2], [w - a, (h - t) / 2], [w - a, 0], [w, h / 2], [w - a, h], [w - a, (h + t) / 2], [0, (h + t) / 2]]); break; }
        case 'leftArrow': { const a = m * adj('adj2', 0.5), t = h * adj('adj1', 0.5); d = P([[w, (h - t) / 2], [a, (h - t) / 2], [a, 0], [0, h / 2], [a, h], [a, (h + t) / 2], [w, (h + t) / 2]]); break; }
        case 'upArrow': { const a = m * adj('adj2', 0.5), t = w * adj('adj1', 0.5); d = P([[(w - t) / 2, h], [(w - t) / 2, a], [0, a], [w / 2, 0], [w, a], [(w + t) / 2, a], [(w + t) / 2, h]]); break; }
        case 'downArrow': { const a = m * adj('adj2', 0.5), t = w * adj('adj1', 0.5); d = P([[(w - t) / 2, 0], [(w - t) / 2, h - a], [0, h - a], [w / 2, h], [w, h - a], [(w + t) / 2, h - a], [(w + t) / 2, 0]]); break; }
        case 'star5': {
          const pts = [];
          for (let i = 0; i < 10; i++) { const ang = -Math.PI / 2 + (i * Math.PI) / 5, rr = i % 2 ? 0.38 : 0.5; pts.push([w / 2 + Math.cos(ang) * w * rr, h / 2 + Math.sin(ang) * h * rr]); }
          d = P(pts);
          break;
        }
        case 'plus': { const a = m * adj('adj', 0.25); d = P([[a, 0], [w - a, 0], [w - a, a], [w, a], [w, h - a], [w - a, h - a], [w - a, h], [a, h], [a, h - a], [0, h - a], [0, a], [a, a]]); break; }
        case 'snip1Rect': { const a = m * adj('adj', 0.16667); d = P([[0, 0], [w - a, 0], [w, a], [w, h], [0, h]]); break; }
        case 'round2SameRect': { const r = m * adj('adj1', 0.16667); d = `M${r} 0 H${w - r} Q${w} 0 ${w} ${r} V${h} H0 V${r} Q0 0 ${r} 0 Z`; break; }
        case 'flowChartTerminator': { const r = h / 2; d = `M${r} 0 H${w - r} A${r} ${r} 0 0 1 ${w - r} ${h} H${r} A${r} ${r} 0 0 1 ${r} 0 Z`; break; }
        case 'flowChartDecision': d = P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]); break;
        case 'can': { const e = h * 0.1; d = `M0 ${e} A${w / 2} ${e} 0 0 1 ${w} ${e} V${h - e} A${w / 2} ${e} 0 0 1 0 ${h - e} Z M0 ${e} A${w / 2} ${e} 0 0 0 ${w} ${e}`; break; }
        default: d = `M0 0 H${w} V${h} H0 Z`;
      }
    }
    const path = document.createElementNS(NS, 'path');
    path.setAttribute('d', d);
    if (fill && !isLine) {
      if (/gradient|url/.test(fill)) {
        // Градиент/картинка в фигуре — через обрезку
        const fo = document.createElementNS(NS, 'foreignObject');
        fo.setAttribute('width', w);
        fo.setAttribute('height', h);
        const id = 'c' + Math.random().toString(36).slice(2);
        const cp = document.createElementNS(NS, 'clipPath');
        cp.id = id;
        const p2 = path.cloneNode();
        cp.append(p2);
        svg.append(cp);
        fo.setAttribute('clip-path', `url(#${id})`);
        const dv = document.createElement('div');
        dv.style.cssText = `width:100%;height:100%;background:${fill}`;
        fo.append(dv);
        svg.append(fo);
        path.setAttribute('fill', 'none');
      } else path.setAttribute('fill', fill);
    } else path.setAttribute('fill', 'none');
    if (line) {
      path.setAttribute('stroke', line.color);
      path.setAttribute('stroke-width', line.w);
      path.setAttribute('vector-effect', 'non-scaling-stroke');
      if (line.dash && line.dash !== 'solid') path.setAttribute('stroke-dasharray', /dot/i.test(line.dash) ? `${line.w} ${line.w * 2}` : `${line.w * 4} ${line.w * 3}`);
      const arrow = (end, at) => {
        if (!end || !end.getAttribute('type') || end.getAttribute('type') === 'none') return;
        const id = 'a' + Math.random().toString(36).slice(2);
        const mk = document.createElementNS(NS, 'marker');
        mk.id = id;
        mk.setAttribute('viewBox', '0 0 10 10');
        mk.setAttribute('refX', '8');
        mk.setAttribute('refY', '5');
        mk.setAttribute('markerWidth', '5');
        mk.setAttribute('markerHeight', '5');
        mk.setAttribute('orient', 'auto-start-reverse');
        const tri = document.createElementNS(NS, 'path');
        tri.setAttribute('d', 'M0 0 L10 5 L0 10 Z');
        tri.setAttribute('fill', line.color);
        mk.append(tri);
        svg.append(mk);
        path.setAttribute(at, `url(#${id})`);
      };
      arrow(line.tail, 'marker-end');
      arrow(line.head, 'marker-start');
    }
    svg.append(path);
    return svg;
  }

  async renderPic(n, shape, ctx, r) {
    const bf = kidAny(n, 'blipFill');
    const url = await this.blipUrl(bf, ctx);
    if (!url) return;
    const img = el('img', 'pp-img');
    img.src = url;
    img.draggable = false;
    const src = kid(bf, 'srcRect');
    if (src) {
      const l = attrN(src, 'l', 0) / 100000, t = attrN(src, 't', 0) / 100000, rr = attrN(src, 'r', 0) / 100000, b = attrN(src, 'b', 0) / 100000;
      const kw = 1 / Math.max(0.01, 1 - l - rr), kh = 1 / Math.max(0.01, 1 - t - b);
      Object.assign(img.style, { width: r.w * kw + 'px', height: r.h * kh + 'px', left: -l * r.w * kw + 'px', top: -t * r.h * kh + 'px' });
      shape.style.overflow = 'hidden';
    }
    const geom = kid(kidAny(n, 'spPr'), 'prstGeom');
    if (geom && geom.getAttribute('prst') === 'ellipse') { shape.style.borderRadius = '50%'; shape.style.overflow = 'hidden'; }
    if (geom && geom.getAttribute('prst') === 'roundRect') { shape.style.borderRadius = Math.min(r.w, r.h) * 0.16667 + 'px'; shape.style.overflow = 'hidden'; }
    const line = this.lineStyle(kidAny(n, 'spPr'), kidAny(n, 'style'), ctx);
    if (line) shape.style.outline = `${line.w}px solid ${line.color}`;
    shape.append(img);
  }

  async renderFrame(n, shape, ctx, sl, own, index, r) {
    const gd = n.getElementsByTagNameNS(PNS.a, 'graphicData')[0];
    const uri = gd ? gd.getAttribute('uri') || '' : '';
    const tbl = gd && kid(gd, 'tbl');
    if (tbl) { this.renderTable(tbl, shape, n, ctx, sl, own, index); return; }
    if (/diagram/.test(uri)) {
      // SmartArt: у PowerPoint есть готовый рисунок из фигур
      const rel = Object.values(sl.rels).find((x) => x.type.endsWith('/diagramDrawing'));
      if (rel && own) {
        const dr = await this.part(this.resolve(rel.target, sl.base));
        const tree = dr.xml.getElementsByTagNameNS(PNS.dsp, 'spTree')[0];
        if (tree) {
          const T = (x, y, w, h) => ({ x, y, w, h });
          for (const k of tree.children) {
            const k2 = k;
            await this.renderNode(k2, shape, { ...ctx, part: dr }, sl, false, index, T, dr);
          }
          return;
        }
      }
    }
    // Диаграммы и прочее — аккуратная заглушка
    const ph = el('div', 'pp-chart');
    ph.innerHTML = ICONS.table + '<span>' + T(/chart/.test(uri) ? 'Диаграмма' : 'Объект') + '</span>';
    shape.append(ph);
  }

  renderTable(tbl, shape, frame, ctx, sl, own, index) {
    const t = el('table', 'pp-tbl');
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => attrN(c, 'w', 0) / EMU);
    const cg = el('colgroup');
    for (const w of cols) { const c = el('col'); c.style.width = w + 'px'; cg.append(c); }
    t.append(cg);
    const tp = kid(tbl, 'tblPr');
    const styled = tp && kid(tp, 'tableStyleId');
    const first = tp && tp.getAttribute('firstRow') === '1', band = tp && tp.getAttribute('bandRow') === '1';
    const accent = this.color({ localName: 'schemeClr', getAttribute: () => 'accent1', children: [] }, ctx) || '#4472C4';
    kids(tbl, 'tr').forEach((tr, ri) => {
      const row = el('tr');
      row.style.height = attrN(tr, 'h', 0) / EMU + 'px';
      kids(tr, 'tc').forEach((tc) => {
        if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') return;
        const td = el('td');
        if (tc.hasAttribute('gridSpan')) td.colSpan = +tc.getAttribute('gridSpan');
        if (tc.hasAttribute('rowSpan')) td.rowSpan = +tc.getAttribute('rowSpan');
        const pr = kid(tc, 'tcPr');
        let bgc = null;
        if (pr) {
          const sf = kid(pr, 'solidFill');
          if (sf) bgc = this.colorIn(sf, ctx);
          for (const [k, css] of [['lnL', 'Left'], ['lnR', 'Right'], ['lnT', 'Top'], ['lnB', 'Bottom']]) {
            const ln = kid(pr, k);
            if (!ln || kid(ln, 'noFill')) continue;
            const c = this.colorIn(kid(ln, 'solidFill'), ctx);
            if (c) td.style['border' + css] = `${Math.max(1, attrN(ln, 'w', 12700) / EMU)}px solid ${c}`;
          }
          td.style.verticalAlign = { ctr: 'middle', b: 'bottom' }[pr.getAttribute('anchor')] || 'top';
          td.style.padding = `${attrN(pr, 'marT', 45720) / EMU}px ${attrN(pr, 'marR', 91440) / EMU}px ${attrN(pr, 'marB', 45720) / EMU}px ${attrN(pr, 'marL', 91440) / EMU}px`;
        } else td.style.padding = '5px 10px';
        let headText = null;
        if (!bgc && styled) {
          if (first && ri === 0) { bgc = accent; headText = '#ffffff'; }
          else if (band) bgc = (first ? ri : ri + 1) % 2 ? 'color-mix(in srgb, ' + accent + ' 22%, white)' : 'color-mix(in srgb, ' + accent + ' 10%, white)';
          if (!td.style.borderBottom) td.style.border = '1px solid #ffffff';
        }
        if (bgc) td.style.background = bgc;
        const tx = kid(tc, 'txBody');
        if (tx) this.renderText(tx, td, frame, { ...ctx, defColor: headText, bold: !!headText }, sl, own, index, null, true);
        row.append(td);
      });
      t.append(row);
    });
    shape.append(t);
  }

  // Текст фигуры: абзацы и фрагменты с унаследованным оформлением.
  renderText(tx, host, sp, ctx, sl, own, index, r, inCell) {
    const ph = this.phOf(sp);
    const m = ctx.m;
    const layoutSp = ph && this.findPh(sl.layout, ph);
    const masterSp = ph && this.findPh(m, ph);
    const kind = !ph ? 'other' : /title/i.test(ph.type) ? 'title' : ph.type === 'body' || ph.type === 'subTitle' || ph.type === 'obj' ? 'body' : 'other';
    const master = m ? (kind === 'title' ? m.txStyles.title : kind === 'body' ? m.txStyles.body : ph ? m.txStyles.other : null) : null;
    const lists = [kid(tx, 'lstStyle'), layoutSp && kid(kidAny(layoutSp, 'txBody'), 'lstStyle'), masterSp && kid(kidAny(masterSp, 'txBody'), 'lstStyle'), master, !ph || !own ? this.defStyle : null].filter(Boolean);
    // Параметры области текста: свои → макет → образец
    const bodyPrs = [kid(tx, 'bodyPr'), layoutSp && kid(kidAny(layoutSp, 'txBody'), 'bodyPr'), masterSp && kid(kidAny(masterSp, 'txBody'), 'bodyPr')].filter(Boolean);
    const bp = (a, d) => { for (const b of bodyPrs) if (b.hasAttribute(a)) return b.getAttribute(a); return d; };
    const fontRef = kidAny(kidAny(sp, 'style'), 'fontRef');
    const styleColor = fontRef ? this.colorIn(fontRef, ctx) : null;
    const auto = bodyPrs.map((b) => kid(b, 'normAutofit')).find(Boolean);
    const scale = auto ? attrN(auto, 'fontScale', 100000) / 100000 : 1;
    const lnRed = auto ? attrN(auto, 'lnSpcReduction', 0) / 100000 : 0;
    const box = el('div', 'pp-text');
    if (!inCell) {
      const ins = (a, d) => +bp(a, d) / EMU;
      Object.assign(box.style, { position: 'absolute', inset: `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px` });
      box.style.justifyContent = { ctr: 'center', b: 'flex-end' }[bp('anchor', kind === 'title' ? 'ctr' : 't')] || 'flex-start';
      if (bp('wrap', 'square') === 'none') box.style.whiteSpace = 'pre';
      const vert = bp('vert', 'horz');
      if (vert === 'vert' || vert === 'eaVert') box.style.writingMode = 'vertical-rl';
      if (vert === 'vert270') { box.style.writingMode = 'vertical-rl'; box.style.transform = 'rotate(180deg)'; }
    }
    const paras = kids(tx, 'p');
    let autoNum = {};
    paras.forEach((p, pi) => {
      const pPrEl = kid(p, 'pPr');
      const lvl = attrN(pPrEl, 'lvl', 0);
      const chain = lists.map((l) => this.lvlFrom(l, lvl, ctx)).reverse();
      const pp = this.merge(...chain, this.parsePPr(pPrEl, ctx));
      const P = el('p', 'pp-p');
      P.dataset.pi = pi;
      const base = { ...(pp.def || {}) };
      if (!base.color) base.color = ctx.defColor || styleColor || null;
      if (ctx.bold && base.b == null) base.b = true;
      if (!base.font) base.font = kind === 'title' ? '+mj-lt' : '+mn-lt';
      const runs = [...p.children].filter((k) => ['r', 'br', 'fld'].includes(k.localName));
      const end = kid(p, 'endParaRPr');
      const firstRPr = runs.length ? kid(runs[0], 'rPr') : end;
      const firstProps = this.merge(base, this.parseRPr(firstRPr, ctx));
      const fpx = (o) => (o.sz || 18) * scale * 4 / 3;
      // Отступы и интервалы
      const sizePx = fpx(firstProps);
      P.style.textAlign = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[pp.algn] || 'left';
      const lh = pp.lnSpc ? (pp.lnSpc.pct ? Math.max(0.5, pp.lnSpc.pct - lnRed) * 1.2 : null) : 1.2 * (1 - lnRed);
      if (pp.lnSpc && pp.lnSpc.pt) P.style.lineHeight = pp.lnSpc.pt * 4 / 3 + 'px';
      else P.style.lineHeight = String(lh);
      const sp2 = (v) => (!v ? 0 : v.pt != null ? v.pt * 4 / 3 : v.pct * sizePx * 1.2);
      if (pi > 0) P.style.marginTop = sp2(pp.spcBef) + 'px';
      P.style.marginBottom = sp2(pp.spcAft) + 'px';
      P.style.paddingLeft = (pp.marL || 0) + 'px';
      P.style.textIndent = (pp.indent || 0) + 'px';
      P.style.fontSize = sizePx + 'px';
      // Маркер
      const hasText = runs.some((k) => k.localName !== 'br' && (k.textContent || '').length);
      if (pp.bu && !pp.bu.none && hasText) {
        let ch = pp.bu.char;
        if (pp.bu.auto) {
          const key = lvl + ':' + pp.bu.auto;
          autoNum[key] = (autoNum[key] || pp.bu.start - 1) + 1;
          for (const k of Object.keys(autoNum)) if (+k.split(':')[0] > lvl) delete autoNum[k];
          ch = numLabel(pp.bu.auto, autoNum[key]);
        }
        if (ch) {
          const b = el('span', 'pp-bu');
          b.contentEditable = 'false';
          b.textContent = ppBullet(ch, pp.buFont);
          b.style.color = pp.buColor || firstProps.color || '#000';
          b.style.fontSize = sizePx * (pp.buSz || 1) + 'px';
          const ind = -(pp.indent || 0);
          if (ind > 0) { b.style.display = 'inline-block'; b.style.minWidth = ind + 'px'; b.style.textIndent = '0'; }
          else b.style.marginRight = '0.4em';
          P.append(b);
        }
      } else autoNum = Object.fromEntries(Object.entries(autoNum).filter(([k]) => +k.split(':')[0] < lvl));
      runs.forEach((rn, ri) => {
        if (rn.localName === 'br') { const br = el('br'); br.dataset.ri = ri; P.append(br); return; }
        const o = this.merge(base, this.parseRPr(kid(rn, 'rPr'), ctx));
        const s = el('span');
        s.dataset.ri = ri;
        let text = (kid(rn, 't') || {}).textContent || '';
        if (rn.localName === 'fld') {
          s.contentEditable = 'false';
          s.dataset.fld = '1';
          const type = rn.getAttribute('type') || '';
          if (/slidenum/i.test(type)) text = String(index + 1);
        }
        s.textContent = text;
        this.runCss(s, o, scale, ctx);
        P.append(s);
      });
      if (!P.childNodes.length || !hasText) {
        const o = this.merge(base, this.parseRPr(end, ctx));
        P.style.fontSize = fpx(o) + 'px';
        if (!P.querySelector('br')) P.append(el('br'));
      }
      box.append(P);
    });
    if (own && !inCell && ph) {
      const hint = { title: 'Заголовок слайда', ctrTitle: 'Заголовок слайда', subTitle: 'Подзаголовок', body: 'Текст слайда', obj: 'Текст слайда' }[ph.type];
      if (hint) box.dataset.hint = T(hint);
    }
    if (own) {
      box.contentEditable = 'true';
      box.spellcheck = false;
      box.txBody = tx;
      box.addEventListener('input', () => { box.changed = true; sl.changed = true; this.markEmpty(box); this.touch(); this.thumbDirty(sl); });
      box.addEventListener('keydown', (e) => this.textKey(e, box));
      box.addEventListener('paste', (e) => {
        // Вставляем только текст, без чужого оформления
        e.preventDefault();
        document.execCommand('insertText', false, e.clipboardData.getData('text/plain'));
      });
      this.markEmpty(box);
    }
    host.append(box);
  }

  runCss(s, o, scale, ctx) {
    const st = s.style;
    st.fontSize = (o.sz || 18) * scale * 4 / 3 + 'px';
    const theme = ctx.m ? ctx.m.theme : { major: 'Calibri Light', minor: 'Calibri' };
    let f = o.font || '+mn-lt';
    if (f === '+mj-lt' || f === '+mj-ea' || f === '+mj-cs') f = theme.major;
    if (f === '+mn-lt' || f === '+mn-ea' || f === '+mn-cs') f = theme.minor;
    st.fontFamily = `"${f}", ${/light/i.test(f) ? '"Segoe UI Light", ' : ''}Calibri, Arial, sans-serif`;
    if (o.b) st.fontWeight = '700';
    if (o.i) st.fontStyle = 'italic';
    const deco = [o.u && 'underline', o.strike && 'line-through'].filter(Boolean).join(' ');
    if (deco) st.textDecoration = deco;
    st.color = o.color || '#000000';
    if (o.baseline) { st.verticalAlign = o.baseline > 0 ? 'super' : 'sub'; st.fontSize = (o.sz || 18) * scale * 4 / 3 * 0.7 + 'px'; }
    if (o.cap === 'all') st.textTransform = 'uppercase';
    if (o.cap === 'small') st.fontVariant = 'small-caps';
    if (o.spc) st.letterSpacing = o.spc * 4 / 3 + 'px';
    if (o.hl) st.background = o.hl;
  }

  markEmpty(box) {
    const empty = !box.textContent.replace(/[​\s]/g, '').length;
    box.classList.toggle('ph-empty', empty && !!box.dataset.hint);
    const p = box.querySelector('p');
    if (p && box.dataset.hint) p.dataset.hint = box.dataset.hint;
  }

  textKey(e, box) {
    if (e.key === 'Escape') { box.blur(); getSelection().removeAllRanges(); this.viewer.focus({ preventScroll: true }); }
    // Tab в списке — уровень ниже/выше (как в PowerPoint)
    if (e.key === 'Tab') {
      e.preventDefault();
      const p = this.curPara(box);
      if (!p) return;
      const lvl = Math.max(0, Math.min(8, (+p.dataset.lvl || this.paraLvl(box, p)) + (e.shiftKey ? -1 : 1)));
      p.dataset.lvl = lvl;
      p.style.paddingLeft = (parseFloat(p.style.paddingLeft) || 0) + (e.shiftKey ? -36 : 36) + 'px';
      if (parseFloat(p.style.paddingLeft) < 0) p.style.paddingLeft = '0px';
      box.changed = true;
      this.touch();
    }
  }
  curPara(box) {
    const s = getSelection();
    if (!s.rangeCount) return null;
    let n = s.anchorNode;
    while (n && n !== box && n.tagName !== 'P') n = n.parentNode;
    return n && n.tagName === 'P' ? n : null;
  }
  paraLvl(box, p) {
    const src = box.txBody && kids(box.txBody, 'p')[+p.dataset.pi];
    return src ? attrN(kid(src, 'pPr'), 'lvl', 0) : 0;
  }

  // ---------- Масштаб и прокрутка ----------
  applyZoom() {
    const z = this.zoom;
    for (const w of this.slideEls || []) {
      w.style.width = this.W * z + 'px';
      w.style.height = this.H * z + 'px';
      w.firstChild.style.transform = `scale(${z})`;
    }
  }
  fitWidth() {
    const avail = this.viewer.clientWidth - 84;
    const availH = this.viewer.clientHeight - 60;
    if (avail <= 0) return;
    void availH;
    this.zoom = clamp(avail / this.W, 0.2, 1.4);
    this.applyZoom();
    this.fitMode = true;
  }
  setZoom(z, quiet) {
    z = clamp(z, 0.2, 3);
    const v = this.viewer;
    const fy = v.scrollTop / Math.max(1, v.scrollHeight);
    this.zoom = z;
    this.fitMode = false;
    this.applyZoom();
    v.scrollTop = fy * v.scrollHeight;
    if (!quiet) this.flashHud(Math.round(z * 100) + '%');
  }
  flashHud(text) {
    this.hud.textContent = text;
    this.hud.classList.add('show');
    clearTimeout(this.hudTimer);
    this.hudTimer = setTimeout(() => this.hud.classList.remove('show'), 900);
  }
  trackCurrent() {
    if (!this.slideEls || !this.slideEls.length) return;
    const mid = this.viewer.scrollTop + this.viewer.clientHeight * 0.4;
    let cur = 0;
    this.slideEls.forEach((w, i) => { if (w.offsetTop <= mid) cur = i; });
    if (cur !== this.curSlide) { this.curSlide = cur; this.markCurThumb(cur); }
  }
  scrollToSlide(i, smooth = true) {
    const w = this.slideEls[i];
    if (!w) return;
    this.viewer.scrollTo({ top: w.offsetTop - 22, behavior: smooth ? 'smooth' : 'auto' });
    this.curSlide = i;
    this.markCurThumb(i);
  }

  // ---------- Миниатюры ----------
  buildThumbs() {
    this.thumbs.textContent = '';
    this.thumbEls = [];
    const k = 150 / this.W;
    this.slides.forEach((sl, i) => {
      const t = el('div', 'thumb pp-thumb');
      t.dataset.i = i;
      const paper = el('div', 't-paper');
      paper.style.width = Math.round(this.W * k) + 'px';
      paper.style.height = Math.round(this.H * k) + 'px';
      const num = el('div', 't-num', String(i + 1));
      t.append(paper, num);
      if (sl.xml.documentElement.getAttribute('show') === '0') t.classList.add('pp-thumb-hidden');
      this.thumbs.append(t);
      this.thumbEls.push(t);
      this.drawThumb(i);
    });
    this.markCurThumb(this.curSlide || 0);
  }
  drawThumb(i) {
    const t = this.thumbEls[i], w = this.slideEls[i];
    if (!t || !w) return;
    const paper = t.firstChild;
    const k = parseFloat(paper.style.width) / this.W;
    const c = w.firstChild.cloneNode(true);
    c.querySelectorAll('[contenteditable]').forEach((n) => n.removeAttribute('contenteditable'));
    c.querySelectorAll('.ph-empty').forEach((n) => n.classList.remove('ph-empty'));
    c.style.transform = `scale(${k})`;
    c.style.transformOrigin = '0 0';
    c.classList.add('pp-thumb-slide');
    paper.replaceChildren(c);
  }
  thumbDirty(sl) {
    const i = this.slides.indexOf(sl);
    clearTimeout(this.thumbTimer);
    this.thumbTimer = setTimeout(() => this.drawThumb(i), 400);
  }
  markCurThumb(i) {
    this.thumbEls.forEach((t, j) => t.classList.toggle('cur', j === i));
    const t = this.thumbEls[i];
    if (t && this.root.classList.contains('side-open')) {
      const tr = this.thumbs.getBoundingClientRect(), r = t.getBoundingClientRect();
      if (r.top < tr.top || r.bottom > tr.bottom) t.scrollIntoView({ block: 'nearest' });
    }
  }

  bindThumbs() {
    const th = this.thumbs;
    th.addEventListener('click', (e) => {
      const t = e.target.closest('.thumb');
      if (t && !this.dragged) this.scrollToSlide(+t.dataset.i);
      this.dragged = false;
    });
    th.addEventListener('contextmenu', (e) => {
      const t = e.target.closest('.thumb');
      if (!t) return;
      e.preventDefault();
      const i = +t.dataset.i;
      this.scrollToSlide(i, false);
      this.slideMenu(e.clientX, e.clientY, i);
    });
    // Перетаскивание миниатюр — новый порядок слайдов
    th.addEventListener('pointerdown', (e) => {
      const t = e.target.closest('.thumb');
      if (!t || e.button !== 0) return;
      const from = +t.dataset.i;
      const sy = e.clientY;
      let started = false, to = from, ghost = null, mark = null;
      const mv = (ev) => {
        if (!started && Math.abs(ev.clientY - sy) < 6) return;
        if (!started) {
          started = true;
          th.setPointerCapture(e.pointerId);
          ghost = t.cloneNode(true);
          ghost.classList.add('pp-ghost');
          document.body.append(ghost);
          mark = el('div', 'pp-drop');
          th.append(mark);
          t.classList.add('pp-dragging');
        }
        ghost.style.left = ev.clientX - 60 + 'px';
        ghost.style.top = ev.clientY - 30 + 'px';
        to = this.thumbEls.length;
        for (let i = 0; i < this.thumbEls.length; i++) {
          const r = this.thumbEls[i].getBoundingClientRect();
          if (ev.clientY < r.top + r.height / 2) { to = i; break; }
        }
        const ref = this.thumbEls[to] || null;
        const tr = th.getBoundingClientRect();
        const y = ref ? ref.getBoundingClientRect().top - 6 : this.thumbEls[this.thumbEls.length - 1].getBoundingClientRect().bottom + 4;
        mark.style.top = y - tr.top + th.scrollTop + 'px';
        if (ev.clientY < tr.top + 30) th.scrollTop -= 12;
        if (ev.clientY > tr.bottom - 30) th.scrollTop += 12;
      };
      const up = () => {
        th.removeEventListener('pointermove', mv);
        th.removeEventListener('pointerup', up);
        th.removeEventListener('pointercancel', up);
        if (!started) return;
        this.dragged = true;
        ghost.remove();
        mark.remove();
        t.classList.remove('pp-dragging');
        const dest = to > from ? to - 1 : to;
        if (dest !== from) this.moveSlide(from, dest);
      };
      th.addEventListener('pointermove', mv);
      th.addEventListener('pointerup', up);
      th.addEventListener('pointercancel', up);
    });
  }

  slideMenu(x, y, i) {
    const items = [
      ['Новый слайд', () => this.newSlideMenu(null, i)],
      ['Дублировать слайд', () => this.duplicateSlide(i), 'Ctrl+D'],
      ['Удалить слайд', () => this.deleteSlide(i), 'Delete'],
      null,
    ];
    if (i > 0) items.push(['Переместить выше', () => this.moveSlide(i, i - 1)]);
    if (i < this.slides.length - 1) items.push(['Переместить ниже', () => this.moveSlide(i, i + 1)]);
    const hidden = this.slides[i].xml.documentElement.getAttribute('show') === '0';
    items.push([hidden ? 'Показывать слайд' : 'Скрыть слайд при показе', () => this.toggleHidden(i)]);
    showMenu(x, y, items);
  }

  // ---------- Операции со слайдами ----------
  snapshot() { return { slides: this.slides.slice(), deleted: this.deleted.slice(), hidden: this.slides.map((s) => s.xml.documentElement.getAttribute('show')) }; }
  pushUndo(label) { this.undoStack.push({ ...this.snapshot(), label }); if (this.undoStack.length > 60) this.undoStack.shift(); }
  async undo() {
    const s = this.undoStack.pop();
    if (!s) return;
    this.slides = s.slides;
    this.deleted = s.deleted;
    this.slides.forEach((sl, i) => { const v = s.hidden[i]; if (v == null) sl.xml.documentElement.removeAttribute('show'); else sl.xml.documentElement.setAttribute('show', v); });
    this.orderChanged = true;
    await this.rebuild(this.curSlide);
    this.touch();
    toast('Отменено: ' + T(s.label));
  }

  async rebuild(focusIdx) {
    const top = this.viewer.scrollTop;
    // Перерисовываем: номера слайдов в полях меняются
    this.syncTextBeforeRebuild();
    await this.renderAll();
    this.viewer.scrollTop = top;
    if (focusIdx != null) this.scrollToSlide(clamp(focusIdx, 0, this.slides.length - 1), false);
  }

  // Правки текста переносим в XML до перерисовки, иначе они пропадут
  syncTextBeforeRebuild() {
    for (const sl of this.slides) if (sl.el) this.flushText(sl);
  }

  moveSlide(from, to) {
    this.pushUndo('Перемещение слайда');
    const [s] = this.slides.splice(from, 1);
    this.slides.splice(to, 0, s);
    this.orderChanged = true;
    this.touch();
    this.rebuild(to);
  }

  async deleteSlide(i) {
    if (this.slides.length <= 1) { toast('В презентации должен остаться хотя бы один слайд'); return; }
    this.pushUndo('Удаление слайда');
    const [s] = this.slides.splice(i, 1);
    this.deleted.push(s);
    this.orderChanged = true;
    this.touch();
    await this.rebuild(Math.min(i, this.slides.length - 1));
  }

  async duplicateSlide(i) {
    this.flushText(this.slides[i]);
    const src = this.slides[i];
    const copy = await this.cloneSlidePart(src);
    this.pushUndo('Копия слайда');
    this.slides.splice(i + 1, 0, copy);
    this.orderChanged = true;
    this.touch();
    await this.rebuild(i + 1);
  }

  // Новая часть слайда в памяти; путь и связи назначим при сохранении.
  async cloneSlidePart(src) {
    const xml = new DOMParser().parseFromString(new XMLSerializer().serializeToString(src.xml), 'application/xml');
    const relsXml = src.relsXml ? new DOMParser().parseFromString(new XMLSerializer().serializeToString(src.relsXml), 'application/xml') : null;
    const p = { path: null, base: src.base, xml, relsXml, rels: {}, isNew: true, changed: true, src };
    for (const r of relsXml ? [...relsXml.documentElement.children] : []) {
      const type = r.getAttribute('Type') || '';
      if (/notesSlide$|comments$|commentAuthors$/.test(type)) { r.remove(); continue; }
      p.rels[r.getAttribute('Id')] = { type, target: r.getAttribute('Target') || '', ext: r.getAttribute('TargetMode') === 'External', el: r };
    }
    p.layout = src.layout;
    return p;
  }

  async newSlideMenu(anchor, at) {
    const sl = this.slides[at != null ? at : this.curSlide || 0];
    const m = sl.layout && sl.layout.master;
    if (!m) return;
    const items = [];
    for (const lid of kids(kid(m.xml.documentElement, 'sldLayoutIdLst', PNS.p), 'sldLayoutId', PNS.p)) {
      const rid = lid.getAttributeNS(PNS.r, 'id') || lid.getAttribute('r:id');
      const rel = m.rels[rid];
      if (!rel) continue;
      const path = this.resolve(rel.target, m.base);
      const lay = await this.part(path);
      lay.master = m;
      const cs = kid(lay.xml.documentElement, 'cSld', PNS.p);
      const name = (cs && cs.getAttribute('name')) || path.split('/').pop();
      items.push([layoutName(name), () => this.addSlideFromLayout(lay, (at != null ? at : this.curSlide || 0) + 1)]);
    }
    if (anchor) this.menuAt(anchor, items); else showMenu(innerWidth / 2 - 110, innerHeight / 2 - items.length * 16, items);
  }

  async addSlideFromLayout(lay, at) {
    const L = lay.xml.documentElement;
    const tree = L.getElementsByTagNameNS(PNS.p, 'spTree')[0];
    let sps = '';
    let id = 2;
    for (const sp of tree ? [...tree.children].filter((k) => k.localName === 'sp') : []) {
      const ph = this.phOf(sp);
      if (!ph || ['dt', 'ftr', 'sldNum'].includes(ph.type)) continue;
      const name = (kidAny(kidAny(sp, 'nvSpPr'), 'cNvPr') || { getAttribute: () => 'Placeholder' }).getAttribute('name');
      sps += `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${escXml(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${ph.type !== 'body' ? ` type="${ph.type}"` : ''}${ph.idx != null ? ` idx="${ph.idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></p:txBody></p:sp>`;
    }
    const xmlText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${PNS.a}" xmlns:r="${PNS.r}" xmlns:p="${PNS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
    const xml = new DOMParser().parseFromString(xmlText, 'application/xml');
    const relsXml = new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}slideLayout" Target="../slideLayouts/${lay.path.split('/').pop()}"/></Relationships>`, 'application/xml');
    const p = { path: null, base: 'ppt/slides/', xml, relsXml, rels: {}, isNew: true, changed: true, layout: lay };
    p.rels.rId1 = { type: PP_REL + 'slideLayout', target: '../slideLayouts/' + lay.path.split('/').pop(), el: relsXml.documentElement.firstElementChild };
    await this.linkSlide(p);
    this.flushText(this.slides[this.curSlide || 0]);
    this.pushUndo('Новый слайд');
    this.slides.splice(at, 0, p);
    this.orderChanged = true;
    this.touch();
    await this.rebuild(at);
    // Курсор — в заголовок нового слайда
    const box = this.slideEls[at] && this.slideEls[at].querySelector('.pp-text[contenteditable]');
    if (box) { box.focus(); const r = document.createRange(); r.selectNodeContents(box.querySelector('p') || box); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); }
  }

  toggleHidden(i) {
    this.pushUndo('Скрытие слайда');
    const R = this.slides[i].xml.documentElement;
    if (R.getAttribute('show') === '0') R.removeAttribute('show'); else R.setAttribute('show', '0');
    this.slides[i].changed = true;
    this.slideEls[i].classList.toggle('pp-hidden', R.getAttribute('show') === '0');
    this.thumbEls[i].classList.toggle('pp-thumb-hidden', R.getAttribute('show') === '0');
    this.touch();
  }

  touch() {
    if (!this.dirty) { this.dirty = true; this.onChange && this.onChange(); }
  }

  // ---------- Текст → XML ----------
  flushText(sl) {
    if (!sl || !sl.el) return;
    for (const box of sl.el.querySelectorAll('.pp-text[contenteditable]')) {
      if (!box.changed || !box.txBody) continue;
      this.writeTxBody(box);
      box.changed = false;
      sl.changed = true;
    }
  }

  writeTxBody(box) {
    const tx = box.txBody;
    const doc = tx.ownerDocument;
    const A = (n) => doc.createElementNS(PNS.a, 'a:' + n);
    const oldPs = kids(tx, 'p');
    let lastSrc = oldPs[0] || null;
    const out = [];
    for (const P of box.querySelectorAll(':scope > p, :scope > div')) {
      const src = P.dataset.pi != null && oldPs[+P.dataset.pi] ? oldPs[+P.dataset.pi] : lastSrc;
      lastSrc = src;
      const np = A('p');
      const srcPPr = src && kid(src, 'pPr');
      if (srcPPr) np.append(srcPPr.cloneNode(true));
      if (P.dataset.lvl != null) {
        let pPr = kid(np, 'pPr');
        if (!pPr) { pPr = A('pPr'); np.prepend(pPr); }
        if (+P.dataset.lvl) pPr.setAttribute('lvl', P.dataset.lvl); else pPr.removeAttribute('lvl');
      }
      const srcRuns = src ? [...src.children].filter((k) => ['r', 'br', 'fld'].includes(k.localName)) : [];
      const baseRPr = () => { const r0 = srcRuns.find((k) => k.localName === 'r'); const e = r0 ? kid(r0, 'rPr') : src && kid(src, 'endParaRPr'); return e; };
      const runs = [];
      const walk = (node, fmt) => {
        for (const c of node.childNodes) {
          if (c.nodeType === 3) {
            if (!c.data) continue;
            runs.push({ text: c.data.replace(/​/g, ''), fmt });
            continue;
          }
          if (c.nodeType !== 1 || c.classList.contains('pp-bu')) continue;
          if (c.tagName === 'BR') { runs.push({ br: true, fmt, ri: c.dataset.ri }); continue; }
          const f = { ...fmt };
          if (c.dataset.ri != null && src && P.dataset.pi != null) f.ri = +c.dataset.ri;
          if (c.dataset.fld) { runs.push({ fld: srcRuns[f.ri], fmt: f }); continue; }
          const tag = c.tagName;
          if (tag === 'B' || tag === 'STRONG') f.b = true;
          if (tag === 'I' || tag === 'EM') f.i = true;
          if (tag === 'U') f.u = true;
          const cs = c.style;
          if (cs.fontWeight) f.b = +cs.fontWeight >= 600 || cs.fontWeight === 'bold';
          if (cs.fontStyle) f.i = cs.fontStyle === 'italic';
          if (cs.textDecorationLine || cs.textDecoration) f.u = /underline/.test(cs.textDecorationLine || cs.textDecoration);
          walk(c, f);
        }
      };
      walk(P, {});
      const rPrFor = (fmt) => {
        const srcRun = fmt.ri != null ? srcRuns[fmt.ri] : null;
        const e = srcRun && srcRun.localName !== 'br' ? kid(srcRun, 'rPr') : baseRPr();
        const r = e ? e.cloneNode(true) : A('rPr');
        if (!e || r.localName !== 'rPr') { const x = A('rPr'); if (e) for (const at of e.attributes) x.setAttribute(at.name, at.value); for (const k of e ? e.children : []) x.append(k.cloneNode(true)); return applyFmt(x, fmt); }
        return applyFmt(r, fmt);
      };
      const applyFmt = (r, fmt) => {
        if (fmt.b != null) { if (fmt.b) r.setAttribute('b', '1'); else r.setAttribute('b', '0'); }
        if (fmt.i != null) { if (fmt.i) r.setAttribute('i', '1'); else r.setAttribute('i', '0'); }
        if (fmt.u != null) { if (fmt.u) r.setAttribute('u', 'sng'); else r.removeAttribute('u'); }
        if (!r.hasAttribute('lang')) r.setAttribute('lang', /[А-Яа-яЁё]/.test(box.textContent) ? 'ru-RU' : 'en-US');
        r.removeAttribute('dirty');
        return r;
      };
      // склеиваем соседние кусочки с одинаковым оформлением
      const merged = [];
      for (const r of runs) {
        const prev = merged[merged.length - 1];
        if (prev && prev.text != null && r.text != null && JSON.stringify(prev.fmt) === JSON.stringify(r.fmt)) prev.text += r.text;
        else merged.push({ ...r });
      }
      // хвостовой <br> от contenteditable — не перенос строки
      if (merged.length && merged[merged.length - 1].br && merged.filter((r) => r.text).length && merged.length > 1) {
        if (P.lastChild && P.lastChild.tagName === 'BR' && !P.lastChild.dataset.ri) merged.pop();
      }
      if (merged.length === 1 && merged[0].br && !merged[0].ri) merged.length = 0;
      for (const r of merged) {
        if (r.fld) { np.append(r.fld.cloneNode(true)); continue; }
        if (r.br) { const b = A('br'); b.append(rPrFor(r.fmt)); np.append(b); continue; }
        if (!r.text) continue;
        const rn = A('r');
        rn.append(rPrFor(r.fmt));
        const t = A('t');
        t.textContent = r.text;
        rn.append(t);
        np.append(rn);
      }
      const end = src && kid(src, 'endParaRPr');
      if (end) np.append(end.cloneNode(true));
      else if (!np.getElementsByTagNameNS(PNS.a, 'r').length) { const e = A('endParaRPr'); e.setAttribute('lang', 'ru-RU'); np.append(e); }
      out.push(np);
    }
    if (!out.length) { const np = A('p'); np.append(A('endParaRPr')); out.push(np); }
    for (const p of oldPs) p.remove();
    for (const p of out) tx.append(p);
  }

  // ---------- Сохранение ----------
  async buildPptx() {
    for (const sl of this.slides) this.flushText(sl);
    const zip = this.zip;
    const ser = (x) => xmlString(x);
    const P = this.pres.xml.documentElement;
    const lst = kid(P, 'sldIdLst', PNS.p);
    let ctTouched = false, presRelsTouched = false;
    const ctOverride = (part, type) => {
      const o = this.ct.createElementNS(PNS.ct, 'Override');
      o.setAttribute('PartName', '/' + part);
      o.setAttribute('ContentType', type);
      this.ct.documentElement.append(o);
      ctTouched = true;
    };
    const ctRemove = (part) => {
      const o = [...this.ct.documentElement.children].find((e) => e.getAttribute('PartName') === '/' + part);
      if (o) { o.remove(); ctTouched = true; }
    };
    // Удалённые: слайд, его связи и заметки
    for (const sl of this.deleted) {
      if (sl.isNew || !sl.path) continue;
      if (this.slides.includes(sl)) continue;
      const nr = this.relOf(sl, 'notesSlide');
      if (nr) { zip.remove(nr.path); zip.remove(this.relsPathOf(nr.path)); ctRemove(nr.path); }
      zip.remove(sl.path);
      zip.remove(this.relsPathOf(sl.path));
      ctRemove(sl.path);
      if (sl.rid && this.pres.rels[sl.rid]) { this.pres.rels[sl.rid].el.remove(); delete this.pres.rels[sl.rid]; presRelsTouched = true; }
      sl.path = null;
    }
    this.deleted = [];
    // Новые
    let maxId = Math.max(255, ...this.slides.filter((s) => s.id).map((s) => s.id));
    for (const sl of this.slides) {
      if (!sl.isNew) continue;
      let n = 1;
      while (zip.file(`ppt/slides/slide${n}.xml`) || this.slides.some((s) => s.path === `ppt/slides/slide${n}.xml`)) n++;
      sl.path = `ppt/slides/slide${n}.xml`;
      sl.base = 'ppt/slides/';
      let rn = 1;
      while (this.pres.rels['rId' + rn]) rn++;
      sl.rid = 'rId' + rn;
      const rel = this.pres.relsXml.createElementNS(PNS.rel, 'Relationship');
      rel.setAttribute('Id', sl.rid);
      rel.setAttribute('Type', PP_REL + 'slide');
      rel.setAttribute('Target', `slides/slide${n}.xml`);
      this.pres.relsXml.documentElement.append(rel);
      this.pres.rels[sl.rid] = { type: PP_REL + 'slide', target: `slides/slide${n}.xml`, el: rel };
      presRelsTouched = true;
      ctOverride(sl.path, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
      sl.id = ++maxId;
      const sid = this.pres.xml.createElementNS(PNS.p, 'p:sldId');
      sid.setAttribute('id', sl.id);
      sid.setAttributeNS(PNS.r, 'r:id', sl.rid);
      sl.sldIdEl = sid;
      this.parts.set(sl.path, sl);
    }
    // Порядок
    if (this.orderChanged || this.slides.some((s) => s.isNew)) {
      for (const s of [...lst.children]) s.remove();
      for (const sl of this.slides) lst.append(sl.sldIdEl);
      this.fixSections();
    }
    for (const sl of this.slides) {
      if (sl.changed || sl.isNew) {
        zip.file(sl.path, ser(sl.xml));
        if (sl.relsXml) zip.file(this.relsPathOf(sl.path), ser(sl.relsXml));
      }
      sl.changed = false;
      sl.isNew = false;
    }
    zip.file(this.pres.path, ser(this.pres.xml));
    if (presRelsTouched) zip.file(this.relsPathOf(this.pres.path), ser(this.pres.relsXml));
    if (ctTouched) zip.file('[Content_Types].xml', ser(this.ct));
    this.orderChanged = false;
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // Разделы (PowerPoint 2010+) перечисляют слайды по номерам — приводим в соответствие.
  fixSections() {
    const secs = [...this.pres.xml.getElementsByTagNameNS(PNS.p14, 'section')];
    if (!secs.length) return;
    const ids = this.slides.map((s) => String(s.id));
    const where = new Map();
    secs.forEach((s, si) => { for (const e of s.getElementsByTagNameNS(PNS.p14, 'sldId')) where.set(e.getAttribute('id'), si); });
    const lists = secs.map((s) => { let l = s.getElementsByTagNameNS(PNS.p14, 'sldIdLst')[0]; if (!l) { l = this.pres.xml.createElementNS(PNS.p14, 'p14:sldIdLst'); s.append(l); } return l; });
    const assign = [];
    let cur = 0;
    ids.forEach((id) => { if (where.has(id)) cur = where.get(id); assign.push([id, cur]); });
    lists.forEach((l) => { while (l.firstChild) l.firstChild.remove(); });
    for (const [id, si] of assign) {
      const e = this.pres.xml.createElementNS(PNS.p14, 'p14:sldId');
      e.setAttribute('id', id);
      lists[si].append(e);
    }
  }

  // ---------- В PDF и показ ----------
  async buildPrintHtml() {
    for (const sl of this.slides) this.flushText(sl);
    const pages = [];
    for (const w of this.slideEls) {
      const c = w.firstChild.cloneNode(true);
      c.style.transform = '';
      c.querySelectorAll('[contenteditable]').forEach((n) => n.removeAttribute('contenteditable'));
      c.querySelectorAll('.ph-empty').forEach((n) => n.classList.remove('ph-empty'));
      for (const img of c.querySelectorAll('img')) if (img.src.startsWith('blob:')) img.src = await readAsDataUrl(await (await fetch(img.src)).blob());
      for (const n of c.querySelectorAll('[style*="blob:"]')) {
        const st = n.getAttribute('style');
        const urls = [...new Set(st.match(/blob:[^")]+/g) || [])];
        let out = st;
        for (const u of urls) out = out.split(u).join(await readAsDataUrl(await (await fetch(u)).blob()));
        n.setAttribute('style', out);
      }
      pages.push(`<div class="pg">${c.outerHTML}</div>`);
    }
    const css = `@page { size: ${this.W}px ${this.H}px; margin: 0; }
html, body { margin: 0; background: #fff; }
.pg { width: ${this.W}px; height: ${this.H}px; overflow: hidden; break-after: page; position: relative; }
.pg:last-child { break-after: auto; }
${PP_CSS}`;
    return `<!doctype html><html><head><meta charset="utf-8"><style>${css}</style></head><body>${pages.join('')}</body></html>`;
  }

  async save(asNew) {
    let target = this.path && /\.pptx$/i.test(this.path) && !asNew ? this.path : null;
    if (!target) {
      const base = (this.path || T('Презентация')).replace(/\.[^.\\/]+$/, '');
      target = await window.folio.saveDialog(base + '.pptx', [
        { name: T('Презентация PowerPoint'), extensions: ['pptx'] },
        { name: 'PDF', extensions: ['pdf'] },
      ]);
      if (!target) return false;
    }
    if (/\.pdf$/i.test(target)) {
      const bytes = await window.folio.htmlToPdf(await this.buildPrintHtml());
      await window.folio.writeFile(target, bytes);
      toast('Сохранено в PDF');
      return false;
    }
    const bytes = await this.buildPptx();
    await window.folio.writeFile(target, bytes);
    this.path = target;
    this.dirty = false;
    this.onChange && this.onChange();
    return true;
  }

  // ---------- Панель ----------
  buildDock() {
    const dock = el('div', 'dock pp-dock');
    dock.addEventListener('pointerdown', (e) => e.preventDefault());
    const btn = (icon, tip, fn, cls = '') => {
      const b = el('button', 'tool ' + cls, icon);
      b.dataset.tip = tip;
      b.addEventListener('click', (e) => fn(e, b));
      dock.append(b);
      return b;
    };
    const sep = () => dock.append(el('div', 'sep'));
    const ui = (this.ui = {});
    ui.pages = btn(ICONS.pages, 'Слайды', () => this.toggleSide());
    ui.pages.classList.add('on');
    sep();
    ui.b = btn(ICONS.bold, 'Жирный · Ctrl+B', () => document.execCommand('bold'));
    ui.i = btn(ICONS.italic, 'Курсив · Ctrl+I', () => document.execCommand('italic'));
    ui.u = btn(ICONS.underline, 'Подчёркнутый · Ctrl+U', () => document.execCommand('underline'));
    sep();
    btn(ICONS.plus, 'Новый слайд', (e, b) => this.newSlideMenu(b, this.curSlide || 0));
    btn(ICONS.dup, 'Дублировать слайд · Ctrl+D', () => this.duplicateSlide(this.curSlide || 0));
    btn(ICONS.trash, 'Удалить слайд', () => this.deleteSlide(this.curSlide || 0));
    sep();
    btn(ICONS.play, 'Показ с этого слайда · F5', () => { if (typeof present === 'function') present(); });
    sep();
    btn(ICONS.hide, 'Скрыть панель · Ctrl+.', () => window.setDockHidden && window.setDockHidden(true));
    const handle = el('button', 'dock-handle');
    handle.dataset.tip = 'Показать панель · Ctrl+.';
    handle.addEventListener('click', () => window.setDockHidden && window.setDockHidden(false));
    this.root.append(dock, handle);
    this.dock = dock;
    this.dockMover = makeDockMovable(this.root, dock, null, 'pptx');
  }

  menuAt(b, items) { DocxEditor.prototype.menuAt.call(this, b, items); }

  syncTools() {
    if (!this.ui || this.root.hidden) return;
    const s = getSelection();
    const inside = s.rangeCount && this.root.contains(s.anchorNode);
    for (const k of ['b', 'i', 'u']) this.ui[k].classList.toggle('on', !!inside && document.queryCommandState({ b: 'bold', i: 'italic', u: 'underline' }[k]));
  }

  toggleSide(force) {
    const open = force != null ? force : !this.root.classList.contains('side-open');
    this.root.classList.toggle('side-open', open);
    this.ui.pages.classList.toggle('on', open);
    setTimeout(() => this.fitMode && this.fitWidth(), 230);
  }

  endEdit() { for (const sl of this.slides || []) this.flushText(sl); }

  onKey(e) {
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && (e.key === '=' || e.key === '+')) { this.setZoom(this.zoom * 1.1); return true; }
    if (ctrl && e.key === '-') { this.setZoom(this.zoom / 1.1); return true; }
    if (ctrl && e.key === '0') { this.fitWidth(); return true; }
    if (ctrl && e.code === 'KeyZ') { this.undo(); return true; }
    if (ctrl && e.code === 'KeyD') { this.duplicateSlide(this.curSlide || 0); return true; }
    if (e.key === 'Delete' && this.thumbs.contains(document.activeElement || null)) { this.deleteSlide(this.curSlide || 0); return true; }
    if (e.key === 'PageDown' || (e.key === 'ArrowDown' && !ctrl)) { this.scrollToSlide(Math.min((this.curSlide || 0) + 1, this.slides.length - 1)); return true; }
    if (e.key === 'PageUp' || (e.key === 'ArrowUp' && !ctrl)) { this.scrollToSlide(Math.max((this.curSlide || 0) - 1, 0)); return true; }
    if (e.key === 'Home') { this.scrollToSlide(0); return true; }
    if (e.key === 'End') { this.scrollToSlide(this.slides.length - 1); return true; }
    return false;
  }

  destroy() {
    document.removeEventListener('selectionchange', this.onSel);
    this.ro.disconnect();
    for (const p of this.media.values()) p.then((u) => URL.revokeObjectURL(u));
    this.root.remove();
  }
}

const escXml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// Номер в нумерованном списке
function numLabel(type, n) {
  const roman = (k) => { let s = ''; for (const [v, r] of ROMAN) while (k >= v) { s += r; k -= v; } return s; };
  const alpha = (k) => { let s = ''; while (k > 0) { k--; s = String.fromCharCode(97 + (k % 26)) + s; k = Math.floor(k / 26); } return s; };
  let core = /roman/i.test(type) ? roman(n) : /alpha/i.test(type) ? alpha(n) : String(n);
  if (/Uc/.test(type)) core = core.toUpperCase();
  if (/ParenBoth/.test(type)) return `(${core})`;
  if (/ParenR/.test(type)) return core + ')';
  if (/Period/.test(type)) return core + '.';
  if (/Plain/.test(type)) return core;
  return core + '.';
}
// Маркеры из символьных шрифтов (Wingdings) — на понятные символы
function ppBullet(ch, font) {
  if (/wingdings/i.test(font || '')) return { 'Ø': '➢', 'ü': '✓', '§': '▪', 'v': '❖', 'q': '❑', 'n': '■', 'l': '●', 'Ü': '➤', 'à': '→' }[ch] || '•';
  if (/symbol/i.test(font || '')) return { '·': '•', 'o': '○' }[ch] || ch;
  return ch;
}
// Названия макетов — по-русски, если стандартные
function layoutName(n) {
  const map = { 'Title Slide': 'Титульный слайд', 'Title and Content': 'Заголовок и объект', 'Section Header': 'Заголовок раздела', 'Two Content': 'Два объекта', Comparison: 'Сравнение', 'Title Only': 'Только заголовок', Blank: 'Пустой слайд', 'Content with Caption': 'Объект с подписью', 'Picture with Caption': 'Рисунок с подписью', 'Title and Vertical Text': 'Заголовок и вертикальный текст', 'Vertical Title and Text': 'Вертикальный заголовок и текст' };
  return map[n] || n;
}

const PP_CSS = `
.pp-slide { position: relative; overflow: hidden; transform-origin: 0 0; background: #fff; }
.pp-sp, .pp-grp { position: absolute; box-sizing: border-box; }
.pp-svg { position: absolute; left: 0; top: 0; overflow: visible; }
.pp-img { position: absolute; left: 0; top: 0; width: 100%; height: 100%; display: block; }
.pp-text { display: flex; flex-direction: column; overflow: visible; word-wrap: break-word; outline: none; }
.pp-text p { margin: 0; min-height: 1em; }
.pp-tbl { border-collapse: collapse; table-layout: fixed; width: 100%; height: 100%; }
.pp-tbl td { overflow: hidden; }
.pp-tbl .pp-text { display: block; }
.pp-chart { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; color: #8a9099; font: 14px "Segoe UI", sans-serif; border: 1px dashed #c4c8ce; border-radius: 6px; background: rgba(240, 242, 245, .6); }
.pp-chart svg { width: 34px; height: 34px; }
`;

(() => {
  const s = document.createElement('style');
  s.textContent = PP_CSS;
  document.head.append(s);
})();

PptxEditor.prototype.setSideRight = PdfEditor.prototype.setSideRight;
PptxEditor.prototype.bindSideGrip = PdfEditor.prototype.bindSideGrip;
PptxEditor.prototype.isShown = PdfEditor.prototype.isShown;

// Пустая презентация 16:9: тема Office, образец и три макета.
async function makeBlankPptx() {
  const zip = new JSZip();
  const A = PNS.a, Pn = PNS.p, R = PNS.r;
  const ns = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${Pn}"`;
  const grp = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  const ph = (id, name, type, idx, x, y, cx, cy, extra = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type ? ` type="${type}"` : ''}${idx != null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr>${x != null ? `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` : ''}</p:spPr><p:txBody><a:bodyPr${extra}/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></p:txBody></p:sp>`;
  const lvl = (i, sz, bu) => `<a:lvl${i}pPr marL="${228600 + (i - 1) * 457200}" indent="-228600" algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>${bu ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : ''}<a:defRPr sz="${sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${i}pPr>`;
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${PNS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${ns} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="ru-RU"/></a:defPPr>${[1, 2, 3].map((i) => `<a:lvl${i}pPr marL="${(i - 1) * 457200}" algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${i}pPr>`).join('')}</p:defaultTextStyle></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${PP_REL}slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="${PP_REL}theme" Target="theme/theme1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${ns}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${grp}${ph(2, 'Title Placeholder 1', 'title', null, 838200, 365125, 10515600, 1325563, ' anchor="ctr"')}${ph(3, 'Text Placeholder 2', 'body', 1, 838200, 1825625, 10515600, 4351338)}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>${lvl(1, 2800, true)}${lvl(2, 2400, true)}${lvl(3, 2000, true)}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="ru-RU"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${PP_REL}slideLayout" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId3" Type="${PP_REL}slideLayout" Target="../slideLayouts/slideLayout3.xml"/><Relationship Id="rId4" Type="${PP_REL}theme" Target="../theme/theme1.xml"/></Relationships>`);
  const layout = (name, type, body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${ns} type="${type}" preserve="1"><p:cSld name="${name}"><p:spTree>${grp}${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  zip.file('ppt/slideLayouts/slideLayout1.xml', layout('Title Slide', 'title', ph(2, 'Title 1', 'ctrTitle', null, 1524000, 1122363, 9144000, 2387600, ' anchor="b"').replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>') + ph(3, 'Subtitle 2', 'subTitle', 1, 1524000, 3602038, 9144000, 1655762).replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>')));
  zip.file('ppt/slideLayouts/slideLayout2.xml', layout('Title and Content', 'obj', ph(2, 'Title 1', 'title', null) + ph(3, 'Content Placeholder 2', null, 1)));
  zip.file('ppt/slideLayouts/slideLayout3.xml', layout('Blank', 'blank', ''));
  for (let i = 1; i <= 3; i++) zip.file(`ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${ns}><p:cSld><p:spTree>${grp}${ph(2, 'Title 1', 'ctrTitle', null)}${ph(3, 'Subtitle 2', 'subTitle', 1)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PNS.rel}"><Relationship Id="rId1" Type="${PP_REL}slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  const sch = (n, v) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>${sch('dk2', '0E2841')}${sch('lt2', 'E8E8E8')}${sch('accent1', '156082')}${sch('accent2', 'E97132')}${sch('accent3', '196B24')}${sch('accent4', '0F9ED5')}${sch('accent5', 'A02B93')}${sch('accent6', '4EA72E')}${sch('hlink', '467886')}${sch('folHlink', '96607D')}</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`);
  return zip.generateAsync({ type: 'uint8array' });
}
