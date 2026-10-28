// Редактор таблиц Excel (.xlsx). Сетка рисуется на холсте — тысячи строк листаются без задержек.
// Как и в Word: нетронутые ячейки сохраняются исходным XML, меняем только то, что поправили.

const XNS = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
};
const XL_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const XL_HDR_W = 46, XL_HDR_H = 22; // шапки строк и столбцов, px
const XL_MDW = 7; // ширина цифры Calibri 11 — так Excel переводит ширину столбца в пиксели
const XL_DEF_COLW = 64, XL_DEF_ROWH = 20;
const xkey = (r, c) => r * 16384 + c;
const XL_INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const XL_FILLS = ['#FFFFFF', '#F2F2F2', '#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#EDE7F6', '#FFFF00'];
const XL_INK = ['#000000', '#595959', '#C00000', '#E36C09', '#00B050', '#0070C0', '#7030A0', '#1F3864'];
const XL_NUMFMTS = [
  ['Общий', 0], ['Число', 4], ['Целое', 3], ['Деньги, ₸', '#,##0.00\\ "₸"'], ['Процент', 10], ['Дата', 14], ['Дата и время', 22], ['Текст', 49],
];
const xlLocale = () => (typeof LANG !== 'undefined' && LANG === 'en' ? { dec: '.', grp: ',' } : { dec: ',', grp: '\u00a0' });

class XlsxEditor {
  constructor(host, { onChange }) {
    this.kind = 'xlsx';
    this.onChange = onChange;
    this.dirty = false;
    this.zoom = 1;
    this.undoStack = [];
    this.redoStack = [];

    this.root = el('div', 'doc xl');
    this.bar = el('div', 'xl-bar');
    this.nameBox = el('input', 'xl-name');
    this.nameBox.spellcheck = false;
    this.nameBox.dataset.tip = 'Адрес ячейки — введите, чтобы перейти';
    this.fx = el('span', 'xl-fx', 'fx');
    this.fBox = el('input', 'xl-formula');
    this.fBox.spellcheck = false;
    this.tabsEl = el('div', 'xl-tabs');
    this.addSheetBtn = el('button', 'xl-addsheet', ICONS.plus);
    this.addSheetBtn.dataset.tip = 'Новый лист';
    this.bar.append(this.nameBox, this.fx, this.fBox, this.tabsEl, this.addSheetBtn);

    this.grid = el('div', 'xl-grid');
    this.viewer = el('div', 'xl-scroll');
    this.viewer.tabIndex = -1;
    this.space = el('div', 'xl-space');
    this.viewer.append(this.space);
    this.canvas = el('canvas', 'xl-canvas');
    this.editBox = el('textarea', 'xl-edit');
    this.editBox.spellcheck = false;
    this.editBox.hidden = true;
    this.grid.append(this.viewer, this.canvas, this.editBox);
    this.hud = el('div', 'hud');
    this.root.append(this.bar, this.grid, this.hud);
    host.append(this.root);

    this.buildDock();
    this.bindGrid();
    this.bindBar();
    this.ro = new ResizeObserver(() => this.draw());
    this.ro.observe(this.grid);
    this.onCopy = (e) => this.handleCopy(e, false);
    this.onCut = (e) => this.handleCopy(e, true);
    this.onPaste = (e) => this.handlePaste(e);
    document.addEventListener('copy', this.onCopy);
    document.addEventListener('cut', this.onCut);
    document.addEventListener('paste', this.onPaste);
  }

  // ---------- Загрузка ----------
  async open(files) {
    const loading = el('div', 'loading', 'Открываю…');
    this.root.append(loading);
    try {
      await this.load(files[0].bytes);
      this.computeMissing();
      this.showSheet(this.activeIdx || 0);
    } finally {
      loading.remove();
    }
  }

  // Формулы без сохранённого результата (так пишут многие программы, кроме Excel) — считаем сразу.
  computeMissing() {
    const list = [];
    for (const sh of this.sheets) if (sh.cells) for (const [k, c] of sh.cells) if (c.f != null && c.v == null) list.push({ sh, r: Math.floor(k / 16384), c: k % 16384 });
    if (!list.length) return;
    this.changedCells = list;
    this.recalc();
    for (const sh of this.sheets) if (sh.cells) this.usedOf(sh);
  }

  async blank() {
    await this.load(await makeBlankXlsx());
    this.showSheet(0);
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

  rels(x) {
    const map = {};
    for (const r of x ? x.documentElement.children : []) map[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: r.getAttribute('Target') || '', el: r };
    return map;
  }

  async load(bytes) {
    this.zip = await JSZip.loadAsync(bytes);
    const rootRels = this.rels(await this.readXml('_rels/.rels'));
    const main = Object.values(rootRels).find((r) => r.type.endsWith('/officeDocument'));
    this.wbPath = this.resolve(main ? main.target : 'xl/workbook.xml', '');
    this.base = this.wbPath.replace(/[^/]*$/, '');
    this.wb = await this.readXml(this.wbPath);
    this.wbRelsPath = this.base + '_rels/' + this.wbPath.split('/').pop() + '.rels';
    this.wbRelsXml = await this.readXml(this.wbRelsPath);
    this.wbRels = this.rels(this.wbRelsXml);
    this.ct = await this.readXml('[Content_Types].xml');
    const relOf = (t) => Object.entries(this.wbRels).find(([, r]) => r.type.endsWith('/' + t));

    // Общие строки
    this.sst = [];
    this.sstIndex = null;
    const sstRel = relOf('sharedStrings');
    this.sstPath = sstRel ? this.resolve(sstRel[1].target, this.base) : null;
    this.sstXml = this.sstPath ? await this.readXml(this.sstPath) : null;
    if (this.sstXml) {
      for (const si of this.sstXml.documentElement.children) {
        if (si.localName !== 'si') continue;
        let s = '';
        for (const t of si.getElementsByTagNameNS(XNS.main, 't')) if (!t.closest || !t.parentNode || t.parentNode.localName !== 'rPh') s += t.textContent;
        this.sst.push(s);
      }
    }

    // Тема (цвета) и стили
    const themeRel = relOf('theme');
    this.themeColors = [];
    if (themeRel) {
      const th = await this.readXml(this.resolve(themeRel[1].target, this.base));
      const cs = th && th.getElementsByTagNameNS(XNS.a, 'clrScheme')[0];
      if (cs) {
        const pick = (n) => { const e = [...cs.children].find((x) => x.localName === n); if (!e || !e.firstElementChild) return '000000'; const c = e.firstElementChild; return c.getAttribute('lastClr') || c.getAttribute('val') || '000000'; };
        this.themeColors = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(pick);
      }
    }
    const stRel = relOf('styles');
    this.stylesPath = stRel ? this.resolve(stRel[1].target, this.base) : null;
    this.stylesXml = this.stylesPath ? await this.readXml(this.stylesPath) : null;
    this.parseStyles();

    // Листы и имена
    this.sheets = [];
    const sheetEls = [...this.wb.getElementsByTagNameNS(XNS.main, 'sheet')];
    for (const se of sheetEls) {
      const rid = se.getAttributeNS(XNS.r, 'id') || se.getAttribute('r:id');
      const rel = this.wbRels[rid];
      const sh = { name: se.getAttribute('name'), el: se, rid, state: se.getAttribute('state') || 'visible' };
      if (!rel || !/worksheet$/.test(rel.type)) { sh.kind = 'other'; this.sheets.push(sh); continue; }
      sh.path = this.resolve(rel.target, this.base);
      sh.xml = await this.readXml(sh.path);
      this.parseSheet(sh);
      this.sheets.push(sh);
    }
    this.origOrder = this.sheets.slice();
    this.names = [...this.wb.getElementsByTagNameNS(XNS.main, 'definedName')].map((d) => ({ el: d, name: d.getAttribute('name'), text: d.textContent }));
    const bv = this.wb.getElementsByTagNameNS(XNS.main, 'workbookView')[0];
    const at = bv ? +bv.getAttribute('activeTab') || 0 : 0;
    this.activeIdx = this.sheets[at] && this.sheets[at].kind !== 'other' && this.sheets[at].state === 'visible' ? at : Math.max(0, this.sheets.findIndex((s) => s.kind !== 'other' && s.state === 'visible'));
    this.formulas = null;
    this.drawTabs();
  }

  // ---------- Стили ----------
  color(e) {
    if (!e) return null;
    if (e.getAttribute('auto') === '1') return null;
    let hex = null;
    if (e.hasAttribute('rgb')) hex = e.getAttribute('rgb').slice(-6);
    else if (e.hasAttribute('theme')) hex = this.themeColors[+e.getAttribute('theme')] || null;
    else if (e.hasAttribute('indexed')) { const i = +e.getAttribute('indexed'); if (i === 64 || i === 65) return null; hex = XL_INDEXED[i] || null; }
    if (!hex) return null;
    const tint = parseFloat(e.getAttribute('tint') || '0');
    if (tint) {
      const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
      hex = ch.map((v) => Math.round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint).toString(16).padStart(2, '0')).join('');
    }
    return '#' + hex;
  }

  parseStyles() {
    const x = this.stylesXml;
    const kids = (p, n) => (p ? [...p.children].filter((e) => e.localName === n) : []);
    const one = (p, n) => (p ? [...p.children].find((e) => e.localName === n) : null);
    const root = x ? x.documentElement : null;
    this.numFmts = {};
    for (const n of kids(one(root, 'numFmts'), 'numFmt')) this.numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode');
    this.fonts = kids(one(root, 'fonts'), 'font').map((f) => this.parseFont(f));
    if (!this.fonts.length) this.fonts = [{ name: 'Calibri', sz: 11 }];
    this.fills = kids(one(root, 'fills'), 'fill').map((f) => {
      const p = one(f, 'patternFill');
      if (!p || !p.getAttribute('patternType') || p.getAttribute('patternType') === 'none') {
        const g = one(f, 'gradientFill');
        if (g) { const st = g.getElementsByTagNameNS(XNS.main, 'color')[0]; return { color: this.color(st) }; }
        return {};
      }
      const t = p.getAttribute('patternType');
      return { color: this.color(one(p, t === 'solid' ? 'fgColor' : 'bgColor') || one(p, 'fgColor')) || (t === 'solid' ? '#000000' : null) };
    });
    this.borders = kids(one(root, 'borders'), 'border').map((b) => {
      const side = (n) => { const s = one(b, n); if (!s || !s.getAttribute('style')) return null; return { style: s.getAttribute('style'), color: this.color(one(s, 'color')) || '#000000' }; };
      return { l: side('left') || side('start'), r: side('right') || side('end'), t: side('top'), b: side('bottom') };
    });
    this.xfs = kids(one(root, 'cellXfs'), 'xf').map((xf) => this.parseXf(xf));
    if (!this.xfs.length) this.xfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, align: {} }];
    this.xfCache = new Map();
    this.xfPatchCache = new Map();
  }

  parseFont(f) {
    const one = (n) => [...f.children].find((e) => e.localName === n);
    const on = (n) => { const e = one(n); return !!e && e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false'; };
    const sz = one('sz'), nm = one('name'), u = one('u');
    return { b: on('b'), i: on('i'), u: !!u && u.getAttribute('val') !== 'none', strike: on('strike'), sz: sz ? parseFloat(sz.getAttribute('val')) : 11, name: nm ? nm.getAttribute('val') : 'Calibri', color: this.color(one('color')) };
  }

  parseXf(xf) {
    const al = [...xf.children].find((e) => e.localName === 'alignment');
    return {
      numFmtId: +(xf.getAttribute('numFmtId') || 0), fontId: +(xf.getAttribute('fontId') || 0), fillId: +(xf.getAttribute('fillId') || 0), borderId: +(xf.getAttribute('borderId') || 0),
      align: al ? { h: al.getAttribute('horizontal'), v: al.getAttribute('vertical'), wrap: al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true', indent: +(al.getAttribute('indent') || 0), rot: +(al.getAttribute('textRotation') || 0) } : {},
    };
  }

  fmtCode(id) {
    return this.numFmts[id] != null ? this.numFmts[id] : BUILTIN_FMT[id] || 'General';
  }

  // Готовый для рисования стиль ячейки (кэш по номеру xf).
  st(s) {
    s = s || 0;
    let c = this.xfCache.get(s);
    if (c) return c;
    const xf = this.xfs[s] || this.xfs[0];
    const f = this.fonts[xf.fontId] || this.fonts[0];
    c = {
      font: f, fill: (this.fills[xf.fillId] || {}).color || null, border: this.borders[xf.borderId] || {},
      fmt: this.fmtCode(xf.numFmtId), align: xf.align || {},
    };
    this.xfCache.set(s, c);
    return c;
  }

  // ---------- Лист ----------
  parseSheet(sh) {
    const x = sh.xml.documentElement;
    const one = (p, n) => (p ? [...p.children].find((e) => e.localName === n) : null);
    sh.cells = new Map();
    sh.rowsInfo = new Map();
    sh.cols = [];
    sh.merges = [];
    sh.ops = [];
    const fmtPr = one(x, 'sheetFormatPr');
    sh.defRowH = fmtPr && fmtPr.getAttribute('defaultRowHeight') ? Math.round(parseFloat(fmtPr.getAttribute('defaultRowHeight')) * 4 / 3) : XL_DEF_ROWH;
    sh.defColW = fmtPr && fmtPr.getAttribute('defaultColWidth') ? Math.round(parseFloat(fmtPr.getAttribute('defaultColWidth')) * XL_MDW) : fmtPr && fmtPr.getAttribute('baseColWidth') ? Math.round((+fmtPr.getAttribute('baseColWidth') + 0.71) * XL_MDW + 5) : XL_DEF_COLW;
    for (const c of [...(one(x, 'cols') || { children: [] }).children]) {
      sh.cols.push({ min: +c.getAttribute('min') - 1, max: +c.getAttribute('max') - 1, w: c.hasAttribute('width') ? Math.round(parseFloat(c.getAttribute('width')) * XL_MDW) : null, hidden: c.getAttribute('hidden') === '1', el: c });
    }
    const sv = one(one(x, 'sheetViews'), 'sheetView');
    sh.grid = !sv || sv.getAttribute('showGridLines') !== '0';
    const pane = one(sv, 'pane');
    sh.frozen = { r: 0, c: 0 };
    if (pane && /frozen/i.test(pane.getAttribute('state') || '')) sh.frozen = { r: +(pane.getAttribute('ySplit') || 0), c: +(pane.getAttribute('xSplit') || 0) };
    const sel = one(sv, 'selection');
    const a = sel && parseAddr((sel.getAttribute('activeCell') || 'A1').split(':')[0]);
    sh.sel = a ? { r: a.r, c: a.c, r1: a.r, c1: a.c, r2: a.r, c2: a.c } : { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    const masters = {};
    let maxR = 0, maxC = 0;
    const data = one(x, 'sheetData');
    let rIdx = -1;
    for (const row of data ? data.children : []) {
      rIdx = row.hasAttribute('r') ? +row.getAttribute('r') - 1 : rIdx + 1;
      const info = { el: row };
      if (row.getAttribute('customHeight') === '1' || row.hasAttribute('ht')) info.h = row.hasAttribute('ht') ? Math.round(parseFloat(row.getAttribute('ht')) * 4 / 3) : null;
      if (row.getAttribute('hidden') === '1') info.hidden = true;
      if (row.hasAttribute('s') && row.getAttribute('customFormat') === '1') info.s = +row.getAttribute('s');
      sh.rowsInfo.set(rIdx, info);
      let cIdx = -1;
      for (const c of row.children) {
        if (c.localName !== 'c') continue;
        const ad = c.getAttribute('r');
        const p = ad ? parseAddr(ad) : null;
        cIdx = p ? p.c : cIdx + 1;
        const cell = { s: +(c.getAttribute('s') || 0), el: c, v: null };
        const t = c.getAttribute('t') || 'n';
        let vEl = null, fEl = null, isEl = null;
        for (const k of c.children) { if (k.localName === 'v') vEl = k; else if (k.localName === 'f') fEl = k; else if (k.localName === 'is') isEl = k; }
        const raw = vEl ? vEl.textContent : null;
        if (t === 's') cell.v = raw != null ? this.sst[+raw] ?? '' : null;
        else if (t === 'b') cell.v = raw === '1' || raw === 'true';
        else if (t === 'e') cell.v = raw != null ? errOf(raw) : null;
        else if (t === 'str') cell.v = raw ?? '';
        else if (t === 'inlineStr') cell.v = isEl ? [...isEl.getElementsByTagNameNS(XNS.main, 't')].map((n) => n.textContent).join('') : '';
        else cell.v = raw != null && raw !== '' ? parseFloat(raw) : null;
        if (fEl) {
          const ft = fEl.getAttribute('t');
          if (ft === 'shared') {
            const si = fEl.getAttribute('si');
            cell.si = si;
            if (fEl.textContent) { masters[si] = { r: rIdx, c: cIdx, f: fEl.textContent }; cell.f = fEl.textContent; }
            else if (masters[si]) cell.f = shiftFormula(masters[si].f, rIdx - masters[si].r, cIdx - masters[si].c);
          } else if (ft !== 'dataTable') cell.f = fEl.textContent;
          if (ft === 'array') cell.arr = true;
        }
        if (cell.v == null && cell.f == null && !cell.s) continue;
        sh.cells.set(xkey(rIdx, cIdx), cell);
        if (cell.v != null || cell.f != null) { if (rIdx > maxR) maxR = rIdx; if (cIdx > maxC) maxC = cIdx; }
      }
    }
    sh.used = { r: maxR, c: maxC };
    for (const m of x.getElementsByTagNameNS(XNS.main, 'mergeCell')) { const g = parseRange(m.getAttribute('ref') || ''); if (g) sh.merges.push(g); }
    this.indexMerges(sh);
    sh.geo = null;
  }

  indexMerges(sh) {
    sh.mergeAt = new Map();
    for (const m of sh.merges) {
      if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) > 200000) continue;
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) sh.mergeAt.set(xkey(r, c), m);
    }
  }

  get sh() { return this.sheets[this.activeIdx]; }
  cell(r, c, sh = this.sh) { return sh.cells.get(xkey(r, c)); }

  usedOf(sh) {
    let r = 0, c = 0;
    for (const [k, v] of sh.cells) if (v.v != null || v.f != null) { const rr = Math.floor(k / 16384), cc = k % 16384; if (rr > r) r = rr; if (cc > c) c = cc; }
    sh.used = { r, c };
    return sh.used;
  }

  // ---------- Геометрия ----------
  colW(c, sh = this.sh) {
    for (const k of sh.cols) if (c >= k.min && c <= k.max) return k.hidden ? 0 : k.w != null ? k.w : sh.defColW;
    return sh.defColW;
  }
  rowH(r, sh = this.sh) {
    const i = sh.rowsInfo.get(r);
    if (i) { if (i.hidden) return 0; if (i.h != null) return i.h; }
    return sh.defRowH;
  }
  // Накопленные координаты; расширяются по мере прокрутки.
  geo(sh = this.sh) {
    if (!sh.geo) sh.geo = { cx: [0], ry: [0] };
    return sh.geo;
  }
  colX(c, sh = this.sh) {
    const g = this.geo(sh);
    while (g.cx.length <= c) g.cx.push(g.cx[g.cx.length - 1] + this.colW(g.cx.length - 1, sh));
    return g.cx[c];
  }
  rowY(r, sh = this.sh) {
    const g = this.geo(sh);
    while (g.ry.length <= r) g.ry.push(g.ry[g.ry.length - 1] + this.rowH(g.ry.length - 1, sh));
    return g.ry[r];
  }
  resetGeo(sh = this.sh) { sh.geo = null; }
  colAt(x, sh = this.sh) {
    let lo = 0, hi = 1;
    while (this.colX(hi, sh) <= x && hi < XL_MAXC - 1) hi *= 2;
    hi = Math.min(hi, XL_MAXC - 1);
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.colX(m, sh) <= x) lo = m; else hi = m - 1; }
    return lo;
  }
  rowAt(y, sh = this.sh) {
    let lo = 0, hi = 1;
    while (this.rowY(hi, sh) <= y && hi < XL_MAXR - 1) hi *= 2;
    hi = Math.min(hi, XL_MAXR - 1);
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.rowY(m, sh) <= y) lo = m; else hi = m - 1; }
    return lo;
  }

  // Экранные координаты ячейки с учётом закреплённых областей и прокрутки.
  view() {
    const z = this.zoom, sh = this.sh;
    const fr = sh.frozen;
    const fw = this.colX(fr.c) * z, fh = this.rowY(fr.r) * z;
    return { z, fr, fw, fh, sl: this.viewer.scrollLeft, st: this.viewer.scrollTop, hx: XL_HDR_W, hy: XL_HDR_H };
  }
  xOf(c, v = this.view()) {
    if (c < v.fr.c) return v.hx + this.colX(c) * v.z;
    return v.hx + v.fw + (this.colX(c) - this.colX(v.fr.c)) * v.z - v.sl;
  }
  yOf(r, v = this.view()) {
    if (r < v.fr.r) return v.hy + this.rowY(r) * v.z;
    return v.hy + v.fh + (this.rowY(r) - this.rowY(v.fr.r)) * v.z - v.st;
  }
  hit(x, y, v = this.view()) {
    let c, r;
    if (x < v.hx) c = -1;
    else if (x < v.hx + v.fw) c = this.colAt((x - v.hx) / v.z);
    else c = this.colAt((x - v.hx - v.fw + v.sl) / v.z + this.colX(v.fr.c));
    if (y < v.hy) r = -1;
    else if (y < v.hy + v.fh) r = this.rowAt((y - v.hy) / v.z);
    else r = this.rowAt((y - v.hy - v.fh + v.st) / v.z + this.rowY(v.fr.r));
    return { r, c };
  }

  updateSpace() {
    const sh = this.sh, z = this.zoom;
    const v = this.view();
    const lastR = this.rowAt((this.grid.clientHeight + v.st) / z);
    const lastC = this.colAt((this.grid.clientWidth + v.sl) / z);
    const R = Math.min(XL_MAXR - 1, Math.max(sh.used.r + 40, lastR + 60, sh.sel.r2 + 20));
    const C = Math.min(XL_MAXC - 1, Math.max(sh.used.c + 12, lastC + 12, sh.sel.c2 + 6));
    this.space.style.width = Math.ceil(XL_HDR_W + this.colX(R > 0 ? C : C) * z) + 'px';
    this.space.style.height = Math.ceil(XL_HDR_H + this.rowY(R) * z) + 'px';
  }

  // ---------- Рисование ----------
  draw() {
    if (!this.sheets || !this.sh || this.root.hidden) return;
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => { this.raf = 0; this.paint(); });
    clearTimeout(this.rafFallback);
    this.rafFallback = setTimeout(() => { if (this.raf) { cancelAnimationFrame(this.raf); this.raf = 0; this.paint(); } }, 120);
  }

  paint() {
    const sh = this.sh;
    if (!sh || !sh.cells) return;
    const W = this.viewer.clientWidth, H = this.viewer.clientHeight;
    if (!W || !H) return;
    this.updateSpace();
    const dpr = window.devicePixelRatio || 1;
    const cv = this.canvas;
    if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) {
      cv.width = Math.round(W * dpr);
      cv.height = Math.round(H * dpr);
      cv.style.width = W + 'px';
      cv.style.height = H + 'px';
    }
    const g = cv.getContext('2d');
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.fillStyle = '#ffffff';
    g.fillRect(0, 0, W, H);
    const v = this.view();
    const fr = sh.frozen;
    const firstC = this.colAt((v.sl) / v.z + this.colX(fr.c)), lastC = this.colAt((v.sl + W) / v.z + this.colX(fr.c));
    const firstR = this.rowAt((v.st) / v.z + this.rowY(fr.r)), lastR = this.rowAt((v.st + H) / v.z + this.rowY(fr.r));
    const cols = [[0, fr.c - 1], [Math.max(fr.c, firstC), lastC]];
    const rows = [[0, fr.r - 1], [Math.max(fr.r, firstR), lastR]];
    const clipX = [[v.hx, v.hx + v.fw], [v.hx + v.fw, W]];
    const clipY = [[v.hy, v.hy + v.fh], [v.hy + v.fh, H]];
    for (let i = 0; i < 2; i++) for (let j = 0; j < 2; j++) {
      const [r0, r1] = rows[i], [c0, c1] = cols[j];
      if (r1 < r0 || c1 < c0) continue;
      g.save();
      g.beginPath();
      g.rect(clipX[j][0], clipY[i][0], clipX[j][1] - clipX[j][0], clipY[i][1] - clipY[i][0]);
      g.clip();
      this.paintCells(g, v, r0, r1, c0, c1);
      g.restore();
    }
    this.paintSelection(g, v, W, H);
    this.paintHeaders(g, v, W, H, rows, cols);
    if (fr.r || fr.c) {
      g.strokeStyle = '#a6a6a6';
      g.lineWidth = 1;
      g.beginPath();
      if (fr.r) { g.moveTo(0, v.hy + v.fh + 0.5); g.lineTo(W, v.hy + v.fh + 0.5); }
      if (fr.c) { g.moveTo(v.hx + v.fw + 0.5, 0); g.lineTo(v.hx + v.fw + 0.5, H); }
      g.stroke();
    }
    this.placeEditBox();
  }

  fontCss(f, z) {
    const px = (f.sz || 11) * 4 / 3 * z;
    return `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${px.toFixed(2)}px "${f.name || 'Calibri'}", Calibri, Carlito, Arial, sans-serif`;
  }

  display(cell) {
    if (!cell || cell.v == null) return null;
    const s = this.st(cell.s);
    const loc = xlLocale();
    return formatValue(cell.v, s.fmt, loc.dec, loc.grp);
  }

  paintCells(g, v, r0, r1, c0, c1) {
    const sh = this.sh, z = v.z;
    const X = [], Wd = [];
    for (let c = c0; c <= c1 + 1; c++) X[c - c0] = this.xOf(c, v);
    for (let c = c0; c <= c1; c++) Wd[c - c0] = X[c - c0 + 1] - X[c - c0];
    const Y = [];
    for (let r = r0; r <= r1 + 1; r++) Y[r - r0] = this.yOf(r, v);
    const x0 = X[0], x1 = X[c1 - c0 + 1], y0 = Y[0], y1 = Y[r1 - r0 + 1];
    // Заливки
    const mergesSeen = new Set();
    for (let r = r0; r <= r1; r++) {
      const rs = (sh.rowsInfo.get(r) || {}).s;
      for (let c = c0; c <= c1; c++) {
        const m = sh.mergeAt.get(xkey(r, c));
        if (m) {
          if (mergesSeen.has(m)) continue;
          mergesSeen.add(m);
          const cell = this.cell(m.r1, m.c1);
          const st = this.st(cell ? cell.s : 0);
          const mx = this.xOf(m.c1, v), my = this.yOf(m.r1, v);
          const mw = this.xOf(m.c2 + 1, v) - mx, mh = this.yOf(m.r2 + 1, v) - my;
          g.fillStyle = st.fill || '#ffffff';
          g.fillRect(mx, my, mw, mh);
          continue;
        }
        const cell = this.cell(r, c);
        const s = cell ? cell.s : rs || this.colStyle(c);
        if (!s) continue;
        const st = this.st(s);
        if (st.fill) { g.fillStyle = st.fill; g.fillRect(X[c - c0], Y[r - r0], Wd[c - c0], Y[r - r0 + 1] - Y[r - r0]); }
      }
    }
    // Сетка
    if (sh.grid) {
      g.strokeStyle = '#e1e3e6';
      g.lineWidth = 1;
      g.beginPath();
      for (let c = c0; c <= c1 + 1; c++) { const x = Math.round(X[c - c0]) + 0.5; g.moveTo(x, y0); g.lineTo(x, y1); }
      for (let r = r0; r <= r1 + 1; r++) { const y = Math.round(Y[r - r0]) + 0.5; g.moveTo(x0, y); g.lineTo(x1, y); }
      g.stroke();
      // Внутри объединённых ячеек линий нет
      for (const m of mergesSeen) {
        const cell = this.cell(m.r1, m.c1);
        const st = this.st(cell ? cell.s : 0);
        const mx = this.xOf(m.c1, v), my = this.yOf(m.r1, v);
        g.fillStyle = st.fill || '#ffffff';
        g.fillRect(Math.round(mx) + 1, Math.round(my) + 1, this.xOf(m.c2 + 1, v) - mx - 1, this.yOf(m.r2 + 1, v) - my - 1);
      }
    }
    // Рамки
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const cell = this.cell(r, c);
      if (!cell || !cell.s) continue;
      const b = this.st(cell.s).border;
      if (!b.l && !b.r && !b.t && !b.b) continue;
      const m = sh.mergeAt.get(xkey(r, c));
      const xa = Math.round(X[c - c0]), xb = Math.round(m ? this.xOf(m.c2 + 1, v) : X[c - c0 + 1]);
      const ya = Math.round(Y[r - r0]), yb = Math.round(m ? this.yOf(m.r2 + 1, v) : Y[r - r0 + 1]);
      this.line(g, b.t, xa, ya, xb, ya);
      this.line(g, b.b, xa, yb, xb, yb);
      this.line(g, b.l, xa, ya, xa, yb);
      this.line(g, b.r, xb, ya, xb, yb);
    }
    // Текст
    const textDone = new Set();
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const m = sh.mergeAt.get(xkey(r, c));
        if (m) { if (textDone.has(m)) continue; textDone.add(m); }
        const tr = m ? m.r1 : r, tc = m ? m.c1 : c;
        const cell = this.cell(tr, tc);
        if (!cell || cell.v == null || cell.v === '') continue;
        const bx = this.xOf(tc, v), by = this.yOf(tr, v);
        const bw = (m ? this.xOf(m.c2 + 1, v) : this.xOf(tc + 1, v)) - bx;
        const bh = (m ? this.yOf(m.r2 + 1, v) : this.yOf(tr + 1, v)) - by;
        if (bw <= 0 || bh <= 0) continue;
        this.paintText(g, v, cell, tr, tc, bx, by, bw, bh, !m, c1);
      }
    }
  }

  colStyle(c) {
    for (const k of this.sh.cols) if (c >= k.min && c <= k.max) return +(k.el && k.el.getAttribute('style')) || 0;
    return 0;
  }

  line(g, side, x1, y1, x2, y2) {
    if (!side) return;
    const w = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 }[side.style] || 1;
    g.strokeStyle = side.color || '#000';
    g.lineWidth = w;
    g.setLineDash(/dot/i.test(side.style) && !/dash/i.test(side.style) ? [1, 2] : /dash/i.test(side.style) ? [4, 2] : []);
    const o = w % 2 ? 0.5 : 0;
    g.beginPath();
    g.moveTo(x1 + (x1 === x2 ? o : 0), y1 + (y1 === y2 ? o : 0));
    g.lineTo(x2 + (x1 === x2 ? o : 0), y2 + (y1 === y2 ? o : 0));
    g.stroke();
    g.setLineDash([]);
  }

  paintText(g, v, cell, r, c, bx, by, bw, bh, canOverflow, lastC) {
    const st = this.st(cell.s);
    const d = this.display(cell);
    if (!d) return;
    const z = v.z;
    g.font = this.fontCss(st.font, z);
    const al = st.align;
    let h = al.h;
    if (!h || h === 'general') h = d.num || typeof cell.v === 'number' ? 'right' : d.bool || d.err ? 'center' : 'left';
    if (h === 'centerContinuous' || h === 'distributed') h = 'center';
    if (h === 'fill' || h === 'justify') h = 'left';
    const pad = 3 * z + (al.indent || 0) * 9 * z;
    const fpx = (st.font.sz || 11) * 4 / 3 * z;
    const lh = fpx * 1.2;
    let lines = [d.text];
    if (al.wrap || /\n/.test(d.text)) lines = this.wrapLines(g, d.text, al.wrap ? bw - pad * 2 : Infinity);
    let text = d.text;
    let clipL = bx, clipR = bx + bw;
    const tw = lines.length === 1 ? g.measureText(text).width : 0;
    if (lines.length === 1 && tw > bw - pad * 2) {
      if (d.num) { text = '#'.repeat(Math.max(1, Math.floor((bw - pad) / Math.max(1, g.measureText('#').width)))); lines = [text]; }
      else if (canOverflow && !al.wrap) {
        // Текст переливается в пустые соседние ячейки
        const sh = this.sh;
        if (h !== 'right') {
          let cc = c + 1, x = bx + bw;
          while (x < bx + tw + pad * 2 && cc <= lastC + 30) {
            const n = this.cell(r, cc);
            if ((n && n.v != null && n.v !== '') || sh.mergeAt.get(xkey(r, cc))) break;
            x = this.xOf(cc + 1, v);
            cc++;
          }
          clipR = x;
        }
        if (h !== 'left') {
          let cc = c - 1, x = bx;
          while (x > bx + bw - tw - pad * 2 && cc >= 0) {
            const n = this.cell(r, cc);
            if ((n && n.v != null && n.v !== '') || sh.mergeAt.get(xkey(r, cc))) break;
            x = this.xOf(cc, v);
            cc--;
          }
          clipL = x;
        }
      }
    }
    g.save();
    g.beginPath();
    g.rect(clipL + 1, by + 1, clipR - clipL - 1, bh - 1);
    g.clip();
    g.fillStyle = d.color || st.font.color || '#000000';
    g.textBaseline = 'alphabetic';
    const total = lines.length * lh;
    const va = al.v || 'bottom';
    let ty = va === 'top' ? by + 2 * z : va === 'center' ? by + (bh - total) / 2 : by + bh - total - 2 * z;
    for (const ln of lines) {
      const w = g.measureText(ln).width;
      const tx = h === 'right' ? bx + bw - pad - w : h === 'center' ? bx + (bw - w) / 2 : bx + pad;
      const base = ty + lh * 0.5 + fpx * 0.36;
      g.fillText(ln, tx, base);
      if (st.font.u || st.font.strike) {
        g.fillRect(tx, st.font.u ? base + 1.5 * z : base - fpx * 0.3, w, Math.max(1, z));
      }
      ty += lh;
    }
    g.restore();
  }

  wrapLines(g, text, maxW) {
    const out = [];
    for (const para of String(text).split('\n')) {
      if (maxW === Infinity) { out.push(para); continue; }
      const words = para.split(/(\s+)/);
      let cur = '';
      for (const w of words) {
        const t = cur + w;
        if (cur && g.measureText(t).width > maxW && w.trim()) { out.push(cur.trimEnd()); cur = w.trimStart(); }
        else cur = t;
      }
      out.push(cur);
    }
    return out;
  }

  paintSelection(g, v, W, H) {
    const s = this.sh.sel;
    const ac = getComputedStyle(this.root).getPropertyValue('--xl-sel').trim() || '#1f7a4a';
    const box = (g1) => {
      const x = this.xOf(g1.c1, v), y = this.yOf(g1.r1, v);
      return { x, y, w: this.xOf(g1.c2 + 1, v) - x, h: this.yOf(g1.r2 + 1, v) - y };
    };
    const b = box(s);
    g.save();
    g.beginPath();
    g.rect(v.hx, v.hy, W - v.hx, H - v.hy);
    g.clip();
    if (s.r1 !== s.r2 || s.c1 !== s.c2) {
      const m = this.sh.mergeAt.get(xkey(s.r, s.c));
      const a = box(m || { r1: s.r, c1: s.c, r2: s.r, c2: s.c });
      const p = new Path2D();
      p.rect(b.x, b.y, b.w, b.h);
      p.rect(a.x, a.y, a.w, a.h);
      g.fillStyle = 'rgba(31, 122, 74, .1)';
      g.fill(p, 'evenodd');
    }
    if (this.cutRange && this.cutRange.sheet === this.sh) {
      const q = box(this.cutRange);
      g.strokeStyle = ac;
      g.lineWidth = 1.5;
      g.setLineDash([5, 3]);
      g.strokeRect(q.x + 1, q.y + 1, q.w - 2, q.h - 2);
      g.setLineDash([]);
    }
    // Ссылки формулы, которую сейчас набирают, — цветными рамками
    if (this.editing && this.refHints) {
      this.refHints.forEach((h, i) => {
        if (h.sheet && h.sheet.toLowerCase() !== this.sh.name.toLowerCase()) return;
        const q = box(h);
        const col = XL_REF_COLORS[i % XL_REF_COLORS.length];
        g.strokeStyle = col;
        g.lineWidth = 1.5;
        g.setLineDash([4, 2]);
        g.strokeRect(q.x + 1, q.y + 1, q.w - 2, q.h - 2);
        g.setLineDash([]);
      });
    }
    g.strokeStyle = ac;
    g.lineWidth = 2;
    g.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1);
    if (!this.editing) {
      g.fillStyle = ac;
      g.fillRect(b.x + b.w - 4, b.y + b.h - 4, 6, 6);
      g.fillStyle = '#ffffff';
      g.fillRect(b.x + b.w - 4, b.y + b.h - 5, 6, 1);
      g.fillRect(b.x + b.w - 5, b.y + b.h - 4, 1, 6);
    }
    g.restore();
  }

  paintHeaders(g, v, W, H, rows, cols) {
    const cs = getComputedStyle(this.root);
    const bg = cs.getPropertyValue('--tab').trim() || '#f8f8f9';
    const ink = cs.getPropertyValue('--ink-2').trim() || '#636a73';
    const line = cs.getPropertyValue('--xl-hline').trim() || 'rgba(0,0,0,.12)';
    const on = cs.getPropertyValue('--xl-hsel').trim() || 'rgba(31,122,74,.16)';
    const onInk = cs.getPropertyValue('--xl-sel').trim() || '#1f7a4a';
    const s = this.sh.sel;
    g.fillStyle = bg;
    g.fillRect(0, 0, W, v.hy);
    g.fillRect(0, 0, v.hx, H);
    g.font = `${(11.5).toFixed(1)}px "Segoe UI Variable Text", "Segoe UI", sans-serif`;
    g.textBaseline = 'middle';
    g.textAlign = 'center';
    const all = s.r1 === 0 && s.r2 >= XL_MAXR - 1, allC = s.c1 === 0 && s.c2 >= XL_MAXC - 1;
    // Столбцы
    g.save();
    g.beginPath();
    g.rect(v.hx, 0, W - v.hx, v.hy);
    g.clip();
    for (const [c0, c1] of cols) for (let c = c0; c <= c1; c++) {
      const x = this.xOf(c, v), w = this.xOf(c + 1, v) - x;
      if (w <= 0) continue;
      if (c >= s.c1 && c <= s.c2) { g.fillStyle = on; g.fillRect(x, 0, w, v.hy); }
      if (c >= s.c1 && c <= s.c2 && all) { g.fillStyle = onInk; g.fillRect(x, v.hy - 2, w, 2); }
      g.fillStyle = c >= s.c1 && c <= s.c2 ? onInk : ink;
      if (w > 14) g.fillText(colName(c), x + w / 2, v.hy / 2 + 0.5);
      g.fillStyle = line;
      g.fillRect(Math.round(x + w) - 0.5, 4, 1, v.hy - 8);
    }
    g.restore();
    // Строки
    g.save();
    g.beginPath();
    g.rect(0, v.hy, v.hx, H - v.hy);
    g.clip();
    for (const [r0, r1] of rows) for (let r = r0; r <= r1; r++) {
      const y = this.yOf(r, v), h = this.yOf(r + 1, v) - y;
      if (h <= 0) continue;
      if (r >= s.r1 && r <= s.r2) { g.fillStyle = on; g.fillRect(0, y, v.hx, h); }
      if (r >= s.r1 && r <= s.r2 && allC) { g.fillStyle = onInk; g.fillRect(v.hx - 2, y, 2, h); }
      g.fillStyle = r >= s.r1 && r <= s.r2 ? onInk : ink;
      if (h > 9) g.fillText(String(r + 1), v.hx / 2, y + h / 2 + 0.5);
      g.fillStyle = line;
      g.fillRect(6, Math.round(y + h) - 0.5, v.hx - 12, 1);
    }
    g.restore();
    g.textAlign = 'start';
    g.fillStyle = line;
    g.fillRect(0, v.hy - 0.5, W, 1);
    g.fillRect(v.hx - 0.5, 0, 1, H);
  }

  // ---------- Выделение ----------
  select(r, c, extend) {
    const sh = this.sh, s = sh.sel;
    r = clamp(r, 0, XL_MAXR - 1);
    c = clamp(c, 0, XL_MAXC - 1);
    if (extend) {
      s.r1 = Math.min(s.r, r); s.r2 = Math.max(s.r, r);
      s.c1 = Math.min(s.c, c); s.c2 = Math.max(s.c, c);
      this.growToMerges(s);
    } else {
      const m = sh.mergeAt.get(xkey(r, c));
      if (m) { r = m.r1; c = m.c1; }
      Object.assign(s, { r, c, r1: r, c1: c, r2: m ? m.r2 : r, c2: m ? m.c2 : c });
    }
    this.syncBar();
    this.draw();
  }

  growToMerges(s) {
    let grown = true;
    while (grown) {
      grown = false;
      for (const m of this.sh.merges) {
        if (m.r2 < s.r1 || m.r1 > s.r2 || m.c2 < s.c1 || m.c1 > s.c2) continue;
        if (m.r1 < s.r1) { s.r1 = m.r1; grown = true; }
        if (m.r2 > s.r2) { s.r2 = m.r2; grown = true; }
        if (m.c1 < s.c1) { s.c1 = m.c1; grown = true; }
        if (m.c2 > s.c2) { s.c2 = m.c2; grown = true; }
      }
    }
  }

  // Прокрутить так, чтобы ячейка была видна.
  reveal(r, c) {
    const v = this.view(), sh = this.sh;
    const W = this.viewer.clientWidth, H = this.viewer.clientHeight;
    if (c >= sh.frozen.c) {
      const x = this.xOf(c, v), x2 = this.xOf(c + 1, v);
      if (x < v.hx + v.fw) this.viewer.scrollLeft -= v.hx + v.fw - x;
      else if (x2 > W) this.viewer.scrollLeft += Math.min(x2 - W + 4, x - v.hx - v.fw);
    }
    if (r >= sh.frozen.r) {
      const y = this.yOf(r, v), y2 = this.yOf(r + 1, v);
      if (y < v.hy + v.fh) this.viewer.scrollTop -= v.hy + v.fh - y;
      else if (y2 > H) { this.updateSpaceFor(r); this.viewer.scrollTop += Math.min(y2 - H + 4, y - v.hy - v.fh); }
    }
  }
  updateSpaceFor(r) {
    const cur = parseFloat(this.space.style.height) || 0;
    const need = XL_HDR_H + this.rowY(r + 60) * this.zoom;
    if (need > cur) this.space.style.height = need + 'px';
  }

  move(dr, dc, extend, jump) {
    const s = this.sh.sel;
    let r = extend ? (s.r === s.r1 ? s.r2 : s.r1) : s.r, c = extend ? (s.c === s.c1 ? s.c2 : s.c1) : s.c;
    if (!extend) { const m = this.sh.mergeAt.get(xkey(s.r, s.c)); if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; } }
    if (jump) [r, c] = this.edge(r, c, dr, dc);
    else { r += dr; c += dc; }
    // пропускаем скрытые
    while (dr && r > 0 && r < XL_MAXR - 1 && this.rowH(r) === 0) r += Math.sign(dr);
    while (dc && c > 0 && c < XL_MAXC - 1 && this.colW(c) === 0) c += Math.sign(dc);
    r = clamp(r, 0, XL_MAXR - 1);
    c = clamp(c, 0, XL_MAXC - 1);
    if (extend) {
      const sel = this.sh.sel;
      sel.r1 = Math.min(sel.r, r); sel.r2 = Math.max(sel.r, r);
      sel.c1 = Math.min(sel.c, c); sel.c2 = Math.max(sel.c, c);
      this.growToMerges(sel);
      this.syncBar();
      this.reveal(r, c);
      this.draw();
    } else {
      this.select(r, c);
      this.reveal(this.sh.sel.r, this.sh.sel.c);
    }
  }

  // Ctrl+стрелка: к краю блока с данными.
  edge(r, c, dr, dc) {
    const has = (rr, cc) => { const x = this.cell(rr, cc); return !!x && x.v != null && x.v !== ''; };
    const lim = this.usedOf(this.sh);
    const inb = (rr, cc) => rr >= 0 && cc >= 0 && rr < XL_MAXR && cc < XL_MAXC;
    let nr = r + dr, nc = c + dc;
    if (!inb(nr, nc)) return [r, c];
    if (has(r, c) && has(nr, nc)) {
      while (inb(nr + dr, nc + dc) && has(nr + dr, nc + dc)) { nr += dr; nc += dc; }
      return [nr, nc];
    }
    while (inb(nr, nc) && !has(nr, nc)) {
      if ((dr > 0 && nr > lim.r) || (dc > 0 && nc > lim.c)) return [dr ? XL_MAXR - 1 : r, dc ? XL_MAXC - 1 : c].map((x, i) => (i === 0 ? (dr ? Math.min(x, lim.r + 1000) : x) : dc ? Math.min(x, lim.c + 100) : x));
      nr += dr; nc += dc;
    }
    if (!inb(nr, nc)) return [clamp(nr, 0, XL_MAXR - 1), clamp(nc, 0, XL_MAXC - 1)];
    return [nr, nc];
  }

  selRangeStr() {
    const s = this.sh.sel;
    if (s.r1 === s.r2 && s.c1 === s.c2) return addr(s.r, s.c);
    const m = this.sh.mergeAt.get(xkey(s.r1, s.c1));
    if (m && m.r1 === s.r1 && m.c1 === s.c1 && m.r2 === s.r2 && m.c2 === s.c2) return addr(s.r, s.c);
    if (s.r1 === 0 && s.r2 >= XL_MAXR - 1) return colName(s.c1) + ':' + colName(s.c2);
    if (s.c1 === 0 && s.c2 >= XL_MAXC - 1) return s.r1 + 1 + ':' + (s.r2 + 1);
    return rangeStr(s);
  }

  syncBar() {
    if (!this.sh) return;
    if (document.activeElement !== this.nameBox) this.nameBox.value = this.selRangeStr();
    if (!this.editing) {
      const s = this.sh.sel, cell = this.cell(s.r, s.c);
      this.fBox.value = cell ? (cell.f != null ? '=' + cell.f : editText(cell.v, this.st(cell.s).fmt, xlLocale().dec)) : '';
    }
    this.syncTools();
    this.updateStatus();
  }

  // Сумма и среднее выделенного — тихо, в подсказке у панели.
  updateStatus() {
    const s = this.sh.sel;
    if (s.r1 === s.r2 && s.c1 === s.c2) { this.status.hidden = true; return; }
    let sum = 0, n = 0, cnt = 0;
    const lim = this.sh.used;
    for (const [k, cell] of this.sh.cells) {
      const r = Math.floor(k / 16384), c = k % 16384;
      if (r < s.r1 || r > Math.min(s.r2, lim.r) || c < s.c1 || c > Math.min(s.c2, lim.c)) continue;
      if (cell.v == null || cell.v === '') continue;
      cnt++;
      if (typeof cell.v === 'number') { sum += cell.v; n++; }
    }
    if (cnt < 2) { this.status.hidden = true; return; }
    const loc = xlLocale();
    const f = (x) => formatValue(+x.toPrecision(12), '#,##0.##########', loc.dec, loc.grp).text.replace(/[.,]$/, '');
    this.status.hidden = false;
    this.status.textContent = n ? `${T('Сумма')}: ${f(sum)}   ${T('Среднее')}: ${f(sum / n)}   ${T('Количество')}: ${cnt}` : `${T('Количество')}: ${cnt}`;
  }

  // ---------- Мышь ----------
  bindGrid() {
    const vw = this.viewer;
    vw.addEventListener('scroll', () => { this.draw(); });
    vw.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    const local = (e) => { const r = this.grid.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
    // Граница столбца/строки в шапке — потянуть размер
    const edgeAt = (p) => {
      const v = this.view();
      if (p.y < v.hy && p.x > v.hx) {
        const h = this.hit(p.x, v.hy + 1, v);
        const xr = this.xOf(h.c + 1, v), xl = this.xOf(h.c, v);
        if (Math.abs(p.x - xr) <= 4) return { col: h.c };
        if (Math.abs(p.x - xl) <= 4 && h.c > 0) { let c = h.c - 1; while (c > 0 && this.colW(c) === 0) c--; return { col: c }; }
      }
      if (p.x < v.hx && p.y > v.hy) {
        const h = this.hit(v.hx + 1, p.y, v);
        const yb = this.yOf(h.r + 1, v), yt = this.yOf(h.r, v);
        if (Math.abs(p.y - yb) <= 3) return { row: h.r };
        if (Math.abs(p.y - yt) <= 3 && h.r > 0) { let r = h.r - 1; while (r > 0 && this.rowH(r) === 0) r--; return { row: r }; }
      }
      return null;
    };
    vw.addEventListener('pointermove', (e) => {
      if (e.buttons) return;
      const p = local(e);
      const ed = edgeAt(p);
      vw.style.cursor = ed ? (ed.col != null ? 'col-resize' : 'row-resize') : p.y < XL_HDR_H && p.x > XL_HDR_W ? 'default' : p.x < XL_HDR_W && p.y > XL_HDR_H ? 'default' : 'cell';
    });
    vw.addEventListener('dblclick', (e) => {
      const p = local(e);
      const ed = edgeAt(p);
      if (ed && ed.col != null) { this.autoFit(ed.col); return; }
      if (ed) return;
      const v = this.view();
      if (p.x < v.hx || p.y < v.hy) return;
      this.startEdit(false);
    });
    vw.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const p = local(e);
      const v = this.view();
      const h = this.hit(p.x, p.y, v);
      const s = this.sh.sel;
      const inside = (h.r < 0 || (h.r >= s.r1 && h.r <= s.r2)) && (h.c < 0 || (h.c >= s.c1 && h.c <= s.c2));
      if (!inside) {
        if (h.r < 0 && h.c >= 0) this.selectCols(h.c, h.c);
        else if (h.c < 0 && h.r >= 0) this.selectRows(h.r, h.r);
        else if (h.r >= 0 && h.c >= 0) this.select(h.r, h.c);
      }
      this.cellMenu(e.clientX, e.clientY, h.r < 0 ? 'cols' : h.c < 0 ? 'rows' : 'cells');
    });
    vw.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      // По полосам прокрутки не выделяем
      if (e.offsetX > vw.clientWidth || e.offsetY > vw.clientHeight) return;
      const p = local(e);
      const v = this.view();
      const ed = edgeAt(p);
      if (ed) { e.preventDefault(); this.dragResize(e, ed); return; }
      const h = this.hit(p.x, p.y, v);
      // Набираем формулу — щелчок по ячейке вставляет ссылку на неё
      if (this.editing && this.canPoint()) { e.preventDefault(); this.pointRef(e, h); return; }
      if (this.editing) this.commitEdit();
      vw.focus({ preventScroll: true });
      if (h.r < 0 && h.c < 0) { this.selectAll(); return; }
      if (h.r < 0) { this.dragHeader(e, 'c', h.c); return; }
      if (h.c < 0) { this.dragHeader(e, 'r', h.r); return; }
      // Маркер заполнения в правом нижнем углу
      const s = this.sh.sel;
      const bx = this.xOf(s.c2 + 1, v), by = this.yOf(s.r2 + 1, v);
      if (Math.abs(p.x - bx) <= 5 && Math.abs(p.y - by) <= 5) { e.preventDefault(); this.dragFill(e); return; }
      this.select(h.r, h.c, e.shiftKey);
      vw.setPointerCapture(e.pointerId);
      let last = null;
      const mv = (ev) => {
        const q = local(ev);
        const k = this.hit(clamp(q.x, XL_HDR_W + 1, this.grid.clientWidth - 2), clamp(q.y, XL_HDR_H + 1, this.grid.clientHeight - 2));
        this.autoScroll(q);
        if (last && last.r === k.r && last.c === k.c) return;
        last = k;
        const sel = this.sh.sel;
        sel.r1 = Math.min(sel.r, k.r); sel.r2 = Math.max(sel.r, k.r);
        sel.c1 = Math.min(sel.c, k.c); sel.c2 = Math.max(sel.c, k.c);
        this.growToMerges(sel);
        this.syncBar();
        this.draw();
      };
      const up = () => {
        vw.removeEventListener('pointermove', mv);
        vw.removeEventListener('pointerup', up);
        clearInterval(this.scrollTimer);
      };
      vw.addEventListener('pointermove', mv);
      vw.addEventListener('pointerup', up);
    });
  }

  autoScroll(q) {
    const W = this.viewer.clientWidth, H = this.viewer.clientHeight;
    const dx = q.x > W - 10 ? 24 : q.x < XL_HDR_W ? -24 : 0;
    const dy = q.y > H - 10 ? 24 : q.y < XL_HDR_H ? -24 : 0;
    if (dx) this.viewer.scrollLeft += dx;
    if (dy) { if (dy > 0) this.updateSpaceFor(this.sh.sel.r2 + 2); this.viewer.scrollTop += dy; }
  }

  selectAll() {
    Object.assign(this.sh.sel, { r1: 0, c1: 0, r2: XL_MAXR - 1, c2: XL_MAXC - 1 });
    this.syncBar();
    this.draw();
  }
  selectCols(a, b) {
    Object.assign(this.sh.sel, { r: 0, c: a, r1: 0, r2: XL_MAXR - 1, c1: Math.min(a, b), c2: Math.max(a, b) });
    this.syncBar();
    this.draw();
  }
  selectRows(a, b) {
    Object.assign(this.sh.sel, { r: a, c: 0, c1: 0, c2: XL_MAXC - 1, r1: Math.min(a, b), r2: Math.max(a, b) });
    this.syncBar();
    this.draw();
  }

  dragHeader(e, axis, start) {
    const vw = this.viewer;
    if (axis === 'c') this.selectCols(start, start); else this.selectRows(start, start);
    vw.setPointerCapture(e.pointerId);
    const local = (ev) => { const r = this.grid.getBoundingClientRect(); return { x: ev.clientX - r.left, y: ev.clientY - r.top }; };
    const mv = (ev) => {
      const q = local(ev);
      this.autoScroll(q);
      const h = this.hit(Math.max(q.x, XL_HDR_W + 1), Math.max(q.y, XL_HDR_H + 1));
      if (axis === 'c') { this.sh.sel.c1 = Math.min(start, h.c); this.sh.sel.c2 = Math.max(start, h.c); }
      else { this.sh.sel.r1 = Math.min(start, h.r); this.sh.sel.r2 = Math.max(start, h.r); }
      this.syncBar();
      this.draw();
    };
    const up = () => { vw.removeEventListener('pointermove', mv); vw.removeEventListener('pointerup', up); };
    vw.addEventListener('pointermove', mv);
    vw.addEventListener('pointerup', up);
  }

  dragResize(e, ed) {
    const vw = this.viewer;
    vw.setPointerCapture(e.pointerId);
    const z = this.zoom;
    const s = this.sh.sel;
    const start = ed.col != null ? this.colW(ed.col) : this.rowH(ed.row);
    const x0 = e.clientX, y0 = e.clientY;
    // Если выделено несколько столбцов и тянем один из них — меняем все
    const many = ed.col != null ? (s.r1 === 0 && s.r2 >= XL_MAXR - 1 && ed.col >= s.c1 && ed.col <= s.c2 ? [s.c1, s.c2] : [ed.col, ed.col])
      : (s.c1 === 0 && s.c2 >= XL_MAXC - 1 && ed.row >= s.r1 && ed.row <= s.r2 ? [s.r1, s.r2] : [ed.row, ed.row]);
    this.beginOp();
    const snap = this.snapshot();
    let val = start;
    const mv = (ev) => {
      val = Math.max(0, Math.round(start + (ed.col != null ? ev.clientX - x0 : ev.clientY - y0) / z));
      if (ed.col != null) this.setColWidth(many[0], many[1], val, true);
      else this.setRowHeight(many[0], many[1], val, true);
      this.flashHud(ed.col != null ? `${T('Ширина')}: ${val} px` : `${T('Высота')}: ${val} px`);
      this.draw();
    };
    const up = () => {
      vw.removeEventListener('pointermove', mv);
      vw.removeEventListener('pointerup', up);
      if (val !== start) { this.pushUndo({ snap, after: this.snapshot() }); this.touch(); }
      this.op = null;
    };
    vw.addEventListener('pointermove', mv);
    vw.addEventListener('pointerup', up);
  }

  setColWidth(c1, c2, px, quiet) {
    const sh = this.sh;
    this.splitCols(sh, c1, c2);
    for (const k of sh.cols) if (k.min >= c1 && k.max <= c2) { k.w = px; k.hidden = px === 0; k.mod = true; }
    if (!sh.cols.some((k) => k.min === c1)) sh.cols.push({ min: c1, max: c2, w: px, hidden: px === 0, mod: true });
    sh.cols.sort((a, b) => a.min - b.min);
    sh.colsMod = true;
    this.resetGeo(sh);
    if (!quiet) { this.touch(); this.draw(); }
  }

  // Разрезаем диапазоны <col>, чтобы [c1..c2] шли отдельными записями.
  splitCols(sh, c1, c2) {
    const out = [];
    for (const k of sh.cols) {
      if (k.max < c1 || k.min > c2) { out.push(k); continue; }
      if (k.min < c1) out.push({ ...k, max: c1 - 1 });
      out.push({ ...k, min: Math.max(k.min, c1), max: Math.min(k.max, c2) });
      if (k.max > c2) out.push({ ...k, min: c2 + 1 });
    }
    // Пустые места внутри [c1..c2] — новые записи
    for (let c = c1; c <= c2; c++) {
      if (!out.some((k) => c >= k.min && c <= k.max)) {
        let e = c;
        while (e + 1 <= c2 && !out.some((k) => e + 1 >= k.min && e + 1 <= k.max)) e++;
        out.push({ min: c, max: e, w: null, hidden: false });
        c = e;
      }
    }
    sh.cols = out.sort((a, b) => a.min - b.min);
  }

  setRowHeight(r1, r2, px, quiet) {
    const sh = this.sh;
    for (let r = r1; r <= r2; r++) {
      const i = { ...(sh.rowsInfo.get(r) || {}) };
      i.h = px;
      i.hidden = px === 0;
      i.custom = true;
      sh.rowsInfo.set(r, i);
    }
    sh.touched = true;
    this.resetGeo(sh);
    if (!quiet) { this.touch(); this.draw(); }
  }

  // Как Excel: число или дата не влезает в столбец обычной ширины — столбец расширяется сам.
  widenFor(r, c) {
    const sh = this.sh, cell = this.cell(r, c);
    if (!cell || typeof cell.v !== 'number' || sh.mergeAt.get(xkey(r, c))) return;
    if (sh.cols.some((k) => c >= k.min && c <= k.max && (k.w != null || k.hidden))) return;
    const g = this.canvas.getContext('2d');
    g.font = this.fontCss(this.st(cell.s).font, 1);
    const need = Math.ceil(g.measureText(this.display(cell).text).width + 10);
    if (need > this.colW(c)) { this.setColWidth(c, c, need, true); this.draw(); }
  }

  autoFit(c) {
    const sh = this.sh;
    const g = this.canvas.getContext('2d');
    let w = 0;
    for (const [k, cell] of sh.cells) {
      if (k % 16384 !== c || cell.v == null) continue;
      if (sh.mergeAt.get(k)) continue;
      const d = this.display(cell);
      g.font = this.fontCss(this.st(cell.s).font, 1);
      w = Math.max(w, g.measureText(d.text).width);
    }
    if (!w) return;
    const s = this.sh.sel;
    const [a, b] = s.r1 === 0 && s.r2 >= XL_MAXR - 1 && c >= s.c1 && c <= s.c2 ? [s.c1, s.c2] : [c, c];
    this.beginOp();
    const snap = this.snapshot();
    if (a === b) this.setColWidth(c, c, Math.ceil(w + 10), true);
    else for (let k = a; k <= b; k++) { this.sh.sel.c1 = k; this.autoFitOne(k, g); }
    this.sh.sel.c1 = a;
    this.pushUndo({ snap, after: this.snapshot() });
    this.op = null;
    this.touch();
    this.draw();
  }
  autoFitOne(c, g) {
    let w = 0;
    for (const [k, cell] of this.sh.cells) {
      if (k % 16384 !== c || cell.v == null) continue;
      g.font = this.fontCss(this.st(cell.s).font, 1);
      w = Math.max(w, g.measureText(this.display(cell).text).width);
    }
    if (w) this.setColWidth(c, c, Math.ceil(w + 10), true);
  }

  // Маркер заполнения: тянем вниз/вправо — копируем ячейки (числа 1, 2 → продолжаем ряд).
  dragFill(e) {
    const vw = this.viewer;
    vw.setPointerCapture(e.pointerId);
    const s = { ...this.sh.sel };
    const local = (ev) => { const r = this.grid.getBoundingClientRect(); return { x: ev.clientX - r.left, y: ev.clientY - r.top }; };
    let target = null;
    const mv = (ev) => {
      const q = local(ev);
      this.autoScroll(q);
      const h = this.hit(clamp(q.x, XL_HDR_W + 1, this.grid.clientWidth - 2), clamp(q.y, XL_HDR_H + 1, this.grid.clientHeight - 2));
      const down = h.r - s.r2, right = h.c - s.c2;
      const sel = this.sh.sel;
      Object.assign(sel, s);
      target = null;
      if (down > 0 && down >= right) { sel.r2 = h.r; target = { axis: 'r', to: h.r }; }
      else if (right > 0) { sel.c2 = h.c; target = { axis: 'c', to: h.c }; }
      this.draw();
    };
    const up = () => {
      vw.removeEventListener('pointermove', mv);
      vw.removeEventListener('pointerup', up);
      if (target) this.fill(s, target);
    };
    vw.addEventListener('pointermove', mv);
    vw.addEventListener('pointerup', up);
  }

  fill(src, target) {
    this.beginOp();
    const vertical = target.axis === 'r';
    const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
    const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
    const end = target.to - (vertical ? src.r1 : src.c1);
    for (let l = 0; l < lines; l++) {
      const at = (i) => (vertical ? [src.r1 + i, src.c1 + l] : [src.r1 + l, src.c1 + i]);
      const seq = [];
      for (let i = 0; i < len; i++) seq.push(this.cell(...at(i)));
      // Числовой ряд: шаг по последним двум значениям
      const numsOnly = seq.every((x) => x && typeof x.v === 'number' && x.f == null);
      const step = numsOnly ? (len > 1 ? seq[len - 1].v - seq[len - 2].v : 0) : 0;
      for (let i = len; i <= end; i++) {
        const k = i % len, srcCell = seq[k];
        const [r, c] = at(i);
        if (!srcCell) { this.setCell(r, c, null); continue; }
        const nc = { s: srcCell.s, v: srcCell.v };
        if (srcCell.f != null) {
          const [sr, scc] = at(k);
          nc.f = shiftFormula(srcCell.f, r - sr, c - scc);
        } else if (numsOnly && len > 1) nc.v = +(seq[len - 1].v + step * (i - len + 1)).toPrecision(15);
        else if (numsOnly && len === 1) nc.v = srcCell.v;
        this.setCell(r, c, nc);
      }
    }
    this.commitOp();
  }

  // ---------- Строка формул ----------
  bindBar() {
    this.nameBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const t = this.nameBox.value.trim().replace(/\$/g, '');
        const g = parseRange(t);
        if (g) {
          Object.assign(this.sh.sel, { r: g.r1, c: g.c1, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 });
          this.reveal(g.r1, g.c1);
          this.syncBar();
          this.draw();
          this.viewer.focus({ preventScroll: true });
        } else toast('Такой ячейки нет — пример: B12');
      } else if (e.key === 'Escape') { this.syncBar(); this.viewer.focus(); }
    });
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.nameBox.addEventListener('blur', () => this.syncBar());
    this.fBox.addEventListener('focus', () => {
      if (!this.editing) this.startEdit(false, this.fBox.value, true);
    });
    this.fBox.addEventListener('input', () => {
      if (!this.editing) this.startEdit(false, this.fBox.value, true);
      this.editBox.value = this.fBox.value;
      this.afterInput();
    });
    this.fBox.addEventListener('keydown', (e) => this.editKey(e, this.fBox));
    this.editBox.addEventListener('input', () => {
      this.fBox.value = this.editBox.value;
      this.afterInput();
    });
    this.editBox.addEventListener('keydown', (e) => this.editKey(e, this.editBox));
    this.editBox.addEventListener('blur', () => {
      setTimeout(() => {
        if (this.editing && document.activeElement !== this.editBox && document.activeElement !== this.fBox && !this.pointing) this.commitEdit();
      }, 0);
    });
    this.fBox.addEventListener('blur', () => {
      setTimeout(() => {
        if (this.editing && document.activeElement !== this.editBox && document.activeElement !== this.fBox && !this.pointing) this.commitEdit();
      }, 0);
    });
    this.addSheetBtn.addEventListener('click', () => this.addSheet());
  }

  // ---------- Правка ячейки ----------
  startEdit(quick, initial, fromBar) {
    const s = this.sh.sel;
    const cell = this.cell(s.r, s.c);
    this.editing = { r: s.r, c: s.c, quick: !!quick, sheet: this.activeIdx };
    const text = initial != null ? initial : cell ? (cell.f != null ? '=' + cell.f : editText(cell.v, this.st(cell.s).fmt, xlLocale().dec)) : '';
    this.editBox.value = text;
    this.fBox.value = text;
    this.editBox.hidden = false;
    this.placeEditBox();
    if (!fromBar) {
      this.editBox.focus({ preventScroll: true });
      const n = this.editBox.value.length;
      this.editBox.setSelectionRange(n, n);
    }
    this.afterInput();
    this.draw();
  }

  placeEditBox() {
    if (!this.editing || this.editing.sheet !== this.activeIdx) { if (!this.editing) this.editBox.hidden = true; return; }
    const v = this.view();
    const { r, c } = this.editing;
    const m = this.sh.mergeAt.get(xkey(r, c));
    const x = this.xOf(c, v), y = this.yOf(r, v);
    const w = (m ? this.xOf(m.c2 + 1, v) : this.xOf(c + 1, v)) - x, h = (m ? this.yOf(m.r2 + 1, v) : this.yOf(r + 1, v)) - y;
    const cell = this.cell(r, c);
    const st = this.st(cell ? cell.s : 0);
    const eb = this.editBox;
    eb.style.font = this.fontCss(st.font, this.zoom);
    eb.style.left = x + 'px';
    eb.style.top = y + 'px';
    eb.style.minWidth = w + 1 + 'px';
    eb.style.height = 'auto';
    eb.style.width = w + 1 + 'px';
    const g = this.canvas.getContext('2d');
    g.font = eb.style.font;
    const lines = eb.value.split('\n');
    const tw = Math.max(...lines.map((l) => g.measureText(l).width)) + 14;
    eb.style.width = Math.min(Math.max(w + 1, tw), this.grid.clientWidth - x - 4) + 'px';
    eb.style.height = Math.max(h + 1, lines.length * (st.font.sz || 11) * 1.6 * this.zoom + 4) + 'px';
    eb.style.color = st.font.color || '#000';
    eb.style.background = st.fill || '#fff';
    eb.style.textAlign = st.align.h === 'center' ? 'center' : st.align.h === 'right' ? 'right' : 'left';
  }

  afterInput() {
    this.placeEditBox();
    // Подсветка ссылок формулы
    const t = this.editBox.value;
    this.refHints = null;
    if (t.startsWith('=')) {
      try {
        const hints = [];
        for (const tok of tokenize(t.slice(1).replace(/;/g, ','))) {
          if (tok.t !== 'ref') continue;
          const n = parseRefText(tok.v);
          const a = n.a, b = n.t === 'range' ? n.b : n.a;
          if (a.r == null || a.c == null || b.r == null || b.c == null) continue;
          hints.push({ sheet: n.sheet, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) });
        }
        this.refHints = hints;
      } catch {}
    }
    this.draw();
  }

  // Можно ли щелчком по ячейке вставить ссылку (курсор после «=», «(», «;», оператора).
  canPoint() {
    const box = document.activeElement === this.fBox ? this.fBox : this.editBox;
    const t = box.value;
    if (!t.startsWith('=')) return false;
    const pos = box.selectionStart;
    if (this.lastPoint && this.lastPoint.end === pos && t.slice(this.lastPoint.start, this.lastPoint.end) === this.lastPoint.text) return true;
    return /[=(,;+\-*/^&<>:]\s*$/.test(t.slice(0, pos));
  }

  pointRef(e, h) {
    if (h.r < 0 || h.c < 0) return;
    const box = document.activeElement === this.fBox ? this.fBox : this.editBox;
    this.pointing = true;
    const start = this.lastPoint && this.lastPoint.end === box.selectionStart ? this.lastPoint.start : box.selectionStart;
    const sheetPrefix = this.editing.sheet !== this.activeIdx ? quoteSheet(this.sh.name) + '!' : '';
    const put = (r1, c1, r2, c2) => {
      const g = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
      const ref = sheetPrefix + rangeStr(g);
      const t = box.value;
      const endOld = this.lastPoint && this.lastPoint.start === start ? this.lastPoint.end : box.selectionEnd;
      box.value = t.slice(0, start) + ref + t.slice(endOld);
      box.setSelectionRange(start + ref.length, start + ref.length);
      this.lastPoint = { start, end: start + ref.length, text: ref };
      (box === this.fBox ? this.editBox : this.fBox).value = box.value;
      this.afterInput();
    };
    put(h.r, h.c, h.r, h.c);
    const vw = this.viewer;
    vw.setPointerCapture(e.pointerId);
    const local = (ev) => { const r = this.grid.getBoundingClientRect(); return { x: ev.clientX - r.left, y: ev.clientY - r.top }; };
    const mv = (ev) => {
      const q = local(ev);
      const k = this.hit(Math.max(q.x, XL_HDR_W + 1), Math.max(q.y, XL_HDR_H + 1));
      put(h.r, h.c, k.r, k.c);
    };
    const up = () => {
      vw.removeEventListener('pointermove', mv);
      vw.removeEventListener('pointerup', up);
      this.pointing = false;
      box.focus({ preventScroll: true });
    };
    vw.addEventListener('pointermove', mv);
    vw.addEventListener('pointerup', up);
  }

  editKey(e, box) {
    if (!this.editing && box === this.fBox) {
      if (e.key === 'Escape') { this.syncBar(); this.viewer.focus(); }
      return;
    }
    const k = e.key;
    if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      const p = box.selectionStart;
      box.value = box.value.slice(0, p) + '\n' + box.value.slice(box.selectionEnd);
      box.setSelectionRange(p + 1, p + 1);
      (box === this.fBox ? this.editBox : this.fBox).value = box.value;
      this.afterInput();
      return;
    }
    if (k === 'Enter' || k === 'Tab') {
      e.preventDefault();
      if (this.commitEdit()) {
        if (k === 'Enter') this.move(e.shiftKey ? -1 : 1, 0);
        else this.move(0, e.shiftKey ? -1 : 1);
      }
      return;
    }
    if (k === 'Escape') { e.preventDefault(); this.cancelEdit(); return; }
    if (this.editing.quick && box === this.editBox && /^Arrow/.test(k) && !this.canPoint()) {
      e.preventDefault();
      if (this.commitEdit()) this.move(k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0);
      return;
    }
    if (/^Arrow/.test(k) && this.canPoint() && this.editing.quick) {
      // Стрелками выбираем ячейку для ссылки, как в Excel
      e.preventDefault();
      const base = this.pointCursor || { r: this.editing.r, c: this.editing.c };
      const nr = clamp(base.r + (k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0), 0, XL_MAXR - 1);
      const nc = clamp(base.c + (k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0), 0, XL_MAXC - 1);
      this.pointCursor = { r: nr, c: nc };
      const start = this.lastPoint && this.lastPoint.end === box.selectionStart ? this.lastPoint.start : box.selectionStart;
      const ref = addr(nr, nc);
      const endOld = this.lastPoint && this.lastPoint.start === start ? this.lastPoint.end : box.selectionEnd;
      box.value = box.value.slice(0, start) + ref + box.value.slice(endOld);
      box.setSelectionRange(start + ref.length, start + ref.length);
      this.lastPoint = { start, end: start + ref.length, text: ref };
      this.fBox.value = box.value;
      this.afterInput();
      return;
    }
    this.lastPoint = null;
    this.pointCursor = null;
  }

  cancelEdit() {
    if (!this.editing) return;
    const back = this.editing.sheet;
    this.editing = null;
    this.refHints = null;
    this.lastPoint = null;
    this.editBox.hidden = true;
    if (back !== this.activeIdx) this.showSheet(back);
    this.syncBar();
    this.draw();
    this.viewer.focus({ preventScroll: true });
  }

  // Разбор набранного: число, дата, процент, логическое, формула или текст.
  parseInput(text, cell) {
    const loc = xlLocale();
    const curFmt = this.st(cell ? cell.s : 0).fmt;
    if (text === '') return { v: null };
    if (text.startsWith("'")) return { v: text.slice(1), text: true };
    if (text.startsWith('=') && text.length > 1) {
      try { return { f: normalizeInput(text.slice(1)) }; }
      catch { return { bad: true }; }
    }
    const t = text.trim();
    const up = t.toUpperCase();
    if (up === 'TRUE' || up === 'ИСТИНА') return { v: true };
    if (up === 'FALSE' || up === 'ЛОЖЬ') return { v: false };
    let s = t.replace(/[\s\u00a0]/g, '');
    let pct = false, money = false;
    if (/%$/.test(s)) { pct = true; s = s.slice(0, -1); }
    if (/^-?[\d.,]+\s*(₸|тг|руб|р\.|\$|€)$/i.test(t.replace(/[\s\u00a0]/g, ''))) { money = true; s = s.replace(/(₸|тг|руб|р\.|\$|€)$/i, ''); }
    let num = null;
    if (/^[-+]?\d+([.,]\d+)?([eE][-+]?\d+)?$/.test(s)) num = parseFloat(s.replace(',', '.'));
    else if (loc.dec === ',' && /^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) num = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    else if (loc.dec === '.' && /^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) num = parseFloat(s.replace(/,/g, ''));
    if (num != null && isFinite(num)) {
      if (pct) return { v: num / 100, fmt: /%/.test(curFmt) ? null : Number.isInteger(num) ? 9 : 10 };
      if (money) return { v: num, fmt: /₸|\$|€|р/.test(curFmt) ? null : '#,##0.00\\ "₸"' };
      return { v: num };
    }
    const d = parseDateText(t);
    if (d != null) return { v: d, fmt: fmtIsDate(curFmt) ? null : /\d:\d/.test(t) ? 22 : 14 };
    if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)) {
      const [h, m, sec] = t.split(':').map(Number);
      return { v: (h * 3600 + m * 60 + (sec || 0)) / 86400, fmt: fmtIsDate(curFmt) ? null : 20 };
    }
    return { v: text };
  }

  commitEdit() {
    if (!this.editing) return true;
    const { r, c, sheet } = this.editing;
    const sh = this.sheets[sheet];
    const text = this.editBox.value;
    const old = this.cell(r, c, sh);
    const p = this.parseInput(text, old);
    if (p.bad) {
      toast('В формуле ошибка — проверьте скобки и кавычки');
      this.editBox.focus();
      return false;
    }
    this.editing = null;
    this.refHints = null;
    this.lastPoint = null;
    this.pointCursor = null;
    this.editBox.hidden = true;
    if (sheet !== this.activeIdx) this.showSheet(sheet);
    const before = old ? (old.f != null ? '=' + old.f : editText(old.v, this.st(old.s).fmt, xlLocale().dec)) : '';
    if (text !== before) {
      this.beginOp();
      let s = old ? old.s : (sh.rowsInfo.get(r) || {}).s || this.colStyle(c);
      if (p.fmt != null) s = this.ensureXf(s, { numFmt: p.fmt });
      const nc = p.f != null ? { s, f: p.f, v: old && old.f === p.f ? old.v : null } : { s, v: p.v };
      if (p.f != null) { try { if (!formulaSupported(parseFormula(p.f))) toast('Эту формулу Folio пока не считает — Excel посчитает её при открытии'); } catch {} }
      if (p.v == null && p.f == null && !s) this.setCell(r, c, null);
      else this.setCell(r, c, nc);
      this.commitOp();
      this.widenFor(r, c);
    }
    this.syncBar();
    this.draw();
    this.viewer.focus({ preventScroll: true });
    return true;
  }

  // ---------- Изменения, пересчёт, отмена ----------
  beginOp() {
    if (!this.op) this.op = { cells: new Map(), sheet: this.activeIdx };
  }

  setCell(r, c, cell, sh = this.sh) {
    const k = xkey(r, c);
    const old = sh.cells.get(k) || null;
    if (this.op && this.op.cells) {
      const ok = this.sheets.indexOf(sh) + ':' + k;
      if (!this.op.cells.has(ok)) this.op.cells.set(ok, old);
    }
    if (cell) {
      cell.mod = true;
      if (old && old.el && !cell.el) cell.el = old.el;
      if (old && old.si != null && cell.si == null) this.unshare(sh, old.si);
      sh.cells.set(k, cell);
      if (cell.v != null || cell.f != null) { if (r > sh.used.r) sh.used.r = r; if (c > sh.used.c) sh.used.c = c; }
    } else {
      if (old && old.si != null) this.unshare(sh, old.si);
      sh.cells.delete(k);
    }
    sh.touched = true;
    if (this.formulas && ((old && old.f != null) || (cell && cell.f != null))) this.formulasDirty = true;
    (this.changedCells || (this.changedCells = [])).push({ sh, r, c });
  }

  // Общая формула (одна на диапазон) после правки превращается в обычные — каждая в своей ячейке.
  unshare(sh, si) {
    for (const [k, cell] of sh.cells) {
      if (cell.si !== si) continue;
      const n = { ...cell, mod: true };
      delete n.si;
      sh.cells.set(k, n);
    }
    sh.touched = true;
  }

  commitOp() {
    const op = this.op;
    this.op = null;
    if (!op) return;
    this.recalc();
    if (op.cells.size) this.pushUndo({ cells: op.cells, after: new Map([...op.cells.keys()].map((ok) => { const [si, k] = ok.split(':'); return [ok, this.sheets[+si].cells.get(+k) || null]; })) });
    this.touch();
    this.syncBar();
    this.draw();
  }

  pushUndo(entry) {
    entry.sheet = this.activeIdx;
    entry.sel = { ...this.sh.sel };
    this.undoStack.push(entry);
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  // Снимок для операций со строками/столбцами — проще вернуть всё целиком.
  snapshot() {
    return {
      sheets: this.sheets.map((s) => (s.cells ? { s, cells: new Map(s.cells), cols: s.cols.map((k) => ({ ...k })), rowsInfo: new Map(s.rowsInfo), merges: s.merges.map((m) => ({ ...m })), ops: s.ops.slice(), name: s.name, touched: s.touched, colsMod: s.colsMod, mergesMod: s.mergesMod } : { s })),
      list: this.sheets.slice(),
      names: this.names.map((n) => ({ ...n })),
    };
  }
  restore(snap) {
    this.sheets = snap.list.slice();
    for (const x of snap.sheets) {
      if (!x.cells) continue;
      Object.assign(x.s, { cells: new Map(x.cells), cols: x.cols.map((k) => ({ ...k })), rowsInfo: new Map(x.rowsInfo), merges: x.merges.map((m) => ({ ...m })), ops: x.ops.slice(), name: x.name, touched: true, colsMod: x.colsMod, mergesMod: x.mergesMod });
      x.s.geo = null;
      this.indexMerges(x.s);
      this.usedOf(x.s);
    }
    this.names = snap.names.map((n) => ({ ...n }));
    this.formulas = null;
    this.sheetsMod = true;
  }

  undo(redo) {
    const from = redo ? this.redoStack : this.undoStack, to = redo ? this.undoStack : this.redoStack;
    const e = from.pop();
    if (!e) return;
    if (this.editing) this.cancelEdit();
    if (e.snap) this.restore(redo ? e.after : e.snap);
    else {
      for (const [ok, cell] of redo ? e.after : e.cells) {
        const [si, k] = ok.split(':');
        const sh = this.sheets[+si];
        if (cell) sh.cells.set(+k, cell); else sh.cells.delete(+k);
        sh.touched = true;
        (this.changedCells || (this.changedCells = [])).push({ sh, r: Math.floor(+k / 16384), c: +k % 16384 });
      }
      this.formulasDirty = true;
      this.recalc();
    }
    to.push(e);
    if (e.sheet != null && e.sheet < this.sheets.length) this.activeIdx = e.sheet;
    if (e.sel) Object.assign(this.sh.sel, e.sel);
    this.drawTabs();
    this.resetGeo();
    this.touch();
    this.syncBar();
    this.draw();
  }

  // Все формулы книги: разобранные, со ссылками.
  collectFormulas() {
    const list = [];
    this.sheets.forEach((sh, si) => {
      if (!sh.cells) return;
      for (const [k, cell] of sh.cells) {
        if (cell.f == null) continue;
        let ast = null, ok = false;
        try { ast = parseFormula(cell.f); ok = formulaSupported(ast); } catch {}
        const refs = [];
        if (ast) eachRef(ast, (n) => {
          if (n.bad) return;
          const sheet = n.sheet ? this.sheets.findIndex((s) => s.name.toLowerCase() === n.sheet.toLowerCase()) : si;
          const b = n.t === 'range' ? n.b : n.a;
          refs.push({ si: sheet, r1: n.a.r == null ? 0 : Math.min(n.a.r, b.r), r2: n.a.r == null ? XL_MAXR : Math.max(n.a.r, b.r), c1: n.a.c == null ? 0 : Math.min(n.a.c, b.c), c2: n.a.c == null ? XL_MAXC : Math.max(n.a.c, b.c) });
        });
        list.push({ si, k, r: Math.floor(k / 16384), c: k % 16384, ast, ok, refs, cell });
      }
    });
    this.formulas = list;
    this.formulasDirty = false;
  }

  // Пересчитываем только то, что зависит от изменённых ячеек (и сами изменённые формулы).
  recalc() {
    const changed = this.changedCells || [];
    this.changedCells = [];
    if (!changed.length) return;
    if (!this.formulas || this.formulasDirty) this.collectFormulas();
    const F = this.formulas;
    const dirty = new Set();
    const queue = [];
    for (const ch of changed) {
      const si = this.sheets.indexOf(ch.sh);
      queue.push({ si, r: ch.r, c: ch.c });
      const k = xkey(ch.r, ch.c);
      F.forEach((f, i) => { if (f.si === si && f.k === k) { dirty.add(i); } });
    }
    while (queue.length) {
      const x = queue.pop();
      for (let i = 0; i < F.length; i++) {
        if (dirty.has(i)) continue;
        if (F[i].refs.some((g) => g.si === x.si && x.r >= g.r1 && x.r <= g.r2 && x.c >= g.c1 && x.c <= g.c2)) {
          dirty.add(i);
          queue.push({ si: F[i].si, r: F[i].r, c: F[i].c });
        }
      }
    }
    if (!dirty.size) return;
    const byCell = new Map();
    for (const i of dirty) byCell.set(F[i].si + ':' + F[i].k, F[i]);
    const done = new Set(), busy = new Set();
    const self = this;
    const evalOne = (f) => {
      const id = f.si + ':' + f.k;
      if (done.has(id)) return;
      if (busy.has(id)) return; // круговая ссылка — оставляем как есть
      busy.add(id);
      const sh = this.sheets[f.si];
      const cur = sh.cells.get(f.k);
      if (cur && cur.f != null && f.ok && f.ast) {
        const v = evaluate(f.ast, ctxFor(f));
        if (cur.v !== v && !(isErr(cur.v) && isErr(v) && cur.v.err === v.err)) {
          const n = { ...cur, v, mod: true };
          sh.cells.set(f.k, n);
          f.cell = n;
        }
      } else if (cur && !f.ok) this.stale = true;
      busy.delete(id);
      done.add(id);
    };
    const ctxFor = (f) => ({
      sheet: this.sheets[f.si].name, r: f.r, c: f.c,
      cell(name, r, c) {
        const si = self.sheets.findIndex((s) => s.name.toLowerCase() === String(name).toLowerCase());
        if (si < 0 || !self.sheets[si].cells) return undefined;
        const k = xkey(r, c);
        const dep = byCell.get(si + ':' + k);
        if (dep) evalOne(dep);
        const x = self.sheets[si].cells.get(k);
        return x ? x.v : null;
      },
      dims(name) {
        const s = self.sheets.find((x) => x.name.toLowerCase() === String(name).toLowerCase());
        return s && s.cells ? s.used : null;
      },
    });
    for (const f of byCell.values()) evalOne(f);
  }

  touch() {
    if (!this.dirty) { this.dirty = true; this.onChange && this.onChange(); }
  }

  // ---------- Стили ячеек: новые xf в styles.xml ----------
  ensureXf(base, patch) {
    const key = base + '|' + JSON.stringify(patch);
    if (this.xfPatchCache.has(key)) return this.xfPatchCache.get(key);
    const x = this.stylesXml;
    if (!x) return base;
    const root = x.documentElement;
    const one = (p, n) => [...p.children].find((e) => e.localName === n);
    const ensure = (n, before) => {
      let e = one(root, n);
      if (!e) {
        e = x.createElementNS(XNS.main, n);
        const ref = before.map((b) => one(root, b)).find(Boolean);
        root.insertBefore(e, ref || null);
      }
      return e;
    };
    const cellXfs = ensure('cellXfs', ['cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']);
    const xfEls = [...cellXfs.children].filter((e) => e.localName === 'xf');
    const src = xfEls[base] || xfEls[0];
    const xf = src ? src.cloneNode(true) : x.createElementNS(XNS.main, 'xf');
    const cur = this.xfs[base] || this.xfs[0];
    const add = (listName, elName, node, before) => {
      const list = ensure(listName, before);
      list.append(node);
      const n = [...list.children].filter((e) => e.localName === elName).length;
      list.setAttribute('count', n);
      return n - 1;
    };
    if (patch.b != null || patch.i != null || patch.u != null || patch.color != null) {
      const fontsEl = ensure('fonts', ['fills', 'borders', 'cellStyleXfs', 'cellXfs']);
      const fEls = [...fontsEl.children].filter((e) => e.localName === 'font');
      const f = (fEls[cur.fontId] || fEls[0]).cloneNode(true);
      const order = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
      const setFlag = (n, on, val) => {
        for (const e of [...f.children].filter((k) => k.localName === n)) e.remove();
        if (!on) return;
        const e = x.createElementNS(XNS.main, n);
        if (val) for (const [a, b] of Object.entries(val)) e.setAttribute(a, b);
        const idx = order.indexOf(n);
        const next = [...f.children].find((k) => order.indexOf(k.localName) > idx);
        f.insertBefore(e, next || null);
      };
      const now = this.fonts[cur.fontId] || {};
      if (patch.b != null) setFlag('b', patch.b);
      if (patch.i != null) setFlag('i', patch.i);
      if (patch.u != null) setFlag('u', patch.u);
      if (patch.color != null) setFlag('color', true, { rgb: 'FF' + patch.color.replace('#', '').toUpperCase() });
      void now;
      const id = add('fonts', 'font', f, ['fills', 'borders', 'cellStyleXfs', 'cellXfs']);
      this.fonts.push(this.parseFont(f));
      xf.setAttribute('fontId', id);
      xf.setAttribute('applyFont', '1');
    }
    if (patch.fill !== undefined) {
      let id = 0;
      if (patch.fill) {
        const fill = x.createElementNS(XNS.main, 'fill');
        fill.innerHTML = '';
        const pf = x.createElementNS(XNS.main, 'patternFill');
        pf.setAttribute('patternType', 'solid');
        const fg = x.createElementNS(XNS.main, 'fgColor');
        fg.setAttribute('rgb', 'FF' + patch.fill.replace('#', '').toUpperCase());
        const bg = x.createElementNS(XNS.main, 'bgColor');
        bg.setAttribute('indexed', '64');
        pf.append(fg, bg);
        fill.append(pf);
        id = add('fills', 'fill', fill, ['borders', 'cellStyleXfs', 'cellXfs']);
        this.fills.push({ color: patch.fill });
      }
      xf.setAttribute('fillId', id);
      xf.setAttribute('applyFill', '1');
    }
    if (patch.h !== undefined || patch.wrap !== undefined) {
      let al = one(xf, 'alignment');
      if (!al) { al = x.createElementNS(XNS.main, 'alignment'); xf.insertBefore(al, xf.firstChild); }
      if (patch.h !== undefined) { if (patch.h) al.setAttribute('horizontal', patch.h); else al.removeAttribute('horizontal'); }
      if (patch.wrap !== undefined) { if (patch.wrap) al.setAttribute('wrapText', '1'); else al.removeAttribute('wrapText'); }
      xf.setAttribute('applyAlignment', '1');
    }
    if (patch.numFmt !== undefined) {
      let id = patch.numFmt;
      if (typeof id === 'string') {
        const found = Object.entries(this.numFmts).find(([, code]) => code === id);
        if (found) id = +found[0];
        else {
          const list = ensure('numFmts', ['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs']);
          const next = Math.max(163, ...Object.keys(this.numFmts).map(Number)) + 1;
          const n = x.createElementNS(XNS.main, 'numFmt');
          n.setAttribute('numFmtId', next);
          n.setAttribute('formatCode', id);
          list.append(n);
          list.setAttribute('count', list.children.length);
          root.insertBefore(list, root.firstElementChild === list ? list.nextSibling : root.firstElementChild);
          this.numFmts[next] = id;
          id = next;
        }
      }
      xf.setAttribute('numFmtId', id);
      xf.setAttribute('applyNumberFormat', '1');
    }
    cellXfs.append(xf);
    cellXfs.setAttribute('count', [...cellXfs.children].filter((e) => e.localName === 'xf').length);
    this.xfs.push(this.parseXf(xf));
    const idx = this.xfs.length - 1;
    this.xfPatchCache.set(key, idx);
    this.stylesTouched = true;
    return idx;
  }

  // Применить к выделению (без выхода за заполненную область у целых строк/столбцов).
  styleSel(patchFn) {
    if (this.editing) this.commitEdit();
    const s = this.sh.sel, sh = this.sh;
    const r2 = Math.min(s.r2, Math.max(sh.used.r, s.r1)), c2 = Math.min(s.c2, Math.max(sh.used.c, s.c1));
    if ((r2 - s.r1 + 1) * (c2 - s.c1 + 1) > 400000) { toast('Слишком большая область'); return; }
    this.beginOp();
    for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) {
      const m = sh.mergeAt.get(xkey(r, c));
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const cell = this.cell(r, c);
      const patch = patchFn(cell);
      const ns = this.ensureXf(cell ? cell.s : 0, patch);
      if (cell && cell.s === ns) continue;
      this.setCell(r, c, cell ? { ...cell, s: ns } : { s: ns, v: null });
    }
    this.op.cells.size || (this.op = null);
    if (this.op) { const op = this.op; this.op = null; this.pushUndo({ cells: op.cells, after: new Map([...op.cells.keys()].map((ok) => { const [si, k] = ok.split(':'); return [ok, this.sheets[+si].cells.get(+k) || null]; })) }); }
    this.changedCells = [];
    this.touch();
    this.syncBar();
    this.draw();
  }

  toggleStyle(flag) {
    const s = this.sh.sel;
    const first = this.cell(s.r, s.c);
    const on = !(first && this.st(first.s).font[flag]);
    this.styleSel(() => ({ [flag]: on }));
  }

  // ---------- Строки и столбцы ----------
  // axis 'r' | 'c'; n > 0 — вставить n перед at; n < 0 — удалить |n| начиная с at.
  shiftAxis(axis, at, n) {
    if (this.editing) this.commitEdit();
    const sh = this.sh;
    const snap = this.snapshot();
    const idx = this.activeIdx;
    const shiftKey = (k) => {
      let r = Math.floor(k / 16384), c = k % 16384;
      const v = axis === 'r' ? r : c;
      if (n > 0) { if (v >= at) { if (axis === 'r') r += n; else c += n; } }
      else { const d = -n; if (v >= at && v < at + d) return -1; if (v >= at + d) { if (axis === 'r') r -= d; else c -= d; } }
      if (r >= XL_MAXR || c >= XL_MAXC) return -1;
      return xkey(r, c);
    };
    const moved = new Map();
    for (const [k, cell] of sh.cells) {
      const nk = shiftKey(k);
      if (nk < 0) continue;
      moved.set(nk, nk === k ? cell : { ...cell, mod: true });
    }
    sh.cells = moved;
    // Формулы во всей книге
    for (const s2 of this.sheets) {
      if (!s2.cells) continue;
      const onSheet = (name) => (name ? name.toLowerCase() === sh.name.toLowerCase() : s2 === sh);
      for (const [k, cell] of s2.cells) {
        if (cell.f == null) continue;
        let ast;
        try { ast = parseFormula(cell.f); } catch { continue; }
        if (adjustAst(ast, onSheet, axis, at, n)) s2.cells.set(k, { ...cell, f: astStr(ast), mod: true });
        else if (cell.si != null && s2 === sh) s2.cells.set(k, { ...cell, mod: true });
      }
      for (const [k, cell] of s2.cells) if (cell.si != null && s2 === sh) { const x = { ...cell }; delete x.si; s2.cells.set(k, x); }
      s2.touched = s2.touched || s2 === sh;
    }
    for (const nm of this.names) {
      try {
        const ast = parseFormula(nm.text);
        if (adjustAst(ast, (name) => name && name.toLowerCase() === sh.name.toLowerCase(), axis, at, n)) { nm.text = astStr(ast); nm.mod = true; }
      } catch {}
    }
    // Высоты строк / ширины столбцов
    if (axis === 'r') {
      const ri = new Map();
      for (const [r, info] of sh.rowsInfo) {
        const nr = n > 0 ? (r >= at ? r + n : r) : r >= at && r < at - n ? -1 : r >= at - n ? r + n : r;
        if (nr >= 0 && nr < XL_MAXR) ri.set(nr, info);
      }
      sh.rowsInfo = ri;
    } else {
      const out = [];
      for (const k of sh.cols) {
        let { min, max } = k;
        if (n > 0) { if (min >= at) { min += n; max += n; } else if (max >= at) max += n; }
        else {
          const d = -n, end = at + d - 1;
          if (min >= at && max <= end) continue;
          const nmin = min > end ? min - d : min >= at ? at : min;
          const nmax = max > end ? max - d : max >= at ? at - 1 : max;
          min = nmin; max = nmax;
          if (max < min) continue;
        }
        out.push({ ...k, min: Math.min(min, XL_MAXC - 1), max: Math.min(max, XL_MAXC - 1) });
      }
      sh.cols = out;
      sh.colsMod = true;
    }
    // Объединения
    const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
    sh.merges = sh.merges.map((m) => {
      const x = { ...m };
      if (n > 0) { if (x[k1] >= at) { x[k1] += n; x[k2] += n; } else if (x[k2] >= at) x[k2] += n; return x; }
      const d = -n, end = at + d - 1;
      if (x[k1] >= at && x[k2] <= end) return null;
      x[k1] = x[k1] > end ? x[k1] - d : x[k1] >= at ? at : x[k1];
      x[k2] = x[k2] > end ? x[k2] - d : x[k2] >= at ? at - 1 : x[k2];
      return x[k2] >= x[k1] && (x.r2 > x.r1 || x.c2 > x.c1) ? x : null;
    }).filter(Boolean);
    sh.mergesMod = true;
    this.indexMerges(sh);
    sh.ops.push({ axis, at, n });
    sh.touched = true;
    this.usedOf(sh);
    this.resetGeo(sh);
    // Пересчёт всех формул книги, которые теперь смотрят на другие места
    this.formulas = null;
    const all = [];
    for (const s2 of this.sheets) if (s2.cells) for (const [k, cell] of s2.cells) if (cell.f != null) all.push({ sh: s2, r: Math.floor(k / 16384), c: k % 16384 });
    this.changedCells = all;
    this.recalc();
    this.pushUndo({ snap, after: this.snapshot() });
    this.undoStack[this.undoStack.length - 1].sheet = idx;
    this.touch();
    this.syncBar();
    this.draw();
  }

  insertRows(where) {
    const s = this.sh.sel;
    const cnt = s.r2 >= XL_MAXR - 1 ? 1 : s.r2 - s.r1 + 1;
    const at = where === 'below' ? s.r2 + 1 : s.r1;
    this.shiftAxis('r', at, cnt);
    this.selectRows(at, at + cnt - 1);
  }
  deleteRows() {
    const s = this.sh.sel;
    if (s.r1 === 0 && s.r2 >= XL_MAXR - 1) { this.clearSel(); return; }
    this.shiftAxis('r', s.r1, -(s.r2 - s.r1 + 1));
    this.select(s.r1, s.c);
  }
  insertCols(where) {
    const s = this.sh.sel;
    const cnt = s.c2 >= XL_MAXC - 1 ? 1 : s.c2 - s.c1 + 1;
    const at = where === 'right' ? s.c2 + 1 : s.c1;
    this.shiftAxis('c', at, cnt);
    this.selectCols(at, at + cnt - 1);
  }
  deleteCols() {
    const s = this.sh.sel;
    if (s.c1 === 0 && s.c2 >= XL_MAXC - 1) { this.clearSel(); return; }
    this.shiftAxis('c', s.c1, -(s.c2 - s.c1 + 1));
    this.select(s.r, s.c1);
  }

  clearSel() {
    if (this.editing) { this.cancelEdit(); }
    const s = this.sh.sel, sh = this.sh;
    this.beginOp();
    for (const [k, cell] of [...sh.cells]) {
      const r = Math.floor(k / 16384), c = k % 16384;
      if (r < s.r1 || r > s.r2 || c < s.c1 || c > s.c2) continue;
      if (cell.v == null && cell.f == null) continue;
      this.setCell(r, c, cell.s ? { s: cell.s, v: null } : null);
    }
    this.commitOp();
  }

  // ---------- Буфер обмена ----------
  isActive() {
    return !this.root.hidden && !!this.root.offsetParent && document.activeElement === this.viewer;
  }

  handleCopy(e, cut) {
    if (!this.isActive()) return;
    const s = this.sh.sel, sh = this.sh;
    const r2 = Math.min(s.r2, sh.used.r), c2 = Math.min(s.c2, sh.used.c);
    const rows = [];
    const cells = [];
    for (let r = s.r1; r <= r2; r++) {
      const row = [], crow = [];
      for (let c = s.c1; c <= c2; c++) {
        const cell = this.cell(r, c);
        const d = this.display(cell);
        let t = d ? d.text : '';
        if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        row.push(t);
        crow.push(cell ? { ...cell } : null);
      }
      rows.push(row.join('\t'));
      cells.push(crow);
    }
    const text = rows.join('\r\n');
    e.preventDefault();
    e.clipboardData.setData('text/plain', text);
    XlsxEditor.clip = { text, cells, r: s.r1, c: s.c1, from: this, sheet: sh, cut };
    if (cut) {
      this.cutRange = { ...s, sheet: sh };
      this.draw();
    }
  }

  handlePaste(e) {
    if (!this.isActive()) return;
    const text = e.clipboardData.getData('text/plain');
    const clip = XlsxEditor.clip;
    e.preventDefault();
    const s = this.sh.sel;
    this.beginOp();
    if (clip && clip.text.replace(/\r/g, '') === text.replace(/\r/g, '')) {
      // Скопировано в Folio: переносим формулы со сдвигом и оформление
      const cr = clip.cells.length, cc = clip.cells[0] ? clip.cells[0].length : 0;
      const tr = Math.max(1, Math.floor((s.r2 - s.r1 + 1) / cr)), tc = Math.max(1, Math.floor((s.c2 - s.c1 + 1) / cc));
      for (let a = 0; a < (s.r2 < XL_MAXR - 1 ? tr : 1); a++) for (let b = 0; b < (s.c2 < XL_MAXC - 1 ? tc : 1); b++) {
        clip.cells.forEach((row, i) => row.forEach((cell, j) => {
          const r = s.r1 + a * cr + i, c = s.c1 + b * cc + j;
          if (!cell) { this.setCell(r, c, null); return; }
          const n = { s: clip.from === this ? cell.s : 0, v: cell.v };
          if (cell.f != null) n.f = clip.cut ? cell.f : shiftFormula(cell.f, r - (clip.r + i), c - (clip.c + j));
          this.setCell(r, c, n);
        }));
      }
      if (clip.cut && clip.from === this) {
        const src = clip.sheet;
        for (let i = 0; i < cr; i++) for (let j = 0; j < cc; j++) {
          const r = clip.r + i, c = clip.c + j;
          if (src === this.sh && r >= s.r1 && r < s.r1 + cr && c >= s.c1 && c < s.c1 + cc) continue;
          this.setCell(r, c, null, src);
        }
        XlsxEditor.clip = null;
        this.cutRange = null;
      }
      Object.assign(this.sh.sel, { r2: s.r1 + cr * (s.r2 < XL_MAXR - 1 ? tr : 1) - 1, c2: s.c1 + cc * (s.c2 < XL_MAXC - 1 ? tc : 1) - 1 });
    } else {
      // Текст из другой программы: строки по переносам, ячейки по табуляции
      const rows = parseTsv(text);
      rows.forEach((row, i) => row.forEach((t, j) => {
        const r = s.r1 + i, c = s.c1 + j;
        const old = this.cell(r, c);
        const p = this.parseInput(t, old);
        if (p.bad) { this.setCell(r, c, { s: old ? old.s : 0, v: t }); return; }
        let st = old ? old.s : 0;
        if (p.fmt != null) st = this.ensureXf(st, { numFmt: p.fmt });
        if (p.v == null && p.f == null) this.setCell(r, c, st ? { s: st, v: null } : null);
        else this.setCell(r, c, p.f != null ? { s: st, f: p.f, v: null } : { s: st, v: p.v });
      }));
      if (rows.length) Object.assign(this.sh.sel, { r2: s.r1 + rows.length - 1, c2: s.c1 + Math.max(...rows.map((r) => r.length)) - 1 });
    }
    this.commitOp();
  }

  // ---------- Листы ----------
  drawTabs() {
    this.tabsEl.textContent = '';
    this.sheets.forEach((sh, i) => {
      if (sh.kind === 'other' || sh.state !== 'visible') return;
      const b = el('button', 'xl-tab' + (i === this.activeIdx ? ' on' : ''));
      b.textContent = sh.name;
      b.addEventListener('pointerdown', (e) => { if (this.editing && this.canPoint()) e.preventDefault(); });
      b.addEventListener('click', () => this.showSheet(i));
      b.addEventListener('dblclick', () => this.renameSheet(i, b));
      b.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const vis = this.sheets.filter((s) => s.kind !== 'other' && s.state === 'visible').length;
        const items = [['Переименовать', () => this.renameSheet(i, b)], ['Новый лист', () => this.addSheet()]];
        if (vis > 1) items.push(null, ['Удалить лист', () => this.deleteSheet(i)]);
        showMenu(e.clientX, e.clientY, items);
      });
      this.tabsEl.append(b);
    });
  }

  showSheet(i) {
    if (!this.sheets[i] || this.sheets[i].kind === 'other') return;
    const pointing = this.editing && this.canPoint();
    if (this.editing && !pointing) this.commitEdit();
    if (this.sh) this.sh.scroll = { l: this.viewer.scrollLeft, t: this.viewer.scrollTop };
    this.activeIdx = i;
    this.drawTabs();
    this.updateSpace();
    const sc = this.sh.scroll;
    this.viewer.scrollLeft = sc ? sc.l : 0;
    this.viewer.scrollTop = sc ? sc.t : 0;
    if (pointing) this.editBox.hidden = this.editing.sheet !== i;
    this.syncBar();
    this.draw();
    const tab = this.tabsEl.querySelector('.on');
    if (tab) tab.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  validName(name, except) {
    name = name.trim();
    if (!name || name.length > 31 || /[\\/?*[\]:]/.test(name) || /^'|'$/.test(name)) return null;
    if (this.sheets.some((s, i) => i !== except && s.name.toLowerCase() === name.toLowerCase())) return null;
    return name;
  }

  renameSheet(i, tabBtn) {
    const sh = this.sheets[i];
    const inp = el('input', 'xl-tab-input');
    inp.value = sh.name;
    tabBtn.replaceWith(inp);
    inp.focus();
    inp.select();
    let done = false;
    const finish = (ok) => {
      if (done) return;
      done = true;
      const name = ok ? this.validName(inp.value, i) : null;
      if (ok && !name && inp.value.trim() !== sh.name) toast('Такое имя нельзя: до 31 знака, без : \\ / ? * [ ] и не повторяться');
      if (name && name !== sh.name) {
        const snap = this.snapshot();
        const old = sh.name;
        sh.name = name;
        for (const s2 of this.sheets) {
          if (!s2.cells) continue;
          for (const [k, cell] of s2.cells) {
            if (cell.f == null || !cell.f.toLowerCase().includes(old.toLowerCase())) continue;
            try { const ast = parseFormula(cell.f); if (renameSheetAst(ast, old, name)) { s2.cells.set(k, { ...cell, f: astStr(ast), mod: true }); s2.touched = true; } } catch {}
          }
        }
        for (const nm of this.names) {
          try { const ast = parseFormula(nm.text); if (renameSheetAst(ast, old, name)) { nm.text = astStr(ast); nm.mod = true; } } catch {}
        }
        this.formulas = null;
        this.sheetsMod = true;
        this.pushUndo({ snap, after: this.snapshot() });
        this.touch();
      }
      this.drawTabs();
    };
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') finish(true); else if (e.key === 'Escape') finish(false); e.stopPropagation(); });
    inp.addEventListener('blur', () => finish(true));
  }

  addSheet() {
    if (this.editing) this.commitEdit();
    const snap = this.snapshot();
    let n = this.sheets.length + 1, name;
    const base = T('Лист');
    do { name = base + n++; } while (this.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase()));
    const sh = { name, isNew: true, state: 'visible', cells: new Map(), rowsInfo: new Map(), cols: [], merges: [], ops: [], defRowH: XL_DEF_ROWH, defColW: XL_DEF_COLW, grid: true, frozen: { r: 0, c: 0 }, sel: { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 }, used: { r: 0, c: 0 }, touched: true };
    this.indexMerges(sh);
    this.sheets.push(sh);
    this.sheetsMod = true;
    this.pushUndo({ snap, after: this.snapshot() });
    this.touch();
    this.showSheet(this.sheets.length - 1);
  }

  async deleteSheet(i) {
    const sh = this.sheets[i];
    if (sh.cells && sh.cells.size) {
      const r = await window.folio.ask({ message: T(`Удалить лист «${sh.name}»?`), detail: T('Данные на нём пропадут.'), buttons: [T('Удалить'), T('Отмена')], cancelId: 1 });
      if (r !== 0) return;
    }
    const snap = this.snapshot();
    this.sheets.splice(i, 1);
    (this.deletedSheets || (this.deletedSheets = [])).push(sh);
    this.sheetsMod = true;
    // Ссылки на удалённый лист → #REF!
    this.formulas = null;
    const all = [];
    for (const s2 of this.sheets) if (s2.cells) for (const [k, cell] of s2.cells) if (cell.f != null) all.push({ sh: s2, r: Math.floor(k / 16384), c: k % 16384 });
    this.changedCells = all;
    this.recalc();
    this.pushUndo({ snap, after: this.snapshot() });
    this.touch();
    this.activeIdx = Math.min(i, this.sheets.length - 1);
    while (this.activeIdx > 0 && (this.sheets[this.activeIdx].kind === 'other' || this.sheets[this.activeIdx].state !== 'visible')) this.activeIdx--;
    this.showSheet(this.activeIdx);
  }

  // ---------- Меню ----------
  cellMenu(x, y, mode) {
    const items = [];
    items.push(['Вырезать', () => document.execCommand('cut'), 'Ctrl+X']);
    items.push(['Копировать', () => document.execCommand('copy'), 'Ctrl+C']);
    items.push(['Вставить', () => this.pasteFromMenu(), 'Ctrl+V']);
    items.push(null);
    if (mode !== 'cols') {
      items.push(['Вставить строку выше', () => this.insertRows('above')]);
      items.push(['Вставить строку ниже', () => this.insertRows('below')]);
    }
    if (mode !== 'rows') {
      items.push(['Вставить столбец слева', () => this.insertCols('left')]);
      items.push(['Вставить столбец справа', () => this.insertCols('right')]);
    }
    items.push(null);
    if (mode !== 'cols') items.push(['Удалить строки', () => this.deleteRows()]);
    if (mode !== 'rows') items.push(['Удалить столбцы', () => this.deleteCols()]);
    items.push(['Очистить', () => this.clearSel(), 'Delete']);
    if (mode === 'cols') items.push(null, ['Ширина по содержимому', () => this.autoFit(this.sh.sel.c1)]);
    showMenu(x, y, items);
  }

  async pasteFromMenu() {
    this.viewer.focus();
    try {
      const text = await navigator.clipboard.readText();
      const dt = new DataTransfer();
      dt.setData('text/plain', text);
      this.handlePaste({ clipboardData: dt, preventDefault() {} });
    } catch { toast('Нажмите Ctrl+V'); }
  }

  // ---------- Масштаб ----------
  setZoom(z) {
    z = clamp(z, 0.4, 3);
    const v = this.view();
    const cx = (v.sl) / this.zoom, cy = (v.st) / this.zoom;
    this.zoom = z;
    this.updateSpace();
    this.viewer.scrollLeft = cx * z;
    this.viewer.scrollTop = cy * z;
    this.flashHud(Math.round(z * 100) + '%');
    this.draw();
  }

  flashHud(text) {
    this.hud.textContent = text;
    this.hud.classList.add('show');
    clearTimeout(this.hudTimer);
    this.hudTimer = setTimeout(() => this.hud.classList.remove('show'), 900);
  }

  // ---------- Клавиатура (когда фокус на сетке) ----------
  onKey(e) {
    if (document.activeElement !== this.viewer) return false;
    const ctrl = e.ctrlKey || e.metaKey, k = e.key;
    if (ctrl && (k === '=' || k === '+')) { this.setZoom(this.zoom * 1.1); return true; }
    if (ctrl && k === '-') { this.setZoom(this.zoom / 1.1); return true; }
    if (ctrl && k === '0') { this.setZoom(1); return true; }
    if (ctrl && (e.code === 'KeyZ')) { this.undo(e.shiftKey); return true; }
    if (ctrl && (e.code === 'KeyY')) { this.undo(true); return true; }
    if (ctrl && e.code === 'KeyB') { this.toggleStyle('b'); return true; }
    if (ctrl && e.code === 'KeyI') { this.toggleStyle('i'); return true; }
    if (ctrl && e.code === 'KeyU') { this.toggleStyle('u'); return true; }
    if (ctrl && e.code === 'KeyA') { this.selectAll(); return true; }
    if (ctrl && e.code === 'KeyD') { const s = this.sh.sel; if (s.r2 > s.r1) this.fill({ ...s, r2: s.r1 }, { axis: 'r', to: s.r2 }); return true; }
    if (ctrl && e.code === 'KeyR') { const s = this.sh.sel; if (s.c2 > s.c1) this.fill({ ...s, c2: s.c1 }, { axis: 'c', to: s.c2 }); return true; }
    if (ctrl && k === 'Home') { this.select(0, 0); this.reveal(0, 0); return true; }
    if (ctrl && k === 'End') { const u = this.usedOf(this.sh); this.select(u.r, u.c); this.reveal(u.r, u.c); return true; }
    if (e.altKey && k === '=') { this.autoSum(); return true; }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { this.move(...arrows[k], e.shiftKey, ctrl); return true; }
    if (k === 'Enter') { this.move(e.shiftKey ? -1 : 1, 0); return true; }
    if (k === 'Tab') { this.move(0, e.shiftKey ? -1 : 1); return true; }
    if (k === 'PageDown' || k === 'PageUp') {
      const rows = Math.max(1, Math.floor(this.viewer.clientHeight / (this.sh.defRowH * this.zoom)) - 1);
      this.move(k === 'PageDown' ? rows : -rows, 0, e.shiftKey);
      return true;
    }
    if (k === 'Home') { this.select(this.sh.sel.r, 0, e.shiftKey); this.reveal(this.sh.sel.r, 0); return true; }
    if (k === 'Delete' || k === 'Backspace') {
      if (k === 'Backspace') { this.clearSel(); this.startEdit(true, ''); return true; }
      this.clearSel();
      return true;
    }
    if (k === 'F2') { this.startEdit(false); return true; }
    if (k === 'Escape') { if (this.cutRange) { this.cutRange = null; XlsxEditor.clip = null; this.draw(); } return false; }
    if (!ctrl && !e.altKey && k.length === 1) {
      // Начали печатать — правим ячейку с нуля; сам символ попадёт в поле
      this.startEdit(true, '');
      return false;
    }
    return false;
  }

  // Σ: сумма столбика над ячейкой (или строки слева).
  autoSum() {
    if (this.editing) this.commitEdit();
    const s = this.sh.sel;
    const isNum = (r, c) => { const x = this.cell(r, c); return x && typeof x.v === 'number'; };
    let r = s.r, c = s.c, f = null;
    let r0 = r - 1;
    while (r0 >= 0 && isNum(r0, c)) r0--;
    if (r0 < r - 1) f = `SUM(${addr(r0 + 1, c)}:${addr(r - 1, c)})`;
    else {
      let c0 = c - 1;
      while (c0 >= 0 && isNum(r, c0)) c0--;
      if (c0 < c - 1) f = `SUM(${addr(r, c0 + 1)}:${addr(r, c - 1)})`;
    }
    if (!f) { this.startEdit(true, '=SUM()'); this.editBox.setSelectionRange(5, 5); this.fBox.value = this.editBox.value; return; }
    this.beginOp();
    const old = this.cell(r, c);
    this.setCell(r, c, { s: old ? old.s : 0, f, v: null });
    this.commitOp();
  }

  // ---------- Панель инструментов ----------
  buildDock() {
    const dock = el('div', 'dock xl-dock');
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
    ui.undo = btn(ICONS.undo, 'Отменить · Ctrl+Z', () => this.undo());
    ui.redo = btn(ICONS.redo, 'Вернуть · Ctrl+Y', () => this.undo(true));
    sep();
    ui.b = btn(ICONS.bold, 'Жирный · Ctrl+B', () => this.toggleStyle('b'));
    ui.i = btn(ICONS.italic, 'Курсив · Ctrl+I', () => this.toggleStyle('i'));
    ui.color = btn(ICONS.textColor + '<span class="dot"></span>', 'Цвет текста', (e, b) => this.swatchMenu(b, XL_INK, (c) => this.styleSel(() => ({ color: c }))), 'wd-color');
    ui.fill = btn(ICONS.fill + '<span class="dot"></span>', 'Заливка', (e, b) => this.swatchMenu(b, XL_FILLS, (c) => this.styleSel(() => ({ fill: c === '#FFFFFF' ? null : c })), true), 'wd-color');
    ui.align = btn(ICONS.alignLeft, 'Выравнивание', (e, b) => this.menuAt(b, [
      ['По левому краю', () => this.styleSel(() => ({ h: 'left' }))],
      ['По центру', () => this.styleSel(() => ({ h: 'center' }))],
      ['По правому краю', () => this.styleSel(() => ({ h: 'right' }))],
      ['Как в Excel (числа справа)', () => this.styleSel(() => ({ h: null }))],
      null,
      ['Переносить по словам', () => { const c = this.cell(this.sh.sel.r, this.sh.sel.c); const on = !(c && this.st(c.s).align.wrap); this.styleSel(() => ({ wrap: on })); }],
    ]));
    ui.fmt = btn('<span class="xl-fmt-txt">123</span>', 'Формат чисел', (e, b) => this.menuAt(b, XL_NUMFMTS.map(([n, id]) => [n, () => this.styleSel(() => ({ numFmt: id }))])), 'wd-txt xl-fmtbtn');
    sep();
    btn(ICONS.sigma, 'Автосумма · Alt+=', () => this.autoSum());
    btn(ICONS.rowAdd, 'Строки и столбцы', (e, b) => this.menuAt(b, [
      ['Вставить строку выше', () => this.insertRows('above')],
      ['Вставить строку ниже', () => this.insertRows('below')],
      ['Вставить столбец слева', () => this.insertCols('left')],
      ['Вставить столбец справа', () => this.insertCols('right')],
      null,
      ['Удалить строки', () => this.deleteRows()],
      ['Удалить столбцы', () => this.deleteCols()],
    ]));
    sep();
    btn(ICONS.hide, 'Скрыть панель · Ctrl+.', () => window.setDockHidden && window.setDockHidden(true));
    const handle = el('button', 'dock-handle');
    handle.dataset.tip = 'Показать панель · Ctrl+.';
    handle.addEventListener('click', () => window.setDockHidden && window.setDockHidden(false));
    this.status = el('div', 'xl-status');
    this.status.hidden = true;
    this.root.append(dock, handle, this.status);
    this.dock = dock;
    this.dockMover = makeDockMovable(this.root, dock, null, 'xlsx');
  }

  menuAt(b, items) { DocxEditor.prototype.menuAt.call(this, b, items); }

  swatchMenu(anchor, list, fn, withNone) {
    const row = el('div', 'wd-sw');
    for (const col of list) {
      const none = withNone && col === '#FFFFFF';
      const b = el('button', 'sw', none ? '<i class="none"></i>' : '<i></i>');
      b.style.setProperty('--c', col);
      b.dataset.tip = none ? 'Без заливки' : col;
      b.addEventListener('click', () => { this.pop && this.pop.remove(); fn(col); });
      row.append(b);
    }
    DocxEditor.prototype.popover.call(this, anchor, row);
  }

  syncTools() {
    if (!this.ui) return;
    const s = this.sh.sel, cell = this.cell(s.r, s.c);
    const st = this.st(cell ? cell.s : 0);
    this.ui.b.classList.toggle('on', !!st.font.b);
    this.ui.i.classList.toggle('on', !!st.font.i);
    this.ui.color.style.setProperty('--c', st.font.color || '#000000');
    this.ui.fill.style.setProperty('--c', st.fill || 'transparent');
    this.ui.align.innerHTML = { center: ICONS.alignCenter, right: ICONS.alignRight }[st.align.h] || ICONS.alignLeft;
    const xf = this.xfs[cell ? cell.s : 0] || {};
    const id = xf.numFmtId || 0;
    const code = this.fmtCode(id);
    const cur = (code.replace(/\[[^\]]*\]/g, '').match(/₸|\$|€|£|₽|р\./) || [])[0];
    const label = id === 0 ? '123' : fmtIsDate(code) ? '📅' : /%/.test(code) ? '%' : cur ? cur : id === 49 ? 'Aa' : '0,00';
    this.ui.fmt.firstChild.textContent = label;
  }

  // ---------- Сохранение ----------
  w(doc, name) { return doc.createElementNS(XNS.main, name); }

  sstAdd(s) {
    if (!this.sstXml) return null;
    if (!this.sstIndex) {
      this.sstIndex = new Map();
      this.sst.forEach((t, i) => { if (!this.sstIndex.has(t)) this.sstIndex.set(t, i); });
    }
    if (this.sstIndex.has(s)) return this.sstIndex.get(s);
    const si = this.sstXml.createElementNS(XNS.main, 'si');
    const t = this.sstXml.createElementNS(XNS.main, 't');
    if (/^\s|\s$|\n/.test(s)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = s;
    si.append(t);
    this.sstXml.documentElement.append(si);
    this.sst.push(s);
    this.sstIndex.set(s, this.sst.length - 1);
    this.sstTouched = true;
    return this.sst.length - 1;
  }

  cellEl(doc, r, c, cell) {
    if (!cell.mod && cell.el) {
      if (cell.el.getAttribute('r') !== addr(r, c)) cell.el.setAttribute('r', addr(r, c));
      return cell.el;
    }
    const e = cell.el ? cell.el.cloneNode(false) : this.w(doc, 'c');
    e.setAttribute('r', addr(r, c));
    for (const a of ['t', 'vm', 'cm']) e.removeAttribute(a);
    if (cell.s) e.setAttribute('s', cell.s); else e.removeAttribute('s');
    const v = cell.v;
    const putV = (text) => { const x = this.w(doc, 'v'); x.textContent = text; e.append(x); };
    if (cell.f != null) {
      const f = this.w(doc, 'f');
      if (cell.si != null && cell.el) {
        // общая формула, которую не трогали по существу — оставляем как была
        const old = [...cell.el.children].find((k) => k.localName === 'f');
        if (old) { e.append(old.cloneNode(true)); }
        else { f.textContent = cell.f; e.append(f); }
      } else {
        f.textContent = cell.f;
        if (cell.arr && cell.el) { const old = [...cell.el.children].find((k) => k.localName === 'f'); if (old && old.getAttribute('t') === 'array') { f.setAttribute('t', 'array'); f.setAttribute('ref', old.getAttribute('ref')); } }
        e.append(f);
      }
      if (v == null) {}
      else if (typeof v === 'number') putV(String(v));
      else if (typeof v === 'boolean') { e.setAttribute('t', 'b'); putV(v ? '1' : '0'); }
      else if (isErr(v)) { e.setAttribute('t', 'e'); putV(v.err); }
      else { e.setAttribute('t', 'str'); putV(String(v)); }
      return e;
    }
    if (v == null) return e;
    if (typeof v === 'number') putV(String(v));
    else if (typeof v === 'boolean') { e.setAttribute('t', 'b'); putV(v ? '1' : '0'); }
    else if (isErr(v)) { e.setAttribute('t', 'e'); putV(v.err); }
    else {
      const idx = this.sstAdd(String(v));
      if (idx != null) { e.setAttribute('t', 's'); putV(String(idx)); }
      else {
        e.setAttribute('t', 'inlineStr');
        const is = this.w(doc, 'is'), t = this.w(doc, 't');
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        t.textContent = String(v);
        is.append(t);
        e.append(is);
      }
    }
    return e;
  }

  // Порядок дочерних элементов worksheet по схеме.
  static WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

  wsPlace(doc, node) {
    const root = doc.documentElement;
    const idx = XlsxEditor.WS_ORDER.indexOf(node.localName);
    const next = [...root.children].find((k) => XlsxEditor.WS_ORDER.indexOf(k.localName) > idx);
    root.insertBefore(node, next || null);
  }

  writeSheet(sh) {
    if (!sh.xml) sh.xml = new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${XNS.main}" xmlns:r="${XNS.r}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`, 'application/xml');
    const doc = sh.xml, root = doc.documentElement;
    const one = (n) => [...root.children].find((e) => e.localName === n);
    let data = one('sheetData');
    if (!data) { data = this.w(doc, 'sheetData'); this.wsPlace(doc, data); }
    // Строки по порядку: исходные атрибуты строк сохраняем
    const byRow = new Map();
    for (const [k, cell] of sh.cells) {
      const r = Math.floor(k / 16384);
      if (!byRow.has(r)) byRow.set(r, []);
      byRow.get(r).push([k % 16384, cell]);
    }
    for (const [r, info] of sh.rowsInfo) if (!byRow.has(r) && (info.custom || info.el && (info.el.attributes.length > 1))) byRow.set(r, []);
    const rows = [...byRow.keys()].sort((a, b) => a - b);
    const fresh = this.w(doc, 'sheetData');
    for (const r of rows) {
      const info = sh.rowsInfo.get(r) || {};
      const row = info.el ? info.el.cloneNode(false) : this.w(doc, 'row');
      row.setAttribute('r', r + 1);
      row.removeAttribute('spans');
      if (info.custom) {
        if (info.h) { row.setAttribute('ht', Math.round(info.h * 0.75 * 100) / 100); row.setAttribute('customHeight', '1'); row.removeAttribute('hidden'); }
        else { row.setAttribute('hidden', '1'); }
      }
      const cells = byRow.get(r).sort((a, b) => a[0] - b[0]);
      for (const [c, cell] of cells) row.append(this.cellEl(doc, r, c, cell));
      fresh.append(row);
    }
    data.replaceWith(fresh);
    // Размер заполненной области
    const u = this.usedOf(sh);
    let dim = one('dimension');
    if (!dim) { dim = this.w(doc, 'dimension'); this.wsPlace(doc, dim); }
    dim.setAttribute('ref', u.r || u.c || sh.cells.size ? 'A1:' + addr(u.r, u.c) : 'A1');
    // Ширины столбцов
    if (sh.colsMod) {
      let cols = one('cols');
      const list = sh.cols.filter((k) => k.w != null || k.hidden || k.el);
      if (!list.length) { if (cols) cols.remove(); }
      else {
        const fresh2 = this.w(doc, 'cols');
        for (const k of list) {
          const e = k.el ? k.el.cloneNode(false) : this.w(doc, 'col');
          e.setAttribute('min', k.min + 1);
          e.setAttribute('max', k.max + 1);
          if (k.w != null && !k.hidden) { e.setAttribute('width', Math.round(k.w / XL_MDW * 256) / 256); if (k.mod) e.setAttribute('customWidth', '1'); e.removeAttribute('hidden'); }
          else if (k.hidden) { e.setAttribute('hidden', '1'); if (!e.hasAttribute('width')) e.setAttribute('width', '0'); }
          else if (!e.hasAttribute('width')) e.setAttribute('width', Math.round(sh.defColW / XL_MDW * 256) / 256);
          fresh2.append(e);
        }
        if (cols) cols.replaceWith(fresh2); else this.wsPlace(doc, fresh2);
      }
    }
    // Объединения
    if (sh.mergesMod) {
      let mc = one('mergeCells');
      if (!sh.merges.length) { if (mc) mc.remove(); }
      else {
        const fresh3 = this.w(doc, 'mergeCells');
        fresh3.setAttribute('count', sh.merges.length);
        for (const m of sh.merges) { const e = this.w(doc, 'mergeCell'); e.setAttribute('ref', rangeStr(m)); fresh3.append(e); }
        if (mc) mc.replaceWith(fresh3); else this.wsPlace(doc, fresh3);
      }
    }
    // Диапазоны в условном форматировании, проверке данных, ссылках, фильтре — сдвигаем вслед за строками
    if (sh.ops.length) {
      const fix = (attr) => (e) => {
        const val = e.getAttribute(attr);
        if (!val) return;
        const out = val.split(/\s+/).map((g) => shiftSqref(g, sh.ops)).filter(Boolean);
        if (out.length) e.setAttribute(attr, out.join(' ')); else e.remove();
      };
      for (const e of [...root.getElementsByTagNameNS(XNS.main, 'conditionalFormatting')]) fix('sqref')(e);
      for (const e of [...root.getElementsByTagNameNS(XNS.main, 'dataValidation')]) fix('sqref')(e);
      for (const e of [...root.getElementsByTagNameNS(XNS.main, 'hyperlink')]) fix('ref')(e);
      for (const e of [...root.getElementsByTagNameNS(XNS.main, 'autoFilter')]) fix('ref')(e);
      for (const e of [...root.getElementsByTagNameNS(XNS.main, 'selection')]) { e.setAttribute('activeCell', 'A1'); e.setAttribute('sqref', 'A1'); }
      sh.ops = [];
    }
    return xmlString(doc);
  }

  async buildXlsx() {
    if (this.editing) this.commitEdit();
    const zip = this.zip;
    const ser = xmlString;
    let ctTouched = false, relsTouched = false;
    // Новые и удалённые листы
    const sheetsEl = this.wb.getElementsByTagNameNS(XNS.main, 'sheets')[0];
    for (const sh of this.deletedSheets || []) {
      if (sh.el) sh.el.remove();
      if (sh.rid && this.wbRels[sh.rid]) { this.wbRels[sh.rid].el.remove(); delete this.wbRels[sh.rid]; relsTouched = true; }
      if (sh.path) {
        zip.remove(sh.path);
        const ov = [...this.ct.documentElement.children].find((e) => e.getAttribute('PartName') === '/' + sh.path);
        if (ov) { ov.remove(); ctTouched = true; }
      }
    }
    this.deletedSheets = [];
    let maxId = Math.max(0, ...[...sheetsEl.children].map((e) => +e.getAttribute('sheetId') || 0));
    for (const sh of this.sheets) {
      if (!sh.isNew) continue;
      let n = 1;
      while (zip.file(this.base + `worksheets/sheet${n}.xml`)) n++;
      sh.path = this.base + `worksheets/sheet${n}.xml`;
      let rn = 1;
      while (this.wbRels['rId' + rn]) rn++;
      sh.rid = 'rId' + rn;
      const rel = this.wbRelsXml.createElementNS(XNS.rel, 'Relationship');
      rel.setAttribute('Id', sh.rid);
      rel.setAttribute('Type', XL_REL + 'worksheet');
      rel.setAttribute('Target', `worksheets/sheet${n}.xml`);
      this.wbRelsXml.documentElement.append(rel);
      this.wbRels[sh.rid] = { type: XL_REL + 'worksheet', target: `worksheets/sheet${n}.xml`, el: rel };
      relsTouched = true;
      const ov = this.ct.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override');
      ov.setAttribute('PartName', '/' + sh.path);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      this.ct.documentElement.append(ov);
      ctTouched = true;
      const se = this.wb.createElementNS(XNS.main, 'sheet');
      se.setAttribute('name', sh.name);
      se.setAttribute('sheetId', ++maxId);
      se.setAttributeNS(XNS.r, 'r:id', sh.rid);
      sh.el = se;
      sh.isNew = false;
    }
    if (this.sheetsMod) {
      for (const nm of this.names) {
        const ls = nm.el.getAttribute('localSheetId');
        if (ls == null) continue;
        const sh = (this.origOrder || [])[+ls];
        const ni = this.sheets.indexOf(sh);
        if (ni < 0) nm.el.remove(); else nm.el.setAttribute('localSheetId', ni);
      }
      this.names = this.names.filter((n) => n.el.parentNode);
      this.origOrder = this.sheets.slice();
      // Порядок и имена листов — как сейчас
      for (const sh of this.sheets) { if (sh.el) { sh.el.setAttribute('name', sh.name); sheetsEl.append(sh.el); } }
      // Имена, привязанные к листу (localSheetId), — по новым номерам
      const bv = this.wb.getElementsByTagNameNS(XNS.main, 'workbookView')[0];
      if (bv) { bv.setAttribute('activeTab', Math.max(0, this.activeIdx)); bv.removeAttribute('firstSheet'); }
    }
    for (const nm of this.names) if (nm.mod) { nm.el.textContent = nm.text; nm.mod = false; }
    // Листы
    let anyFormula = false;
    for (const sh of this.sheets) {
      if (!sh.cells) continue;
      if (!sh.touched && !sh.colsMod && !sh.mergesMod) continue;
      for (const c of sh.cells.values()) if (c.f != null && c.mod) anyFormula = true;
      zip.file(sh.path, this.writeSheet(sh));
      for (const c of sh.cells.values()) if (c.mod) { c.mod = false; }
      sh.touched = false;
      sh.colsMod = false;
      sh.mergesMod = false;
    }
    // Порядок вычислений Excel пересоберёт сам; заодно попросим пересчитать формулы при открытии
    const calcRel = Object.entries(this.wbRels).find(([, r]) => r.type.endsWith('/calcChain'));
    if (calcRel) {
      const p = this.resolve(calcRel[1].target, this.base);
      zip.remove(p);
      calcRel[1].el.remove();
      delete this.wbRels[calcRel[0]];
      relsTouched = true;
      const ov = [...this.ct.documentElement.children].find((e) => e.getAttribute('PartName') === '/' + p);
      if (ov) { ov.remove(); ctTouched = true; }
    }
    if (anyFormula || this.stale) {
      let cp = this.wb.getElementsByTagNameNS(XNS.main, 'calcPr')[0];
      if (!cp) {
        cp = this.wb.createElementNS(XNS.main, 'calcPr');
        const after = ['definedNames', 'sheets'].map((n) => this.wb.getElementsByTagNameNS(XNS.main, n)[0]).find(Boolean);
        after.after(cp);
      }
      cp.setAttribute('fullCalcOnLoad', '1');
    }
    zip.file(this.wbPath, ser(this.wb));
    if (relsTouched) zip.file(this.wbRelsPath, ser(this.wbRelsXml));
    if (ctTouched) zip.file('[Content_Types].xml', ser(this.ct));
    if (this.sstTouched) {
      const root = this.sstXml.documentElement;
      root.setAttribute('uniqueCount', this.sst.length);
      root.setAttribute('count', Math.max(+root.getAttribute('count') || 0, this.sst.length));
      zip.file(this.sstPath, ser(this.sstXml));
      this.sstTouched = false;
    }
    if (this.stylesTouched) { zip.file(this.stylesPath, ser(this.stylesXml)); this.stylesTouched = false; }
    this.sheetsMod = false;
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // ---------- В PDF ----------
  async buildPrintHtml() {
    if (this.editing) this.commitEdit();
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const parts = [];
    const pageW = 794 - 2 * 48; // A4, поля ~13 мм
    this.sheets.forEach((sh) => {
      if (!sh.cells || sh.state !== 'visible') return;
      const u = this.usedOf(sh);
      if (!sh.cells.size) return;
      const R = u.r, C = u.c;
      let html = '<table class="xp"><colgroup>';
      let tw = 0;
      for (let c = 0; c <= C; c++) { const w = this.colW(c, sh); tw += w; html += `<col style="width:${w}px">`; }
      html += '</colgroup>';
      const skip = new Set();
      for (let r = 0; r <= R; r++) {
        const h = this.rowH(r, sh);
        if (!h) continue;
        html += `<tr style="height:${h}px">`;
        for (let c = 0; c <= C; c++) {
          if (skip.has(xkey(r, c))) continue;
          if (!this.colW(c, sh)) continue;
          const m = sh.mergeAt.get(xkey(r, c));
          let span = '';
          if (m && m.r1 === r && m.c1 === c) {
            span = ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"`;
            for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) skip.add(xkey(rr, cc));
          }
          const cell = sh.cells.get(xkey(r, c));
          const st = this.st(cell ? cell.s : 0);
          const d = cell ? this.display(cell) : null;
          const css = [];
          const f = st.font;
          css.push(`font:${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${f.sz || 11}pt "${f.name || 'Calibri'}",Calibri,Arial,sans-serif`);
          if (d && (d.color || f.color)) css.push('color:' + (d.color || f.color));
          if (st.fill) css.push('background:' + st.fill);
          let h2 = st.align.h;
          if (!h2 || h2 === 'general') h2 = d && d.num ? 'right' : d && (d.bool || d.err) ? 'center' : 'left';
          css.push('text-align:' + (h2 === 'centerContinuous' ? 'center' : h2));
          css.push('vertical-align:' + ({ top: 'top', center: 'middle' }[st.align.v] || 'bottom'));
          if (st.align.wrap) css.push('white-space:pre-wrap;word-wrap:break-word'); else css.push('white-space:pre;overflow:hidden');
          if (f.u) css.push('text-decoration:underline');
          const bd = st.border;
          const side = (s) => (s ? `${{ medium: 2, thick: 3, double: 3 }[s.style] || 1}px ${s.style === 'double' ? 'double' : /dash/i.test(s.style) ? 'dashed' : /dot/i.test(s.style) ? 'dotted' : 'solid'} ${s.color}` : null);
          for (const [k, n] of [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom']]) { const v = side(bd[k]); if (v) css.push(`border-${n}:${v}`); }
          html += `<td${span} style="${css.join(';')}">${d ? esc(d.text) : ''}</td>`;
        }
        html += '</tr>';
      }
      html += '</table>';
      const land = tw > pageW * 1.15;
      const avail = land ? 1123 - 2 * 48 : pageW;
      const k = Math.min(1, avail / tw);
      parts.push(`<section class="${land ? 'land' : 'port'}" style="zoom:${k.toFixed(3)}">${html}</section>`);
    });
    const css = `@page { size: A4; margin: 13mm; } @page land { size: A4 landscape; margin: 13mm; }
html, body { margin: 0; background: #fff; }
section { break-after: page; } section:last-child { break-after: auto; } section.land { page: land; }
table.xp { border-collapse: collapse; table-layout: fixed; }
table.xp td { padding: 0 3px; overflow: hidden; line-height: 1.15; }
table.xp tr { break-inside: avoid; }`;
    return `<!doctype html><html><head><meta charset="utf-8"><style>${css}</style></head><body>${parts.join('') || '<section></section>'}</body></html>`;
  }

  async save(asNew) {
    if (this.editing && !this.commitEdit()) return false;
    let target = this.path && /\.xlsx$/i.test(this.path) && !asNew ? this.path : null;
    if (!target) {
      const base = (this.path || T('Таблица')).replace(/\.[^.\\/]+$/, '');
      target = await window.folio.saveDialog(base + '.xlsx', [
        { name: T('Таблица Excel'), extensions: ['xlsx'] },
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
    const bytes = await this.buildXlsx();
    await window.folio.writeFile(target, bytes);
    this.path = target;
    this.dirty = false;
    this.onChange && this.onChange();
    return true;
  }

  endEdit() { if (this.editing) this.commitEdit(); }

  isShown() { return !this.root.hidden && !!this.root.offsetParent; }

  destroy() {
    document.removeEventListener('copy', this.onCopy);
    document.removeEventListener('cut', this.onCut);
    document.removeEventListener('paste', this.onPaste);
    this.ro.disconnect();
    if (XlsxEditor.clip && XlsxEditor.clip.from === this) XlsxEditor.clip = null;
    this.root.remove();
  }
}
const XL_REF_COLORS = ['#3a5ccc', '#c0392b', '#7d3c98', '#16a085', '#d35400', '#2c3e50'];

// Сдвиг диапазона вида A1:B5 по журналу вставок/удалений.
function shiftSqref(g, ops) {
  let rg = parseRange(g);
  if (!rg) return g;
  for (const { axis, at, n } of ops) {
    const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
    if (n > 0) { if (rg[k1] >= at) { rg[k1] += n; rg[k2] += n; } else if (rg[k2] >= at) rg[k2] += n; continue; }
    const d = -n, end = at + d - 1;
    if (rg[k1] >= at && rg[k2] <= end) return null;
    rg[k1] = rg[k1] > end ? rg[k1] - d : rg[k1] >= at ? at : rg[k1];
    rg[k2] = rg[k2] > end ? rg[k2] - d : rg[k2] >= at ? at - 1 : rg[k2];
  }
  return rangeStr(rg);
}

function parseTsv(text) {
  const rows = [];
  let row = [], cur = '', q = false;
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') { q = true; continue; }
    if (ch === '\t') { row.push(cur); cur = ''; continue; }
    if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    cur += ch;
  }
  row.push(cur);
  rows.push(row);
  return rows;
}

// Пустая книга: один лист, Calibri 11.
async function makeBlankXlsx() {
  const zip = new JSZip();
  const M = XNS.main;
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${XL_REL}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${XL_REL}styles" Target="styles.xml"/><Relationship Id="rId3" Type="${XL_REL}sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${M}" xmlns:r="${XNS.r}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="${T('Лист')}1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${M}" xmlns:r="${XNS.r}"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${M}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`);
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${M}" count="0" uniqueCount="0"/>`);
  return zip.generateAsync({ type: 'uint8array' });
}
