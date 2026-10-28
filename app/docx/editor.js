// Редактор Word (.docx).
// Главное правило: в файл возвращается только то, что тронули. Абзацы и таблицы,
// которых не касались, уходят обратно исходным XML — со всеми колонтитулами, полями,
// сносками и прочим, что редактор сам не понимает.

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  v: 'urn:schemas-microsoft-com:vml',
};
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const TWIP = 1 / 20; // пунктов в твипе
const PXPT = 4 / 3; // пикселей в пункте
const PAGE_GAP = 22; // промежуток между листами, px

const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];

const WD_HL_MAP = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff' };
const WD_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Verdana', 'Tahoma', 'Courier New'];
const WD_COLORS = ['#000000', '#595959', '#C00000', '#E36C09', '#00B050', '#0070C0', '#7030A0', '#1F3864'];
const WD_MARKS = [['yellow', '#ffff00'], ['green', '#00ff00'], ['cyan', '#00ffff'], ['none', null]];
const STYLE_RU = { normal: 'Обычный', title: 'Название', subtitle: 'Подзаголовок', 'list paragraph': 'Абзац списка', caption: 'Подпись к рисунку', 'no spacing': 'Без интервала', quote: 'Цитата', 'toc heading': 'Заголовок оглавления' };
const WD_STYLES = [['normal', 'Обычный'], ['title', 'Название'], ['heading 1', 'Заголовок 1'], ['heading 2', 'Заголовок 2'], ['heading 3', 'Заголовок 3']];

// Стили документа, которые общие для экрана и для печати в PDF.
const WD_CSS = `
.wd-surface { color: #000; font-family: "Times New Roman", serif; font-size: 10pt; line-height: 1.15; }
.wd-surface p { margin: 0; white-space: pre-wrap; tab-size: 36pt; overflow-wrap: break-word; min-height: 1em; }
.wd-surface p[data-marker]::before { content: attr(data-marker); display: inline-block; min-width: var(--hang, 18pt); text-indent: 0; text-decoration: none; }
.wd-surface table { border-collapse: collapse; table-layout: fixed; margin: 0; }
.wd-surface td { vertical-align: top; padding: 0 5.4pt; overflow-wrap: break-word; }
.wd-surface table.bordered > tbody > tr > td { border: 0.5pt solid #000; }
.wd-surface .atom img { display: inline-block; vertical-align: bottom; max-width: 100%; }
.wd-surface .atom.hid { display: none; }
.wd-surface .atom.link { color: #0563C1; text-decoration: underline; }
.wd-surface .atom.pbreak { display: block; height: 0; }
.wd-surface .noimg { display: inline-flex; align-items: center; justify-content: center; background: #f0f1f3; color: #8a9099; font: 9pt sans-serif; border: 1px dashed #c9cdd3; }
.wd-surface .txbx { display: inline-block; border: 1px dashed #c9cdd3; padding: 2pt 4pt; white-space: pre-wrap; }
`;

const wkids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
const wkid = (el, name) => (el ? [...el.children].find((c) => c.localName === name) || null : null);
const wa = (el, name) => (el ? el.getAttributeNS(NS.w, name) : null);
const onoff = (el) => !!el && !['0', 'false', 'off'].includes(wa(el, 'val'));
const firstFam = (f) => (f || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
const famCss = (n) => `"${n}", ${/arial|calibri|verdana|tahoma|helvetica|segoe|aptos|sans/i.test(n) ? 'sans-serif' : /courier|consol|mono/i.test(n) ? 'monospace' : 'serif'}`;
const rgbHex = (c) => {
  const m = c.match(/\d+(\.\d+)?/g) || [0, 0, 0];
  return m.slice(0, 3).map((v) => Math.round(+v).toString(16).padStart(2, '0')).join('').toUpperCase();
};
const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
function fmtNum(fmt, n) {
  const letters = (abc) => { let s = ''; let k = n; while (k > 0) { k--; s = abc[k % abc.length] + s; k = Math.floor(k / abc.length); } return s; };
  const roman = () => { let s = '', k = n; for (const [v, r] of ROMAN) while (k >= v) { s += r; k -= v; } return s; };
  switch (fmt) {
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'lowerLetter': return letters('abcdefghijklmnopqrstuvwxyz');
    case 'upperLetter': return letters('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
    case 'lowerRoman': return roman();
    case 'upperRoman': return roman().toUpperCase();
    case 'russianLower': return letters('абвгдежзиклмнопрстуфхцчшщэюя');
    case 'russianUpper': return letters('АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ');
    case 'none': case 'bullet': return '';
    default: return String(n);
  }
}
const BULLETS = { '': '•', '': '▪', '': '➢', '': '✓', '': '❖', '': '◦', o: '◦', '': '' };
// XML в строку с одним заголовком <?xml?> (браузер иногда добавляет свой).
const xmlString = (x) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + new XMLSerializer().serializeToString(x).replace(/^<\?xml[^>]*\?>\s*/, '');
const mapBullet = (t) => (t in BULLETS ? BULLETS[t] : /[-]/.test(t) ? '•' : t);

class DocxEditor {
  constructor(host, { onChange }) {
    this.kind = 'docx';
    this.onChange = onChange;
    this.dirty = false;
    this.zoom = 1;
    this.changed = new WeakSet();
    this.tails = new WeakMap();
    this.newImgs = new Map();
    this.listIds = {};

    this.root = el('div', 'doc wd');
    this.viewer = el('div', 'viewer wd-viewer');
    this.viewer.tabIndex = -1;
    this.col = el('div', 'wd-col');
    this.scaleBox = el('div', 'wd-scale');
    this.wrap = el('div', 'wd-wrap');
    this.sheets = el('div', 'wd-sheets');
    this.surface = el('div', 'wd-surface');
    this.surface.contentEditable = 'true';
    this.surface.spellcheck = false;
    this.wrap.append(this.sheets, this.surface);
    this.scaleBox.append(this.wrap);
    this.col.append(this.scaleBox);
    this.viewer.append(this.col);
    this.hud = el('div', 'hud');
    this.side = el('aside', 'side');
    this.sideGrip = el('div', 'side-grip');
    this.sideGrip.dataset.tip = 'Перетащите панель влево или вправо';
    this.thumbs = el('div', 'thumbs');
    this.side.append(this.sideGrip, this.thumbs);
    this.root.append(this.side, this.viewer, this.hud);
    this.thumbEls = [];
    this.thumbs.addEventListener('scroll', () => this.renderVisibleThumbs(), { passive: true });
    this.gapEls = [];
    this.thumbs.addEventListener('click', (e) => {
      const g = e.target.closest('.gap button');
      if (g) { this.insertPageAt(+g.parentNode.dataset.at); return; }
      const t = e.target.closest('.thumb');
      if (t) this.scrollToPage(+t.dataset.i);
    });
    host.append(this.root);

    this.buildDock();
    this.bindSurface();
    this.bindSideGrip();
    this.ro = new ResizeObserver(() => {
      if (!this.viewer.clientWidth || !this.pg) return;
      if (this.fitPending) { this.fitPending = false; this.fitWidth(); }
      else if (this.fitMode) this.fitWidth();
    });
    this.ro.observe(this.viewer);
  }

  // ---------- Загрузка ----------
  async open(files) {
    const loading = el('div', 'loading', 'Открываю…');
    this.root.append(loading);
    try {
      await this.load(files[0].bytes);
      await this.render();
    } finally {
      loading.remove();
    }
  }

  async blank() {
    await this.load(await makeBlankDocx());
    await this.render();
    this.placeCaretStart();
  }

  async readXml(path) {
    const f = this.zip.file(path);
    if (!f) return null;
    const x = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    if (x.getElementsByTagName('parsererror').length) throw new Error('Повреждённый файл: ' + path);
    return x;
  }

  resolvePath(target, base = this.base) {
    if (target.startsWith('/')) return target.slice(1);
    const out = [];
    for (const s of (base + target).split('/')) {
      if (s === '..') out.pop();
      else if (s && s !== '.') out.push(s);
    }
    return out.join('/');
  }

  readRels(x) {
    const map = {};
    for (const r of x ? x.documentElement.children : []) map[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: r.getAttribute('Target') || '', ext: r.getAttribute('TargetMode') === 'External' };
    return map;
  }

  async load(bytes) {
    this.zip = await JSZip.loadAsync(bytes);
    const rootRels = this.readRels(await this.readXml('_rels/.rels'));
    const main = Object.values(rootRels).find((r) => r.type.endsWith('/officeDocument'));
    this.mainPath = main ? main.target.replace(/^\//, '') : 'word/document.xml';
    this.base = this.mainPath.slice(0, this.mainPath.lastIndexOf('/') + 1);
    this.xml = await this.readXml(this.mainPath);
    if (!this.xml) throw new Error('Это не документ Word');
    this.relsPath = this.base + '_rels/' + this.mainPath.slice(this.base.length) + '.rels';
    this.relsXml = await this.readXml(this.relsPath);
    this.rels = this.readRels(this.relsXml);
    this.ct = await this.readXml('[Content_Types].xml');
    const part = (type) => {
      const r = Object.values(this.rels).find((x) => x.type.endsWith('/' + type));
      return r ? this.resolvePath(r.target) : null;
    };
    this.stylesPath = part('styles');
    this.stylesXml = this.stylesPath ? await this.readXml(this.stylesPath) : null;
    this.numPath = part('numbering');
    this.numXml = this.numPath ? await this.readXml(this.numPath) : null;
    const themePath = part('theme');
    const theme = themePath ? await this.readXml(themePath) : null;
    this.themeFonts = {};
    if (theme) {
      const lat = (n) => {
        const f = theme.getElementsByTagNameNS(NS.a, n)[0];
        const l = f && [...f.children].find((c) => c.localName === 'latin');
        return l ? l.getAttribute('typeface') : null;
      };
      this.themeFonts = { major: lat('majorFont'), minor: lat('minorFont') };
    }
    this.mediaCache = new Map();
    this.parseStyles();
    this.parseNumbering();
  }

  // ---------- Стили и нумерация ----------
  parseRPr(rPr) {
    const o = {};
    if (!rPr) return o;
    for (const c of rPr.children) {
      const v = wa(c, 'val');
      switch (c.localName) {
        case 'rStyle': o.rStyle = v; break;
        case 'b': o.b = onoff(c); break;
        case 'i': o.i = onoff(c); break;
        case 'u': o.u = !!v && v !== 'none'; break;
        case 'strike': case 'dstrike': o.strike = onoff(c); break;
        case 'caps': o.caps = onoff(c); break;
        case 'smallCaps': o.smallCaps = onoff(c); break;
        case 'vanish': o.hidden = onoff(c); break;
        case 'sz': if (v) o.sz = +v / 2; break;
        case 'color': o.color = v && v !== 'auto' ? '#' + v : null; break;
        case 'highlight': o.hl = v && v !== 'none' ? WD_HL_MAP[v] || null : null; break;
        case 'shd': { const f = wa(c, 'fill'); if (f && f !== 'auto') o.shd = '#' + f; break; }
        case 'vertAlign': o.va = v; break;
        case 'rFonts': {
          const t = wa(c, 'asciiTheme') || wa(c, 'hAnsiTheme');
          const f = (t && this.themeFonts[t.startsWith('major') ? 'major' : 'minor']) || wa(c, 'ascii') || wa(c, 'hAnsi') || wa(c, 'cs');
          if (f) o.font = f;
          break;
        }
      }
    }
    return o;
  }

  parsePPr(pPr) {
    const o = {};
    if (!pPr) return o;
    for (const c of pPr.children) {
      switch (c.localName) {
        case 'pStyle': o.style = wa(c, 'val'); break;
        case 'jc': {
          const v = wa(c, 'val');
          o.jc = v === 'center' ? 'center' : v === 'right' || v === 'end' ? 'right' : v === 'both' || v === 'distribute' ? 'justify' : 'left';
          break;
        }
        case 'ind': {
          const n = (a) => { const v = wa(c, a); return v == null ? null : +v * TWIP; };
          const l = n('left') ?? n('start');
          const r = n('right') ?? n('end');
          const f = n('firstLine'), h = n('hanging');
          if (l != null) o.indL = l;
          if (r != null) o.indR = r;
          if (h != null) o.indFirst = -h; else if (f != null) o.indFirst = f;
          break;
        }
        case 'spacing': {
          const b = wa(c, 'before'), a = wa(c, 'after'), l = wa(c, 'line');
          if (b != null && wa(c, 'beforeAutospacing') !== '1') o.before = +b * TWIP;
          if (a != null && wa(c, 'afterAutospacing') !== '1') o.after = +a * TWIP;
          if (l != null) { o.line = +l; o.lineRule = wa(c, 'lineRule') || 'auto'; }
          break;
        }
        case 'numPr': {
          const id = wa(wkid(c, 'numId'), 'val'), lv = wa(wkid(c, 'ilvl'), 'val');
          if (id != null) o.numId = id;
          if (lv != null) o.ilvl = +lv;
          break;
        }
        case 'pageBreakBefore': o.pbb = onoff(c); break;
        case 'shd': { const f = wa(c, 'fill'); if (f && f !== 'auto') o.shd = '#' + f; break; }
        case 'sectPr': o.sect = wa(wkid(c, 'type'), 'val') || 'nextPage'; break;
      }
    }
    return o;
  }

  parseStyles() {
    this.styles = {};
    this.styleCache = new Map();
    this.defR = {};
    this.defP = {};
    this.defPStyle = null;
    const root = this.stylesXml && this.stylesXml.documentElement;
    if (!root) return;
    const dd = wkid(root, 'docDefaults');
    this.defR = this.parseRPr(wkid(wkid(dd, 'rPrDefault'), 'rPr'));
    this.defP = this.parsePPr(wkid(wkid(dd, 'pPrDefault'), 'pPr'));
    for (const s of wkids(root, 'style')) {
      const id = wa(s, 'styleId');
      const tblB = wkid(wkid(s, 'tblPr'), 'tblBorders');
      this.styles[id] = {
        id, type: wa(s, 'type'), name: (wa(wkid(s, 'name'), 'val') || id).toLowerCase(),
        basedOn: wa(wkid(s, 'basedOn'), 'val'),
        pPr: this.parsePPr(wkid(s, 'pPr')), rPr: this.parseRPr(wkid(s, 'rPr')),
        borders: tblB ? this.bordersOn(tblB) : undefined,
      };
      if (wa(s, 'type') === 'paragraph' && ['1', 'true'].includes(wa(s, 'default'))) this.defPStyle = id;
    }
  }

  // Свойства стиля с учётом всей цепочки «основан на».
  style(id) {
    if (!id) return { pPr: {}, rPr: {}, borders: false };
    if (this.styleCache.has(id)) return this.styleCache.get(id);
    const chain = [];
    for (let s = this.styles[id], n = 0; s && n < 20; s = this.styles[s.basedOn], n++) chain.unshift(s);
    const out = { pPr: {}, rPr: {}, borders: false };
    for (const s of chain) {
      Object.assign(out.pPr, s.pPr);
      Object.assign(out.rPr, s.rPr);
      if (s.borders !== undefined) out.borders = s.borders;
    }
    this.styleCache.set(id, out);
    return out;
  }

  styleIdByName(name) {
    const s = Object.values(this.styles).find((x) => x.name === name && x.type === 'paragraph');
    return s ? s.id : null;
  }

  bordersOn(b) {
    return [...b.children].some((c) => !['nil', 'none'].includes(wa(c, 'val')));
  }

  parseNumbering() {
    this.abs = {};
    this.nums = {};
    const root = this.numXml && this.numXml.documentElement;
    if (!root) return;
    for (const a of wkids(root, 'abstractNum')) {
      const lvls = {};
      for (const l of wkids(a, 'lvl')) {
        const pp = this.parsePPr(wkid(l, 'pPr'));
        lvls[+wa(l, 'ilvl')] = {
          fmt: wa(wkid(l, 'numFmt'), 'val') || 'decimal', text: wa(wkid(l, 'lvlText'), 'val') ?? '',
          start: +(wa(wkid(l, 'start'), 'val') ?? 1), indL: pp.indL, indFirst: pp.indFirst,
        };
      }
      this.abs[wa(a, 'abstractNumId')] = lvls;
    }
    for (const n of wkids(root, 'num')) {
      const ov = {};
      for (const o of wkids(n, 'lvlOverride')) {
        const s = wa(wkid(o, 'startOverride'), 'val');
        if (s != null) ov[+wa(o, 'ilvl')] = +s;
      }
      this.nums[wa(n, 'numId')] = { abs: wa(wkid(n, 'abstractNumId'), 'val'), ov };
    }
  }

  lvl(numId, ilvl) {
    const n = this.nums[numId];
    return n && this.abs[n.abs] ? this.abs[n.abs][ilvl] || null : null;
  }

  // ---------- Отрисовка ----------
  async render() {
    const body = wkid(this.xml.documentElement, 'body');
    this.store = [];
    this.rprStore = [];
    this.rprIdx = new Map();
    this.atoms = [];
    this.directCache = new Map();
    this.lead = [];
    this.finalSect = null;
    this.sections = [];
    this.ctxRels = this.rels;
    this.ctxBase = this.base;
    this.surface.textContent = '';
    const frag = document.createDocumentFragment();
    let prev = null;
    for (const c of body.children) {
      let node = null;
      if (c.localName === 'p') node = this.renderP(c);
      else if (c.localName === 'tbl') node = this.renderTbl(c);
      else if (c.localName === 'sdt') node = this.renderSdt(c);
      else if (c.localName === 'sectPr') { this.finalSect = c; continue; }
      else {
        const i = this.atoms.push([c]) - 1;
        if (prev) this.tails.set(prev, [...(this.tails.get(prev) || []), i]);
        else this.lead.push(i);
        continue;
      }
      frag.append(node);
      prev = node;
    }
    if (!frag.childNodes.length) frag.append(this.emptyP());
    this.surface.append(frag);
    this.sections.push(this.parseSect(this.finalSect));
    this.setupSections();
    await this.loadHeaderFooter();
    this.renumber();
    this.observe();
    this.fitPending = true;
    if (this.viewer.clientWidth) { this.fitPending = false; this.fitWidth(); }
    this.paginate();
    // Шрифты Windows подгружаются не сразу: когда догрузятся — пересчитываем листы.
    const again = () => { clearTimeout(this.fontTimer); this.fontTimer = setTimeout(() => this.paginate(), 120); };
    if (!this.onFonts) {
      this.onFonts = again;
      document.fonts.addEventListener('loadingdone', this.onFonts);
    }
    document.fonts.ready.then(again);
    setTimeout(again, 900);
    for (const img of this.surface.querySelectorAll('img')) if (!img.complete) img.addEventListener('load', again, { once: true });
  }

  emptyP() {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    this.applyPara(p);
    return p;
  }

  directPPr(k) {
    if (!this.directCache.has(k)) this.directCache.set(k, this.parsePPr(wkid(this.store[k], 'pPr')));
    return this.directCache.get(k);
  }

  renderP(pEl) {
    const k = this.store.push(pEl) - 1;
    const d = this.directPPr(k);
    const p = document.createElement('p');
    p.dataset.k = k;
    if (d.style) p.dataset.style = d.style;
    if (d.jc) p.dataset.jc = d.jc;
    const st = this.style(d.style || this.defPStyle);
    const numId = d.numId ?? st.pPr.numId;
    if (numId && numId !== '0') p.dataset.num = `${numId}:${d.ilvl ?? st.pPr.ilvl ?? 0}`;
    if (d.pbb || st.pPr.pbb) p.dataset.pbb = '1';
    if (d.sect && this.sections) {
      p.dataset.sect = this.sections.length;
      this.sections.push(this.parseSect(wkid(wkid(pEl, 'pPr'), 'sectPr')));
    }
    this.applyPara(p);
    this.renderInline([...pEl.children], p);
    // Последний <br> в абзаце невидим, поэтому для пустой строки в конце нужен ещё один.
    let leaf = p.lastChild;
    while (leaf && leaf.nodeType === 1 && leaf.lastChild && !leaf.classList.contains('atom')) leaf = leaf.lastChild;
    if (!leaf || leaf.nodeName === 'BR' || (!p.textContent && !p.querySelector('.atom:not(.hid)'))) p.append(document.createElement('br'));
    return p;
  }

  // Оформление абзаца: стиль + прямые свойства + то, что поменяли в редакторе.
  applyPara(p) {
    const sid = p.dataset.style || this.defPStyle;
    const st = this.style(sid);
    const d = p.dataset.k != null ? this.directPPr(+p.dataset.k) : {};
    const pp = { ...this.defP, ...st.pPr, ...d };
    pp.jc = p.dataset.jc || { ...this.defP, ...st.pPr }.jc || 'left';
    if (p.dataset.num) {
      const [nid, il] = p.dataset.num.split(':');
      const lv = this.lvl(nid, +il);
      if (lv) {
        if (d.indL == null && lv.indL != null) pp.indL = lv.indL;
        if (d.indFirst == null && lv.indFirst != null) pp.indFirst = lv.indFirst;
      }
    }
    const rp = { ...this.defR, ...st.rPr };
    const s = p.style;
    s.textAlign = pp.jc;
    s.paddingLeft = (pp.indL || 0) + 'pt';
    s.paddingRight = (pp.indR || 0) + 'pt';
    s.textIndent = (pp.indFirst || 0) + 'pt';
    s.paddingTop = (pp.before || 0) + 'pt';
    s.paddingBottom = (pp.after || 0) + 'pt';
    s.lineHeight = !pp.line ? '1.15' : pp.lineRule === 'auto' ? ((pp.line / 240) * 1.15).toFixed(3) : pp.line * TWIP + 'pt';
    s.fontSize = (rp.sz || 10) + 'pt';
    s.fontFamily = famCss(rp.font || 'Times New Roman');
    s.fontWeight = rp.b ? '700' : '400';
    s.fontStyle = rp.i ? 'italic' : 'normal';
    s.color = rp.color || '#000';
    s.textDecoration = [rp.u && 'underline', rp.strike && 'line-through'].filter(Boolean).join(' ') || 'none';
    s.textTransform = rp.caps ? 'uppercase' : '';
    s.background = pp.shd || '';
    if (p.dataset.num) s.setProperty('--hang', Math.max(10, -(pp.indFirst ?? -18)) + 'pt');
    else s.removeProperty('--hang');
  }

  runProps(rPr) {
    const own = this.parseRPr(rPr);
    return own.rStyle ? { ...this.style(own.rStyle).rPr, ...own } : own;
  }

  runCss(o) {
    const c = [];
    if (o.b != null) c.push(`font-weight:${o.b ? 700 : 400}`);
    if (o.i != null) c.push(`font-style:${o.i ? 'italic' : 'normal'}`);
    if (o.u || o.strike) c.push(`text-decoration:${[o.u && 'underline', o.strike && 'line-through'].filter(Boolean).join(' ')}`);
    if (o.sz) c.push(`font-size:${o.va === 'superscript' || o.va === 'subscript' ? o.sz * 0.65 : o.sz}pt`);
    if (o.font) c.push(`font-family:${famCss(o.font)}`);
    if (o.color !== undefined) c.push(`color:${o.color || '#000'}`);
    if (o.hl || o.shd) c.push(`background-color:${o.hl || o.shd}`);
    if (o.va === 'superscript') c.push('vertical-align:super');
    if (o.va === 'subscript') c.push('vertical-align:sub');
    if (o.caps) c.push('text-transform:uppercase');
    if (o.smallCaps) c.push('font-variant:small-caps');
    return c.join(';');
  }

  rprIndex(rPr) {
    const key = new XMLSerializer().serializeToString(rPr).replace(/\s?w:rsid\w*="[^"]*"/g, '');
    if (!this.rprIdx.has(key)) this.rprIdx.set(key, this.rprStore.push(rPr) - 1);
    return this.rprIdx.get(key);
  }

  pushAtom(nodes, target, disp, cls = '') {
    const i = this.atoms.push(nodes) - 1;
    const s = document.createElement('span');
    s.className = 'atom' + (cls ? ' ' + cls : '');
    s.contentEditable = 'false';
    s.dataset.x = i;
    if (disp) s.append(...disp.childNodes);
    target.append(s);
    return s;
  }

  renderInline(nodes, target) {
    let fld = null;
    const finishField = () => {
      const disp = document.createElement('span');
      const instr = fld.instr.trim();
      if (/^PAGE\b/i.test(instr)) disp.append(el('span', 'fld-page', '1'));
      else if (/^NUMPAGES\b/i.test(instr)) disp.append(el('span', 'fld-pages', '1'));
      else for (const r of fld.result) if (r.localName === 'r') this.fillRunDisplay(r, disp); else this.renderInline([r], disp);
      this.pushAtom(fld.nodes, target, disp, 'fld');
      fld = null;
    };
    for (const c of nodes) {
      const name = c.localName;
      if (fld) {
        fld.nodes.push(c);
        // Поля бывают вложенными (номер страницы внутри ссылки оглавления):
        // показываем фрагмент, только если он в «результате» на всех уровнях.
        const fc = name === 'r' ? wkid(c, 'fldChar') : null;
        if (fc) {
          const t = wa(fc, 'fldCharType');
          if (t === 'begin') fld.stack.push(false);
          else if (t === 'separate') fld.stack[fld.stack.length - 1] = true;
          else if (t === 'end') { fld.stack.pop(); if (!fld.stack.length) finishField(); }
        } else {
          const it = name === 'r' ? wkid(c, 'instrText') : null;
          if (it) { if (fld.stack.length === 1) fld.instr += it.textContent; }
          else if (fld.stack.every(Boolean)) fld.result.push(c);
        }
        continue;
      }
      switch (name) {
        case 'r': {
          const fc = wkid(c, 'fldChar');
          if (fc && wa(fc, 'fldCharType') === 'begin') { fld = { nodes: [c], stack: [false], instr: '', result: [] }; break; }
          if (fc || wkid(c, 'instrText')) { this.pushAtom([c], target, null, 'hid'); break; }
          this.renderRun(c, target);
          break;
        }
        case 'hyperlink': case 'fldSimple': case 'sdt': case 'oMath': case 'oMathPara': {
          const disp = document.createElement('span');
          const inner = name === 'sdt' ? [...(wkid(c, 'sdtContent') || { children: [] }).children] : [...c.children];
          if (name === 'oMath' || name === 'oMathPara') disp.textContent = [...c.getElementsByTagNameNS('*', 't')].map((t) => t.textContent).join('');
          else if (name === 'fldSimple' && /^\s*PAGE\b/i.test(wa(c, 'instr') || '')) disp.append(el('span', 'fld-page', '1'));
          else this.renderInline(inner, disp);
          this.pushAtom([c], target, disp, name === 'hyperlink' ? 'link' : '');
          break;
        }
        case 'smartTag': case 'customXml': case 'ins': case 'moveTo':
          this.renderInline([...c.children], target);
          break;
        case 'pPr': case 'bookmarkStart': case 'bookmarkEnd': case 'proofErr': case 'permStart': case 'permEnd':
          break;
        default:
          this.pushAtom([c], target, null, 'hid'); // неизвестное — не показываем, но сохраняем
      }
    }
    if (fld) finishField(); // поле продолжается в следующем абзаце (оглавление)
  }

  renderRun(r, target) {
    const rPr = wkid(r, 'rPr');
    const props = this.runProps(rPr);
    const special = [...r.children].some((c) =>
      ['drawing', 'pict', 'object', 'footnoteReference', 'endnoteReference', 'commentReference', 'annotationRef', 'ptab', 'sym', 'separator', 'continuationSeparator'].includes(c.localName) ||
      (c.localName === 'br' && wa(c, 'type') === 'page'));
    if (props.hidden) { this.pushAtom([r], target, null, 'hid'); return; }
    if (special) {
      const disp = document.createElement('span');
      this.fillRunDisplay(r, disp);
      const isPb = [...r.children].some((c) => c.localName === 'br' && wa(c, 'type') === 'page');
      this.pushAtom([r], target, disp, isPb ? 'pbreak' : '');
      return;
    }
    let host = target;
    if (rPr && rPr.children.length) {
      const idx = String(this.rprIndex(rPr));
      const last = target.lastChild;
      if (last && last.nodeType === 1 && last.tagName === 'SPAN' && last.dataset.r === idx) host = last;
      else {
        host = document.createElement('span');
        host.dataset.r = idx;
        host.style.cssText = this.runCss(props);
        target.append(host);
      }
    }
    const text = (s) => {
      if (host.lastChild && host.lastChild.nodeType === 3) host.lastChild.appendData(s);
      else host.append(s);
    };
    for (const c of r.children) {
      switch (c.localName) {
        case 't': text(c.textContent); break;
        case 'tab': text('\t'); break;
        case 'br': case 'cr': host.append(document.createElement('br')); break;
        case 'noBreakHyphen': text('‑'); break;
      }
    }
  }

  // Что показать на месте «неразборного» фрагмента: картинку, номер сноски, текст.
  fillRunDisplay(r, disp) {
    const props = this.runProps(wkid(r, 'rPr'));
    const box = document.createElement('span');
    box.style.cssText = this.runCss(props);
    for (const c of r.children) {
      switch (c.localName) {
        case 't': box.append(c.textContent); break;
        case 'tab': box.append('\t'); break;
        case 'drawing': box.append(this.drawingEl(c)); break;
        case 'pict': case 'object': box.append(this.vmlEl(c)); break;
        case 'footnoteReference': case 'endnoteReference': {
          const s = document.createElement('sup');
          s.textContent = wa(c, 'id') || '*';
          box.append(s);
          break;
        }
        case 'sym': box.append('•'); break;
      }
    }
    disp.append(box);
  }

  media(rid) {
    const rels = this.ctxRels, base = this.ctxBase;
    const rel = rels[rid];
    if (!rel || rel.ext) return null;
    const path = this.resolvePath(rel.target, base);
    if (!this.mediaCache.has(path)) {
      const f = this.zip.file(path);
      const ext = path.split('.').pop().toLowerCase();
      const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }[ext];
      this.mediaCache.set(path, f && type ? f.async('blob').then((b) => URL.createObjectURL(new Blob([b], { type }))) : null);
    }
    return this.mediaCache.get(path);
  }

  imgBox(src, w, h) {
    if (!src) {
      const ph = el('span', 'noimg', 'рисунок');
      ph.style.width = w + 'pt';
      ph.style.height = h + 'pt';
      return ph;
    }
    const img = document.createElement('img');
    img.draggable = false;
    img.style.width = w + 'pt';
    img.style.height = h + 'pt';
    src.then((u) => (img.src = u));
    return img;
  }

  drawingEl(d) {
    const ext = d.getElementsByTagNameNS(NS.wp, 'extent')[0];
    const w = ext ? +ext.getAttribute('cx') / 12700 : 120;
    const h = ext ? +ext.getAttribute('cy') / 12700 : 80;
    const blip = d.getElementsByTagNameNS(NS.a, 'blip')[0];
    if (blip) return this.imgBox(this.media(blip.getAttributeNS(NS.r, 'embed')), w, h);
    const txt = [...d.getElementsByTagNameNS(NS.w, 'p')].map((p) => p.textContent).join('\n');
    if (txt.trim()) { const b = el('span', 'txbx'); b.textContent = txt; return b; }
    return this.imgBox(null, w, h);
  }

  vmlEl(c) {
    const im = c.getElementsByTagNameNS(NS.v, 'imagedata')[0];
    const shape = c.getElementsByTagNameNS(NS.v, 'shape')[0];
    const st = (shape && shape.getAttribute('style')) || '';
    const num = (k) => { const m = st.match(new RegExp(k + ':\\s*([\\d.]+)(pt|in|px)?')); return m ? +m[1] * (m[2] === 'in' ? 72 : m[2] === 'px' ? 0.75 : 1) : 0; };
    const w = num('width') || 120, h = num('height') || 80;
    return this.imgBox(im ? this.media(im.getAttributeNS(NS.r, 'id')) : null, w, h);
  }

  renderSdt(sdt) {
    const box = el('div', 'atom-block');
    box.contentEditable = 'false';
    box.dataset.x = this.atoms.push([sdt]) - 1;
    for (const c of (wkid(sdt, 'sdtContent') || { children: [] }).children) {
      if (c.localName === 'p') box.append(this.renderP(c));
      else if (c.localName === 'tbl') box.append(this.renderTbl(c));
    }
    return box;
  }

  cellBorders(td, tb) {
    const map = { top: 'Top', bottom: 'Bottom', left: 'Left', start: 'Left', right: 'Right', end: 'Right' };
    for (const c of tb.children) {
      const side = map[c.localName];
      if (!side) continue;
      const v = wa(c, 'val');
      if (v === 'nil' || v === 'none') td.style['border' + side] = 'none';
      else {
        const col = wa(c, 'color');
        td.style['border' + side] = `${Math.max(0.5, (+wa(c, 'sz') || 4) / 8)}pt solid ${col && col !== 'auto' ? '#' + col : '#000'}`;
      }
    }
  }

  renderTbl(tbl) {
    const t = document.createElement('table');
    t.dataset.k = this.store.push(tbl) - 1;
    const tblPr = wkid(tbl, 'tblPr');
    const own = wkid(tblPr, 'tblBorders');
    if (own ? this.bordersOn(own) : this.style(wa(wkid(tblPr, 'tblStyle'), 'val')).borders) t.classList.add('bordered');
    const jc = wa(wkid(tblPr, 'jc'), 'val');
    if (jc === 'center') { t.style.marginLeft = 'auto'; t.style.marginRight = 'auto'; t.dataset.tjc = 'center'; }
    else if (jc === 'right' || jc === 'end') { t.style.marginLeft = 'auto'; t.dataset.tjc = 'right'; }
    const ind = wkid(tblPr, 'tblInd');
    if (ind && wa(ind, 'type') !== 'pct' && +wa(ind, 'w')) { t.style.marginLeft = +wa(ind, 'w') * TWIP + 'pt'; t.dataset.tind = +wa(ind, 'w') * TWIP; }
    const grid = wkids(wkid(tbl, 'tblGrid'), 'gridCol').map((g) => +(wa(g, 'w') || 0) * TWIP);
    if (grid.length && grid.every((x) => x > 0)) {
      const cg = document.createElement('colgroup');
      for (const w of grid) { const c = document.createElement('col'); c.style.width = w + 'pt'; cg.append(c); }
      t.append(cg);
      t.style.width = grid.reduce((a, b) => a + b, 0) + 'pt';
    }
    const tbody = document.createElement('tbody');
    t.append(tbody);
    const open = [];
    wkids(tbl, 'tr').forEach((tr, ri) => {
      const row = document.createElement('tr');
      row.dataset.tr = ri;
      const trPr = wkid(tr, 'trPr');
      const trh = wkid(trPr, 'trHeight');
      if (trh) row.style.height = +wa(trh, 'val') * TWIP + 'pt';
      let col = +(wa(wkid(trPr, 'gridBefore'), 'val') || 0);
      wkids(tr, 'tc').forEach((tc, ci) => {
        const tcPr = wkid(tc, 'tcPr');
        const span = +(wa(wkid(tcPr, 'gridSpan'), 'val') || 1);
        const vm = wkid(tcPr, 'vMerge');
        const vmv = vm ? wa(vm, 'val') || 'continue' : null;
        if (vmv === 'continue') {
          if (open[col]) open[col].rowSpan += 1;
          col += span;
          return;
        }
        const td = document.createElement('td');
        td.dataset.tc = ci;
        if (span > 1) td.colSpan = span;
        for (let i = 0; i < span; i++) open[col + i] = null;
        if (vmv === 'restart') open[col] = td;
        const va = wa(wkid(tcPr, 'vAlign'), 'val');
        if (va) td.style.verticalAlign = va === 'center' ? 'middle' : va;
        const shd = wa(wkid(tcPr, 'shd'), 'fill');
        if (shd && shd !== 'auto') td.style.background = '#' + shd;
        const tb = wkid(tcPr, 'tcBorders');
        if (tb) this.cellBorders(td, tb);
        for (const b of tc.children) {
          if (b.localName === 'p') td.append(this.renderP(b));
          else if (b.localName === 'tbl') td.append(this.renderTbl(b));
          else if (b.localName === 'sdt') td.append(this.renderSdt(b));
        }
        if (!td.firstChild) td.append(this.emptyP());
        row.append(td);
        col += span;
      });
      tbody.append(row);
    });
    return t;
  }

  parseSect(sp) {
    const sz = wkid(sp, 'pgSz'), mar = wkid(sp, 'pgMar');
    const n = (x, a, d) => Math.abs(+(wa(x, a) ?? d)) * TWIP;
    const refs = (kind) => {
      const o = {};
      for (const r of wkids(sp, kind + 'Reference')) o[wa(r, 'type') || 'default'] = r.getAttributeNS(NS.r, 'id');
      return o;
    };
    const pn = wkid(sp, 'pgNumType');
    return {
      pg: {
        w: n(sz, 'w', 11906), h: n(sz, 'h', 16838),
        mt: n(mar, 'top', 1440), mb: n(mar, 'bottom', 1440), ml: n(mar, 'left', 1440), mr: n(mar, 'right', 1440),
        hd: n(mar, 'header', 708), ft: n(mar, 'footer', 708),
      },
      type: wa(wkid(sp, 'type'), 'val') || 'nextPage',
      header: refs('header'), footer: refs('footer'),
      titlePg: onoff(wkid(sp, 'titlePg')),
      start: pn && wa(pn, 'start') != null ? +wa(pn, 'start') : null,
    };
  }

  // Колонтитулы: у раздела без своего — колонтитул предыдущего раздела (как в Word).
  async loadHeaderFooter() {
    this.hfBoxes = {};
    let prev = { header: {}, footer: {} };
    for (const s of this.sections) {
      s.header = { ...prev.header, ...s.header };
      s.footer = { ...prev.footer, ...s.footer };
      prev = s;
    }
    const ids = new Set();
    for (const s of this.sections) for (const k of ['header', 'footer']) for (const id of Object.values(s[k])) ids.add(id);
    for (const rid of ids) {
      const rel = this.rels[rid];
      if (!rel) continue;
      const path = this.resolvePath(rel.target);
      const x = await this.readXml(path);
      if (!x) continue;
      const dir = path.slice(0, path.lastIndexOf('/') + 1);
      this.ctxRels = this.readRels(await this.readXml(dir + '_rels/' + path.slice(dir.length) + '.rels'));
      this.ctxBase = dir;
      const box = el('div', 'wd-hf wd-surface');
      for (const b of x.documentElement.children) {
        if (b.localName === 'p') box.append(this.renderP(b));
        else if (b.localName === 'tbl') box.append(this.renderTbl(b));
        else if (b.localName === 'sdt') box.append(this.renderSdt(b));
      }
      this.ctxRels = this.rels;
      this.ctxBase = this.base;
      if (box.textContent.trim() || box.querySelector('img')) this.hfBoxes[rid] = box;
    }
  }

  setupSections() {
    this.maxW = Math.max(...this.sections.map((s) => s.pg.w)) * PXPT;
    this.pg = this.sections[0].pg;
    this.wrap.style.width = this.maxW + 'px';
    this.surface.style.padding = `${this.pg.mt * PXPT}px 0 0 0`;
  }

  // Номер раздела для каждого блока верхнего уровня.
  blockSections(blocks) {
    const out = [];
    let s = 0;
    for (const b of blocks) {
      out.push(Math.min(s, this.sections.length - 1));
      if (b.dataset.sect != null) s = Math.max(s, +b.dataset.sect + 1);
    }
    return out;
  }

  // ---------- Нумерация списков ----------
  renumber() {
    const counters = {};
    const seen = new Set();
    for (const p of this.surface.querySelectorAll('p[data-num]')) {
      const [numId, il] = p.dataset.num.split(':');
      const ilvl = +il;
      const num = this.nums[numId];
      const lv = num && this.abs[num.abs] && this.abs[num.abs][ilvl];
      if (!lv) { p.removeAttribute('data-marker'); continue; }
      const c = counters[num.abs] || (counters[num.abs] = []);
      if (!seen.has(numId)) {
        seen.add(numId);
        for (const [l, s] of Object.entries(num.ov)) c[+l] = s - 1;
      }
      c[ilvl] = (c[ilvl] ?? lv.start - 1) + 1;
      c.length = ilvl + 1;
      const text = lv.fmt === 'bullet' ? mapBullet(lv.text)
        : lv.text.replace(/%(\d)/g, (_m, dd) => {
          const L = this.abs[num.abs][dd - 1];
          return fmtNum(L ? L.fmt : 'decimal', c[dd - 1] ?? (L ? L.start : 1));
        });
      p.dataset.marker = text;
    }
  }

  // ---------- Страницы ----------
  // Блок, который не помещается на лист, переносим на следующий (как Word переносит абзац).
  // У каждого раздела свой размер листа и поля.
  paginate() {
    if (!this.sections || !this.sections.length) return;
    const G = PAGE_GAP, X = PXPT;
    if (this.mo) { this.noteChanges(this.mo.takeRecords()); this.mo.disconnect(); }
    this.surface.querySelectorAll('tr.wd-pgap').forEach((r) => r.remove());
    const blocks = [...this.surface.children];
    const sec = this.blockSections(blocks);
    blocks.forEach((b, i) => {
      const g = this.sections[sec[i]].pg;
      const left = (this.maxW - g.w * X) / 2 + g.ml * X;
      const cw = (g.w - g.ml - g.mr) * X;
      b.style.marginTop = '';
      b.classList.remove('wd-split');
      if (b.tagName === 'TABLE') {
        const tw = b.offsetWidth;
        const off = b.dataset.tjc === 'center' ? (cw - tw) / 2 : b.dataset.tjc === 'right' ? cw - tw : (+b.dataset.tind || 0) * X;
        b.style.marginLeft = left + off + 'px';
        b.style.marginRight = '0';
      } else {
        b.style.marginLeft = left + 'px';
        b.style.width = cw + 'px';
      }
    });
    const tops = blocks.map((b) => b.offsetTop);
    const hs = blocks.map((b) => b.offsetHeight);
    const pages = [];
    const newPage = (s, prevPage) => {
      const top = prevPage ? prevPage.top + this.sections[prevPage.sec].pg.h * X + G : 0;
      const pg = { top, sec: s };
      pages.push(pg);
      return pg;
    };
    let cur = newPage(sec[0] || 0, null);
    let shift = 0, force = false;
    const push = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i], s = sec[i], h = hs[i];
      const t = tops[i] + shift;
      let d = 0, hh = h;
      const empty = b.tagName === 'P' && !b.textContent.trim() && !b.querySelector('img,.atom:not(.hid)');
      if (s !== cur.sec && this.sections[s].type !== 'continuous') {
        cur = newPage(s, cur);
        d = cur.top + this.sections[s].pg.mt * X - t;
      } else {
        const g = this.sections[cur.sec].pg;
        const cTop = cur.top + g.mt * X, cBot = cur.top + (g.h - g.mb) * X;
        if ((force || b.dataset.pbb) && t > cTop + 1) {
          cur = newPage(cur.sec, cur);
          d = cur.top + g.mt * X - t;
        } else if (t < cTop - 0.5) d = cTop - t;
        else if (t + h > cBot + 0.5 && t > cTop + 1 && !(empty && b.dataset.sect != null)) {
          // Длинный абзац делим между листами, как Word; короткий переносим целиком.
          if (b.tagName === 'TABLE') {
            // Таблицу делим между строками: перед строкой, которая не влезла, — невидимая прокладка.
            const r = this.splitTable(b, t, cur, newPage);
            shift += r.h - h;
            hh = r.h;
            cur = r.cur;
            push[i] = 0;
            force = false;
            continue;
          }
          const split = this.trySplit(b, cBot - t, (g.mb + g.mt) * X + G);
          if (split != null) {
            shift += split - h;
            hh = split;
            cur = newPage(cur.sec, cur);
          } else if (h <= (g.h - g.mt - g.mb) * X) {
            cur = newPage(cur.sec, cur);
            d = cur.top + g.mt * X - t;
          }
        }
      }
      push[i] = d;
      shift += d;
      // Слишком высокий блок (большая таблица) идёт через несколько листов.
      if (!(empty && b.dataset.sect != null)) {
        const bottom = t + d + hh;
        for (let guard = 0; guard < 200; guard++) {
          const g = this.sections[cur.sec].pg;
          if (bottom <= cur.top + (g.h - g.mb) * X + 0.5) break;
          cur = newPage(cur.sec, cur);
        }
      }
      force = b.tagName === 'P' && !!b.querySelector('.atom.pbreak');
    }
    blocks.forEach((b, i) => { if (push[i] > 0.5) b.style.marginTop = push[i] + 'px'; });
    if (this.mo) this.mo.observe(this.surface, { childList: true, characterData: true, subtree: true });
    this.pageList = pages;
    this.drawSheets(pages);
    if (this.root.classList.contains('side-open')) {
      clearTimeout(this.thumbTimer);
      this.thumbTimer = setTimeout(() => this.buildThumbs(), 350);
    }
  }

  // Невидимый «вкладыш» внутри абзаца отодвигает строки ниже границы листа на следующий лист.
  // Сам текст при этом не меняется — это только оформление.
  trySplit(b, room, gap) {
    if (b.tagName !== 'P' || b.dataset.num || b.dataset.marker) return null;
    const cs = getComputedStyle(b);
    const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.15;
    const y = room - parseFloat(cs.paddingTop);
    if (y < lh * 0.95 || b.offsetHeight - parseFloat(cs.paddingTop) - y < lh * 0.5) return null;
    b.style.setProperty('--sy', y + 'px');
    b.style.setProperty('--sg', gap + 'px');
    b.classList.add('wd-split');
    return b.offsetHeight;
  }

  splitTable(t, top, cur, newPage) {
    const X = PXPT, G = PAGE_GAP;
    const rows = [...t.rows];
    const cols = rows.reduce((m, r) => Math.max(m, [...r.cells].reduce((a, c) => a + c.colSpan, 0)), 1);
    let extra = 0;
    for (const row of rows) {
      const g = this.sections[cur.sec].pg;
      const cTop = cur.top + g.mt * X, cBot = cur.top + (g.h - g.mb) * X;
      const rt = top + row.offsetTop + extra, rh = row.offsetHeight;
      if (rt + rh <= cBot + 0.5 || rt <= cTop + 1 || rh > (g.h - g.mt - g.mb) * X) {
        if (rt + rh > cBot + 0.5 && rh > (g.h - g.mt - g.mb) * X) cur = newPage(cur.sec, cur);
        continue;
      }
      if ([...row.cells].some((c) => c.rowSpan > 1)) continue;
      cur = newPage(cur.sec, cur);
      const gap = cur.top + this.sections[cur.sec].pg.mt * X - rt;
      const sp = document.createElement('tr');
      sp.className = 'wd-pgap';
      sp.contentEditable = 'false';
      const td = document.createElement('td');
      td.colSpan = cols;
      td.style.height = gap + 'px';
      sp.append(td);
      row.before(sp);
      extra += gap;
    }
    void G;
    return { h: t.offsetHeight, cur };
  }

  drawSheets(pages) {
    const X = PXPT;
    const key = pages.map((p) => `${Math.round(p.top)}:${p.sec}`).join('|');
    if (key !== this.sheetsKey) {
      this.sheets.textContent = '';
      let num = 0, prevSec = -1;
      pages.forEach((pg) => {
        const S = this.sections[pg.sec], g = S.pg;
        const first = pg.sec !== prevSec;
        if (first && S.start != null) num = S.start - 1;
        prevSec = pg.sec;
        num++;
        const s = el('div', 'wd-sheet');
        Object.assign(s.style, { top: pg.top + 'px', height: g.h * X + 'px', width: g.w * X + 'px', left: (this.maxW - g.w * X) / 2 + 'px', right: 'auto' });
        for (const kind of ['header', 'footer']) {
          const rid = first && S.titlePg ? S[kind].first : S[kind].default;
          const src = rid && this.hfBoxes[rid];
          if (!src) continue;
          const box = src.cloneNode(true);
          box.style.left = g.ml * X + 'px';
          box.style.right = g.mr * X + 'px';
          if (kind === 'header') box.style.top = g.hd * X + 'px';
          else box.style.bottom = g.ft * X + 'px';
          box.querySelectorAll('.fld-page').forEach((f) => (f.textContent = num));
          box.querySelectorAll('.fld-pages').forEach((f) => (f.textContent = pages.length));
          s.append(box);
        }
        this.sheets.append(s);
      });
      this.sheetsKey = key;
    }
    this.pages = pages.length;
    const last = pages[pages.length - 1];
    this.wrapH = last.top + this.sections[last.sec].pg.h * X;
    this.wrap.style.height = this.wrapH + 'px';
    this.applyZoom();
  }

  applyZoom() {
    const z = this.zoom;
    this.wrap.style.transform = `scale(${z})`;
    this.scaleBox.style.width = (this.maxW || this.pg.w * PXPT) * z + 'px';
    this.scaleBox.style.height = (this.wrapH || this.pg.h * PXPT) * z + 'px';
  }

  fitWidth() {
    const avail = this.viewer.clientWidth - 84;
    this.setZoom(clamp(avail / (this.maxW || this.pg.w * PXPT), 0.3, 1.25), true);
    this.fitMode = true;
  }

  setZoom(z, quiet) {
    z = clamp(z, 0.3, 3);
    this.fitMode = false;
    const v = this.viewer;
    const fy = v.scrollTop / Math.max(1, v.scrollHeight);
    this.zoom = z;
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

  // ---------- Изменения ----------
  observe() {
    if (this.mo) this.mo.disconnect();
    this.mo = new MutationObserver((recs) => this.noteChanges(recs));
    this.mo.observe(this.surface, { childList: true, characterData: true, subtree: true });
  }

  noteChanges(recs) {
    let any = false;
    for (const r of recs) {
      const n = r.target.nodeType === 3 ? r.target.parentElement : r.target;
      if (!n || !this.surface.contains(n)) continue;
      this.markUp(n);
      for (const a of r.addedNodes) if (a.nodeType === 1) { this.changed.add(a); a.querySelectorAll && a.querySelectorAll('p,table').forEach((x) => this.changed.add(x)); }
      any = true;
    }
    if (any) this.touch();
  }

  // Помечаем изменённым абзац и все таблицы вокруг него.
  markUp(n) {
    for (let e = n; e && e !== this.surface; e = e.parentElement) {
      if (e.tagName === 'P' || e.tagName === 'TABLE') this.changed.add(e);
    }
  }

  touch() {
    if (!this.dirty) { this.dirty = true; this.onChange && this.onChange(); }
    clearTimeout(this.pagTimer);
    this.pagTimer = setTimeout(() => { this.renumber(); this.paginate(); }, 180);
  }

  flushChanges() {
    if (this.mo) this.noteChanges(this.mo.takeRecords());
  }

  // ---------- Сохранение в .docx ----------
  w(name, attrs) {
    const e = this.xml.createElementNS(NS.w, 'w:' + name);
    if (attrs) for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(NS.w, 'w:' + k, v);
    return e;
  }

  setChild(parent, name, node, order) {
    for (const c of wkids(parent, name)) parent.removeChild(c);
    if (!node) return;
    const idx = order.indexOf(name);
    const before = [...parent.children].find((c) => order.indexOf(c.localName) > idx) || null;
    parent.insertBefore(node, before);
  }

  serialize() {
    this.flushChanges();
    const doc = this.xml;
    const body = wkid(doc.documentElement, 'body');
    const kCount = new Map();
    for (const e of this.surface.querySelectorAll('[data-k]')) kCount.set(e.dataset.k, (kCount.get(e.dataset.k) || 0) + 1);
    this.ser = { kCount, kSeen: new Map(), bm: new Set(), cs: new Map(), imgs: [], samples: new Map() };
    const out = [];
    for (const i of this.lead) out.push(...this.atoms[i].map((n) => n.cloneNode(true)));
    for (const b of this.surface.children) {
      out.push(...this.emitBlock(b));
      for (const i of this.tails.get(b) || []) out.push(...this.atoms[i].map((n) => n.cloneNode(true)));
    }
    if (!out.length || out[out.length - 1].localName !== 'p') out.push(this.w('p'));
    while (body.firstChild) body.removeChild(body.firstChild);
    for (const n of out) body.append(n);
    if (this.finalSect) body.append(this.finalSect.cloneNode(true));
    if (this.mo) this.mo.takeRecords(); // временные образцы — не правки пользователя
    return xmlString(doc);
  }

  emitBlock(b) {
    if (b.tagName === 'TABLE') return [this.emitTable(b)];
    if (b.classList.contains('atom-block')) return this.atoms[+b.dataset.x].map((n) => n.cloneNode(true));
    if (b.classList.contains('atom') && b.dataset.x != null) return this.atoms[+b.dataset.x].map((n) => n.cloneNode(true));
    return [this.emitPara(b)];
  }

  emitPara(p) {
    const k = p.dataset.k;
    const orig = k != null ? this.store[+k] : null;
    const seen = (this.ser.kSeen.get(k) || 0) + 1;
    if (k != null) this.ser.kSeen.set(k, seen);
    if (orig && orig.localName === 'p' && !this.changed.has(p) && this.ser.kCount.get(k) === 1) return orig.cloneNode(true);
    const P = this.w('p');
    const opPr = orig && wkid(orig, 'pPr');
    const pPr = opPr ? opPr.cloneNode(true) : this.w('pPr');
    // Разбили абзац с разрывом раздела — разрыв остаётся у последней части.
    if (k != null && seen < this.ser.kCount.get(k)) this.setChild(pPr, 'sectPr', null, PPR_ORDER);
    const sid = p.dataset.style;
    this.setChild(pPr, 'pStyle', sid ? this.w('pStyle', { val: sid }) : null, PPR_ORDER);
    const jc = p.dataset.jc;
    this.setChild(pPr, 'jc', jc ? this.w('jc', { val: jc === 'justify' ? 'both' : jc }) : null, PPR_ORDER);
    if (p.dataset.pbb) this.setChild(pPr, 'pageBreakBefore', this.w('pageBreakBefore'), PPR_ORDER);
    if (p.dataset.num) {
      const [nid, il] = p.dataset.num.split(':');
      const np = this.w('numPr');
      np.append(this.w('ilvl', { val: il }), this.w('numId', { val: nid }));
      this.setChild(pPr, 'numPr', np, PPR_ORDER);
    } else {
      const st = this.style(sid || this.defPStyle);
      if (st.pPr.numId && st.pPr.numId !== '0') {
        const np = this.w('numPr');
        np.append(this.w('numId', { val: '0' }));
        this.setChild(pPr, 'numPr', np, PPR_ORDER);
      } else this.setChild(pPr, 'numPr', null, PPR_ORDER);
    }
    if (pPr.children.length) P.append(pPr);
    let ends = [];
    if (orig && !this.ser.bm.has(k)) {
      this.ser.bm.add(k);
      for (const b of wkids(orig, 'bookmarkStart')) P.append(b.cloneNode(true));
      ends = wkids(orig, 'bookmarkEnd');
    }
    this.emitInline(p, P);
    for (const b of ends) P.append(b.cloneNode(true));
    return P;
  }

  emitInline(p, P) {
    const pieces = [];
    const walk = (node) => {
      for (const c of node.childNodes) {
        if (c.nodeType === 3) {
          const t = c.data.replace(/ /g, ' ').replace(/[\r\n​]/g, '');
          if (t) pieces.push({ t: 'text', text: t, el: c.parentElement });
        } else if (c.nodeType === 1) {
          if (c.classList.contains('atom')) {
            if (c.dataset.x != null) pieces.push({ t: 'atom', x: +c.dataset.x });
            else if (c.dataset.img) pieces.push({ t: 'img', id: c.dataset.img });
            else if (c.dataset.pb) pieces.push({ t: 'pb' });
          } else if (c.tagName === 'BR') pieces.push({ t: 'br', el: c.parentElement });
          else if (c.tagName === 'P' || c.tagName === 'DIV') { if (pieces.length) pieces.push({ t: 'br', el: c }); walk(c); }
          else if (c.tagName !== 'IMG' && c.tagName !== 'TABLE') walk(c);
        }
      }
    };
    walk(p);
    if (pieces.length && pieces[pieces.length - 1].t === 'br') pieces.pop(); // последний <br> — служебный
    let run = null, key = null;
    for (const pc of pieces) {
      if (pc.t === 'atom') { run = null; for (const n of this.atoms[pc.x]) P.append(n.cloneNode(true)); continue; }
      if (pc.t === 'img') { run = null; P.append(this.newImageRun(pc.id)); continue; }
      if (pc.t === 'pb') { run = null; const r = this.w('r'); r.append(this.w('br', { type: 'page' })); P.append(r); continue; }
      const rPr = this.rPrFor(pc.el, p);
      const k = rPr ? new XMLSerializer().serializeToString(rPr) : '';
      if (!run || k !== key) {
        run = this.w('r');
        if (rPr) run.append(rPr);
        P.append(run);
        key = k;
      }
      if (pc.t === 'br') { run.append(this.w('br')); continue; }
      pc.text.split('\t').forEach((s, i) => {
        if (i) run.append(this.w('tab'));
        if (s) {
          const t = this.w('t');
          t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
          t.textContent = s;
          run.append(t);
        }
      });
    }
  }

  // Свойства фрагмента текста: исходные свойства прогона + то, что видно иначе на экране.
  rPrFor(elx, p) {
    const cache = this.ser.cs;
    const ck = elx;
    if (cache.has(ck)) { const v = cache.get(ck); return v ? v.cloneNode(true) : null; }
    // Браузер при форматировании копирует исходный <span data-r> и добавляет ему стиль,
    // поэтому сравниваем не с ним, а с «образцом» — как этот прогон выглядел изначально.
    const baseEl = elx.closest('[data-r]');
    const r = baseEl && p.contains(baseEl) ? baseEl.dataset.r : null;
    const rPr = r != null ? this.rprStore[+r].cloneNode(true) : this.w('rPr');
    const a = this.looks(elx, p), b = this.sample(p, r);
    const RO = RPR_ORDER;
    if (a.bold !== b.bold) this.setChild(rPr, 'b', a.bold ? this.w('b') : this.w('b', { val: '0' }), RO);
    if (a.ital !== b.ital) this.setChild(rPr, 'i', a.ital ? this.w('i') : this.w('i', { val: '0' }), RO);
    if (a.u !== b.u) this.setChild(rPr, 'u', this.w('u', { val: a.u ? 'single' : 'none' }), RO);
    if (a.strike !== b.strike) this.setChild(rPr, 'strike', a.strike ? this.w('strike') : this.w('strike', { val: '0' }), RO);
    if (a.color !== b.color) this.setChild(rPr, 'color', this.w('color', { val: rgbHex(a.color) }), RO);
    if (a.hp !== b.hp && a.va === b.va) {
      this.setChild(rPr, 'sz', this.w('sz', { val: String(a.hp) }), RO);
      this.setChild(rPr, 'szCs', this.w('szCs', { val: String(a.hp) }), RO);
    }
    if (a.fam !== b.fam) this.setChild(rPr, 'rFonts', this.w('rFonts', { ascii: a.fam, hAnsi: a.fam, cs: a.fam, eastAsia: a.fam }), RO);
    const bg = a.bg !== b.bg ? a.bg : null;
    if (a.bg !== b.bg && !a.bg) { this.setChild(rPr, 'highlight', null, RO); this.setChild(rPr, 'shd', null, RO); }
    if (bg) {
      const hex = '#' + rgbHex(bg).toLowerCase();
      const name = Object.entries(WD_HL_MAP).find(([, v]) => v === hex);
      this.setChild(rPr, 'shd', null, RO);
      if (name) this.setChild(rPr, 'highlight', this.w('highlight', { val: name[0] }), RO);
      else this.setChild(rPr, 'shd', this.w('shd', { val: 'clear', color: 'auto', fill: hex.slice(1).toUpperCase() }), RO);
    }
    const res = rPr.children.length ? rPr : null;
    cache.set(ck, res);
    return res ? res.cloneNode(true) : null;
  }

  // Как выглядит текст внутри элемента (жирность, цвет, размер, подчёркивание, фон).
  looks(elx, p) {
    const s = getComputedStyle(elx);
    const deco = (what) => {
      for (let e = elx; e; e = e.parentElement) {
        if (getComputedStyle(e).textDecorationLine.includes(what)) return true;
        if (e === p) break;
      }
      return false;
    };
    let bg = null;
    for (let e = elx; e && e !== p; e = e.parentElement) {
      const c = getComputedStyle(e).backgroundColor;
      if (c && !/rgba\(0, 0, 0, 0\)|transparent/.test(c)) { bg = c; break; }
    }
    return {
      bold: +s.fontWeight >= 600, ital: s.fontStyle === 'italic', u: deco('underline'), strike: deco('line-through'),
      color: s.color, hp: Math.round(parseFloat(s.fontSize) * 0.75 * 2), va: s.verticalAlign, fam: firstFam(s.fontFamily), bg,
    };
  }

  // Образец исходного прогона: временный span с исходным оформлением внутри того же абзаца.
  sample(p, r) {
    const key = r == null ? '' : r;
    let m = this.ser.samples.get(p);
    if (!m) this.ser.samples.set(p, (m = new Map()));
    if (m.has(key)) return m.get(key);
    const s = document.createElement('span');
    if (r != null) s.style.cssText = this.runCss(this.runProps(this.rprStore[+r]));
    s.textContent = 'x';
    p.append(s);
    const v = this.looks(s, p);
    s.remove();
    m.set(key, v);
    return v;
  }

  emitTable(t) {
    const k = t.dataset.k;
    const orig = k != null ? this.store[+k] : null;
    if (orig && !this.changed.has(t) && this.ser.kCount.get(k) === 1) return orig.cloneNode(true);
    const tbl = orig ? orig.cloneNode(true) : this.newTblXml(t);
    const origTrs = orig ? wkids(orig, 'tr') : [];
    for (const tr of wkids(tbl, 'tr')) tbl.removeChild(tr);
    for (const row of t.rows) {
      if (row.classList.contains('wd-pgap')) continue;
      const src = row.dataset.tr != null ? origTrs[+row.dataset.tr] : null;
      const xtr = src ? src.cloneNode(true) : this.newTrXml(row, t);
      const tcs = wkids(xtr, 'tc');
      [...row.cells].forEach((td, i) => {
        const tc = src ? tcs[+td.dataset.tc] : tcs[i];
        if (!tc) return;
        for (const c of [...tc.children]) if (['p', 'tbl', 'sdt'].includes(c.localName)) tc.removeChild(c);
        const blocks = [];
        for (const b of td.children) blocks.push(...this.emitBlock(b));
        if (!blocks.length || blocks[blocks.length - 1].localName !== 'p') blocks.push(this.w('p'));
        for (const b of blocks) tc.append(b);
      });
      tbl.append(xtr);
    }
    return tbl;
  }

  contentWidthTw() { return Math.round((this.pg.w - this.pg.ml - this.pg.mr) * 20); }

  newTblXml(t) {
    const cols = t.rows[0] ? t.rows[0].cells.length : 1;
    const tbl = this.w('tbl');
    const pr = this.w('tblPr');
    pr.append(this.w('tblW', { w: '0', type: 'auto' }));
    const bd = this.w('tblBorders');
    for (const s of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) bd.append(this.w(s, { val: 'single', sz: '4', space: '0', color: 'auto' }));
    pr.append(bd, this.w('tblLook', { val: '04A0', firstRow: '1', lastRow: '0', firstColumn: '1', lastColumn: '0', noHBand: '0', noVBand: '1' }));
    const grid = this.w('tblGrid');
    const cw = Math.floor(this.contentWidthTw() / cols);
    for (let i = 0; i < cols; i++) grid.append(this.w('gridCol', { w: String(cw) }));
    tbl.append(pr, grid);
    return tbl;
  }

  newTrXml(row, t) {
    const tr = this.w('tr');
    const cols = row.cells.length;
    const cw = Math.floor(this.contentWidthTw() / cols);
    for (let i = 0; i < cols; i++) {
      const tc = this.w('tc');
      const pr = this.w('tcPr');
      pr.append(this.w('tcW', { w: String(cw), type: 'dxa' }));
      tc.append(pr);
      tr.append(tc);
    }
    void t;
    return tr;
  }

  newImageRun(id) {
    const im = this.newImgs.get(id);
    const n = this.ser.imgs.push(id);
    const rid = 'rIdFolio' + id;
    const cx = Math.round(im.wPt * 12700), cy = Math.round(im.hPt * 12700);
    const xml = `<w:r xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${4000 + n}" name="Рисунок ${n}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="folio${id}.${im.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    return this.xml.importNode(new DOMParser().parseFromString(xml, 'application/xml').documentElement, true);
  }

  addRel(id, type, target) {
    const root = this.relsXml.documentElement;
    if ([...root.children].some((r) => r.getAttribute('Id') === id)) return;
    const r = this.relsXml.createElementNS(NS.rel, 'Relationship');
    r.setAttribute('Id', id);
    r.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + type);
    r.setAttribute('Target', target);
    root.append(r);
    this.rels[id] = { type: r.getAttribute('Type'), target, ext: false };
  }

  ensureDefaultType(ext, type) {
    const root = this.ct.documentElement;
    if ([...root.children].some((c) => c.localName === 'Default' && (c.getAttribute('Extension') || '').toLowerCase() === ext)) return;
    const d = this.ct.createElementNS(NS.ct, 'Default');
    d.setAttribute('Extension', ext);
    d.setAttribute('ContentType', type);
    root.insertBefore(d, root.firstChild);
  }

  ensureOverride(part, type) {
    const root = this.ct.documentElement;
    if ([...root.children].some((c) => c.getAttribute('PartName') === '/' + part)) return;
    const o = this.ct.createElementNS(NS.ct, 'Override');
    o.setAttribute('PartName', '/' + part);
    o.setAttribute('ContentType', type);
    root.append(o);
  }

  async buildDocx() {
    const docXml = this.serialize();
    const zip = this.zip;
    zip.file(this.mainPath, docXml);
    for (const id of new Set(this.ser.imgs)) {
      const im = this.newImgs.get(id);
      const name = `media/folio${id}.${im.ext}`;
      zip.file(this.base + name, dataUrlBytes(im.url));
      this.addRel('rIdFolio' + id, 'image', name);
      this.ensureDefaultType(im.ext, im.ext === 'png' ? 'image/png' : 'image/jpeg');
    }
    const ser = xmlString;
    zip.file(this.relsPath, ser(this.relsXml));
    zip.file('[Content_Types].xml', ser(this.ct));
    if (this.numTouched) zip.file(this.numPath, ser(this.numXml));
    if (this.stylesTouched) zip.file(this.stylesPath, ser(this.stylesXml));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // ---------- Печать в PDF ----------
  async buildPrintHtml() {
    const clone = this.surface.cloneNode(true);
    clone.removeAttribute('contenteditable');
    clone.style.padding = '0';
    const sec = this.blockSections([...this.surface.children]);
    [...clone.children].forEach((b, i) => {
      b.style.marginTop = '';
      b.style.width = '';
      b.style.marginLeft = b.dataset.tjc ? 'auto' : b.dataset.tind ? b.dataset.tind + 'pt' : '';
      b.style.marginRight = b.dataset.tjc === 'center' ? 'auto' : '';
      b.style.page = 's' + sec[i];
      b.classList.remove('wd-split');
    });
    clone.querySelectorAll('.atom.pbreak').forEach((n) => (n.style.breakAfter = 'page'));

    clone.querySelectorAll('[data-pbb]').forEach((n) => (n.style.breakBefore = 'page'));
    clone.querySelectorAll('.atom.hid, tr.wd-pgap').forEach((n) => n.remove());
    for (const img of clone.querySelectorAll('img')) {
      if (img.src.startsWith('blob:')) img.src = await readAsDataUrl(await (await fetch(img.src)).blob());
    }
    const g = this.pg;
    const pages = this.sections.map((s, i) => `@page s${i} { size: ${s.pg.w}pt ${s.pg.h}pt; margin: ${s.pg.mt}pt ${s.pg.mr}pt ${s.pg.mb}pt ${s.pg.ml}pt; }`).join('\n');
    const css = `${WD_CSS}
@page { size: ${g.w}pt ${g.h}pt; margin: ${g.mt}pt ${g.mr}pt ${g.mb}pt ${g.ml}pt; }
${pages}
html, body { margin: 0; background: #fff; }
.wd-surface p, .wd-surface tr { break-inside: avoid-page; }
.wd-surface p { orphans: 2; widows: 2; }`;
    return `<!doctype html><html><head><meta charset="utf-8"><style>${css}</style></head><body><div class="wd-surface">${clone.innerHTML}</div></body></html>`;
  }

  async exportPdf(target) {
    const html = await this.buildPrintHtml();
    const bytes = await window.folio.htmlToPdf(html);
    await window.folio.writeFile(target, bytes);
  }

  async save(asNew) {
    let target = this.path && /\.docx$/i.test(this.path) && !asNew ? this.path : null;
    if (!target) {
      const base = (this.path || T('Документ')).replace(/\.[^.\\/]+$/, '');
      target = await window.folio.saveDialog(base + '.docx', [
        { name: T('Документ Word'), extensions: ['docx'] },
        { name: 'PDF', extensions: ['pdf'] },
      ]);
      if (!target) return false;
    }
    if (/\.pdf$/i.test(target)) {
      await this.exportPdf(target);
      toast('Сохранено в PDF');
      return false;
    }
    const bytes = await this.buildDocx();
    await window.folio.writeFile(target, bytes);
    this.path = target;
    this.dirty = false;
    this.onChange && this.onChange();
    return true;
  }

  // ---------- Ввод ----------
  bindSurface() {
    const s = this.surface;
    document.execCommand('defaultParagraphSeparator', false, 'p');
    s.addEventListener('focus', () => {
      document.execCommand('styleWithCSS', false, true);
      document.execCommand('defaultParagraphSeparator', false, 'p');
    });
    s.addEventListener('keydown', (e) => this.onSurfaceKey(e));
    s.addEventListener('paste', (e) => this.onPaste(e));
    s.addEventListener('drop', (e) => { if (e.dataTransfer && e.dataTransfer.files.length) e.preventDefault(); });
    s.addEventListener('contextmenu', (e) => this.onContext(e));
    // Щелчок по пустому месту листа — курсор в документ.
    this.wrap.addEventListener('mousedown', (e) => {
      if (this.surface.contains(e.target)) return;
      e.preventDefault();
      this.surface.focus();
      const last = this.surface.lastElementChild;
      if (last) { const r = document.createRange(); r.selectNodeContents(last); r.collapse(false); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); }
    });
    this.viewer.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * Math.exp(-e.deltaY * 0.0022));
    }, { passive: false });
    this.viewer.addEventListener('scroll', () => {
      if (!this.pageList || this.pages < 2) return;
      const y = (this.viewer.scrollTop + this.viewer.clientHeight * 0.4) / this.zoom;
      let cur = 1;
      this.pageList.forEach((p, i) => { if (p.top <= y) cur = i + 1; });
      this.flashHud(`${cur} / ${this.pages}`);
      this.markCurThumb(cur - 1);
    }, { passive: true });
    this.onSel = () => this.syncTools();
    document.addEventListener('selectionchange', this.onSel);
  }

  placeCaretStart() {
    const first = this.surface.firstElementChild;
    if (!first) return;
    this.surface.focus();
    const r = document.createRange();
    r.setStart(first, 0);
    r.collapse(true);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }

  curBlock() {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    let n = sel.anchorNode;
    if (n && n.nodeType === 3) n = n.parentElement;
    return n && this.surface.contains(n) ? n.closest('p') : null;
  }

  selBlocks() {
    const sel = getSelection();
    if (!sel.rangeCount || !this.surface.contains(sel.anchorNode)) return [];
    const r = sel.getRangeAt(0);
    const out = [...this.surface.querySelectorAll('p')].filter((p) => r.intersectsNode(p) && !p.closest('[contenteditable="false"]'));
    if (!out.length) { const b = this.curBlock(); if (b) out.push(b); }
    return out;
  }

  onSurfaceKey(e) {
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    if (ctrl && (k === '=' || k === '+')) { e.preventDefault(); this.setZoom(this.zoom * 1.1); return; }
    if (ctrl && k === '-') { e.preventDefault(); this.setZoom(this.zoom / 1.1); return; }
    if (ctrl && k === '0') { e.preventDefault(); this.fitWidth(); this.flashHud(Math.round(this.zoom * 100) + '%'); return; }
    if (ctrl && e.key === 'Enter') { e.preventDefault(); this.insertPageBreak(); return; }
    if (ctrl && (k === 'l' || k === 'д')) { e.preventDefault(); this.setAlign('left'); return; }
    if (ctrl && (k === 'e' || k === 'у')) { e.preventDefault(); this.setAlign('center'); return; }
    if (ctrl && (k === 'r' || k === 'к')) { e.preventDefault(); this.setAlign('right'); return; }
    if (ctrl && (k === 'j' || k === 'о')) { e.preventDefault(); this.setAlign('justify'); return; }
    if (ctrl && (k === 'b' || k === 'и')) { e.preventDefault(); this.exec('bold'); return; }
    if (ctrl && (k === 'i' || k === 'ш')) { e.preventDefault(); this.exec('italic'); return; }
    if (ctrl && (k === 'u' || k === 'г')) { e.preventDefault(); this.exec('underline'); return; }
    const p = this.curBlock();
    if (e.key === 'Tab') {
      e.preventDefault();
      if (p && p.dataset.num) {
        const [nid, il] = p.dataset.num.split(':');
        const nl = clamp(+il + (e.shiftKey ? -1 : 1), 0, 8);
        if (this.lvl(nid, nl)) { p.dataset.num = `${nid}:${nl}`; this.paraChanged([p]); }
        return;
      }
      document.execCommand('insertText', false, '\t');
      return;
    }
    // Enter в пустом пункте списка — выход из списка, как в Word.
    if (e.key === 'Enter' && p && p.dataset.pbb) {
      setTimeout(() => {
        const np = this.curBlock();
        if (np && np !== p && np.dataset.pbb) { delete np.dataset.pbb; this.markUp(np); this.touch(); }
      }, 0);
    }
    if (e.key === 'Enter' && !e.shiftKey && p && p.dataset.num && !p.textContent.trim()) {
      e.preventDefault();
      delete p.dataset.num;
      p.removeAttribute('data-marker');
      this.paraChanged([p]);
    }
  }

  async onPaste(e) {
    e.preventDefault();
    const items = [...(e.clipboardData ? e.clipboardData.items : [])];
    const img = items.find((i) => i.type.startsWith('image/'));
    if (img) {
      const f = img.getAsFile();
      await this.insertImageBytes(new Uint8Array(await f.arrayBuffer()), 'paste.' + (f.type.split('/')[1] || 'png'));
      return;
    }
    const text = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
    const lines = text.replace(/\r/g, '').split('\n');
    lines.forEach((line, i) => {
      if (i) document.execCommand('insertParagraph');
      if (line) document.execCommand('insertText', false, line);
    });
  }

  onContext(e) {
    const td = e.target.closest('td');
    if (!td || !this.surface.contains(td) || td.closest('[contenteditable="false"]')) return;
    e.preventDefault();
    const tr = td.parentElement, table = tr.closest('table');
    showMenu(e.clientX, e.clientY, [
      ['Добавить строку выше', () => this.addRow(tr, false)],
      ['Добавить строку ниже', () => this.addRow(tr, true)],
      ['Удалить строку', () => this.removeRow(tr)],
      null,
      ['Удалить таблицу', () => { const n = table.nextElementSibling || table.previousElementSibling; table.remove(); if (!this.surface.children.length) this.surface.append(this.emptyP()); this.touch(); void n; }],
    ]);
  }

  addRow(tr, after) {
    const row = tr.cloneNode(true);
    for (const td of row.cells) {
      td.rowSpan = 1;
      // Новая ячейка получает оформление первого абзаца исходной, но без текста.
      const first = td.querySelector('p');
      const np = first ? first.cloneNode(false) : this.emptyP();
      np.removeAttribute('data-marker');
      np.textContent = '';
      // Оставляем оформление первого фрагмента, чтобы набранный текст был тем же шрифтом.
      const run = first && first.querySelector('span[data-r]');
      const host = run ? run.cloneNode(false) : np;
      if (run) np.append(host);
      host.append(document.createElement('br'));
      td.textContent = '';
      td.append(np);
    }
    if (after) tr.after(row); else tr.before(row);
    this.markUp(tr);
    this.touch();
  }

  removeRow(tr) {
    const table = tr.closest('table');
    if (table.rows.length <= 1) { table.remove(); if (!this.surface.children.length) this.surface.append(this.emptyP()); }
    else { this.markUp(tr); tr.remove(); }
    this.touch();
  }

  // ---------- Команды форматирования ----------
  saveRange() {
    const sel = getSelection();
    if (sel.rangeCount && this.surface.contains(sel.anchorNode)) this.range = sel.getRangeAt(0).cloneRange();
  }

  restoreRange() {
    // Курсор и так в документе — ничего не трогаем; возвращаем, только если фокус ушёл (меню, окно выбора).
    const cur = getSelection();
    if (document.activeElement === this.surface && cur.rangeCount && this.surface.contains(cur.anchorNode)) return;
    this.surface.focus({ preventScroll: true });
    if (this.range) {
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(this.range);
    }
  }

  exec(cmd, val) {
    this.restoreRange();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd, false, val);
    this.saveRange();
    this.syncTools();
  }

  setFontSize(pt) {
    this.restoreRange();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    for (const n of this.surface.querySelectorAll('font[size="7"], span[style*="xxx-large"]')) {
      if (n.tagName === 'FONT') {
        const s = document.createElement('span');
        s.style.fontSize = pt + 'pt';
        while (n.firstChild) s.append(n.firstChild);
        n.replaceWith(s);
      } else n.style.fontSize = pt + 'pt';
    }
    this.saveRange();
    this.syncTools();
  }

  paraChanged(ps) {
    for (const p of ps) { this.applyPara(p); this.markUp(p); }
    this.touch();
    this.renumber();
    this.syncTools();
  }

  setAlign(jc) {
    this.restoreRange();
    const ps = this.selBlocks();
    for (const p of ps) p.dataset.jc = jc;
    this.paraChanged(ps);
  }

  setStyle(name) {
    this.restoreRange();
    const id = this.ensureStyle(name);
    const ps = this.selBlocks();
    for (const p of ps) {
      if (!id || id === this.defPStyle) delete p.dataset.style;
      else p.dataset.style = id;
    }
    this.paraChanged(ps);
  }

  ensureStyle(name) {
    const have = this.styleIdByName(name);
    if (have) return have;
    if (name === 'normal') return this.defPStyle;
    const tpl = STYLE_TEMPLATES[name];
    if (!tpl || !this.stylesXml) return null;
    const node = new DOMParser().parseFromString(`<w:styles xmlns:w="${NS.w}">${tpl}</w:styles>`, 'application/xml').documentElement.firstElementChild;
    const imp = this.stylesXml.importNode(node, true);
    this.stylesXml.documentElement.append(imp);
    if (this.defPStyle && this.defPStyle !== 'Normal') {
      const b = wkid(imp, 'basedOn');
      if (b) b.setAttributeNS(NS.w, 'w:val', this.defPStyle);
    }
    this.stylesTouched = true;
    this.parseStyles();
    return this.styleIdByName(name);
  }

  // Списки: свои два определения (маркеры и номера), создаются при первом использовании.
  ensureList(kind) {
    if (this.listIds[kind]) return this.listIds[kind];
    if (!this.numXml) {
      this.numXml = new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS.w}"/>`, 'application/xml');
      this.numPath = this.base + 'numbering.xml';
      this.addRel('rIdFolioNum', 'numbering', 'numbering.xml');
      this.ensureOverride(this.numPath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
    }
    const root = this.numXml.documentElement;
    const aid = Math.max(-1, ...wkids(root, 'abstractNum').map((a) => +wa(a, 'abstractNumId'))) + 1;
    const nid = Math.max(0, ...wkids(root, 'num').map((n) => +wa(n, 'numId'))) + 1;
    let lv = '';
    for (let i = 0; i < 9; i++) {
      const fmt = kind === 'bullet' ? 'bullet' : ['decimal', 'lowerLetter', 'lowerRoman'][i % 3];
      const txt = kind === 'bullet' ? ['•', '◦', '▪'][i % 3] : `%${i + 1}.`;
      lv += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    const xml = `<w:numbering xmlns:w="${NS.w}"><w:abstractNum w:abstractNumId="${aid}"><w:multiLevelType w:val="hybridMultilevel"/>${lv}</w:abstractNum><w:num w:numId="${nid}"><w:abstractNumId w:val="${aid}"/></w:num></w:numbering>`;
    const frag = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
    const abs = this.numXml.importNode(frag.children[0], true);
    const num = this.numXml.importNode(frag.children[1], true);
    const firstNum = wkid(root, 'num');
    root.insertBefore(abs, firstNum);
    root.append(num);
    this.numTouched = true;
    this.parseNumbering();
    this.listIds[kind] = String(nid);
    return this.listIds[kind];
  }

  listKind(p) {
    if (!p || !p.dataset.num) return null;
    const [nid, il] = p.dataset.num.split(':');
    const lv = this.lvl(nid, +il);
    return lv ? (lv.fmt === 'bullet' ? 'bullet' : 'number') : null;
  }

  toggleList(kind) {
    this.restoreRange();
    const ps = this.selBlocks();
    if (!ps.length) return;
    const off = ps.every((p) => this.listKind(p) === kind);
    const nid = off ? null : this.ensureList(kind);
    for (const p of ps) {
      if (off) { delete p.dataset.num; p.removeAttribute('data-marker'); }
      else p.dataset.num = `${nid}:${p.dataset.num ? p.dataset.num.split(':')[1] : 0}`;
    }
    this.paraChanged(ps);
  }

  insertPageBreak() {
    this.restoreRange();
    // Делим абзац, а разрыв ставим в конец первой половины: следующая часть начнётся с нового листа.
    document.execCommand('insertParagraph');
    const next = this.curBlock();
    const prev = next && next.previousElementSibling;
    if (prev && prev.tagName === 'P') {
      const pb = el('span', 'atom pbreak');
      pb.contentEditable = 'false';
      pb.dataset.pb = '1';
      if (prev.lastChild && prev.lastChild.nodeName === 'BR' && prev.childNodes.length === 1) prev.lastChild.remove();
      prev.append(pb);
      this.markUp(prev);
    }
    this.touch();
  }

  async pickImage() {
    this.saveRange();
    const paths = await window.folio.openDialog({ filters: [{ name: T('Картинки'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] }] });
    if (!paths.length) return;
    await this.insertImageBytes(await window.folio.readFile(paths[0]), paths[0]);
  }

  async insertImageBytes(bytes, name) {
    const im = await normalizeImage(bytes, name);
    const id = String(Date.now() % 1e9) + Math.floor(Math.random() * 1000);
    const maxW = this.pg.w - this.pg.ml - this.pg.mr;
    const wPt = Math.min(maxW, im.w * 0.75), hPt = (im.h * wPt) / im.w;
    this.newImgs.set(id, { url: im.url, wPt, hPt, ext: im.url.startsWith('data:image/jpeg') ? 'jpeg' : 'png' });
    this.restoreRange();
    document.execCommand('insertHTML', false, `<span class="atom" contenteditable="false" data-img="${id}"><img src="${im.url}" style="width:${wPt}pt;height:${hPt}pt"></span>​`);
    this.touch();
  }

  insertTable(rows, cols) {
    this.restoreRange();
    const b = this.curBlock();
    const top = b ? (b.closest('td') ? b : [...this.surface.children].find((x) => x.contains(b))) : this.surface.lastElementChild;
    const t = document.createElement('table');
    t.className = 'bordered';
    const cw = (this.pg.w - this.pg.ml - this.pg.mr) / cols;
    t.style.width = cw * cols + 'pt';
    const cg = document.createElement('colgroup');
    for (let i = 0; i < cols; i++) { const c = document.createElement('col'); c.style.width = cw + 'pt'; cg.append(c); }
    t.append(cg);
    const tb = document.createElement('tbody');
    for (let r = 0; r < rows; r++) {
      const tr = document.createElement('tr');
      for (let c = 0; c < cols; c++) { const td = document.createElement('td'); td.append(this.emptyP()); tr.append(td); }
      tb.append(tr);
    }
    t.append(tb);
    const after = this.emptyP();
    if (top) top.after(t, after); else this.surface.append(t, after);
    const r = document.createRange();
    r.setStart(t.querySelector('p'), 0);
    r.collapse(true);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    this.touch();
  }

  // ---------- Панель инструментов ----------
  buildDock() {
    const dock = el('div', 'dock wd-dock');
    dock.addEventListener('pointerdown', (e) => { if (!e.target.closest('input')) { e.preventDefault(); this.saveRange(); } });
    const btn = (icon, tip, fn, cls = '') => {
      const b = el('button', 'tool ' + cls, icon);
      b.dataset.tip = tip;
      b.addEventListener('click', (e) => fn(e, b));
      dock.append(b);
      return b;
    };
    const sep = () => dock.append(el('div', 'sep'));
    const ui = (this.ui = {});
    ui.pages = btn(ICONS.pages, 'Страницы', () => this.toggleSide());
    sep();
    ui.style = btn('Обычный', 'Стиль абзаца', (e, b) => this.menuAt(b, WD_STYLES.map(([n, label]) => [label, () => this.setStyle(n)])), 'wd-txt wd-style');
    ui.font = btn('Times New Roman', 'Шрифт', (e, b) => this.menuAt(b, WD_FONTS.map((f) => [f, () => this.exec('fontName', f)])), 'wd-txt wd-font');
    const size = el('div', 'size wd-size');
    size.dataset.tip = 'Размер';
    const minus = el('button', '', '−'), plus = el('button', '', '+');
    ui.size = el('output', '', '11');
    const step = (d) => {
      const cur = parseFloat(ui.size.textContent) || 11;
      const sizes = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];
      const next = d > 0 ? sizes.find((s) => s > cur + 0.01) || cur + 4 : [...sizes].reverse().find((s) => s < cur - 0.01) || Math.max(1, cur - 1);
      this.setFontSize(next);
    };
    minus.addEventListener('click', () => step(-1));
    plus.addEventListener('click', () => step(1));
    size.append(minus, ui.size, plus);
    dock.append(size);
    sep();
    ui.b = btn(ICONS.bold, 'Жирный · Ctrl+B', () => this.exec('bold'));
    ui.i = btn(ICONS.italic, 'Курсив · Ctrl+I', () => this.exec('italic'));
    ui.u = btn(ICONS.underline, 'Подчёркнутый · Ctrl+U', () => this.exec('underline'));
    ui.color = btn(ICONS.textColor + '<span class="dot"></span>', 'Цвет текста', (e, b) => this.swatches(b, WD_COLORS.map((c) => [c, c]), (c) => this.exec('foreColor', c)), 'wd-color');
    ui.mark = btn(ICONS.marker, 'Выделение цветом', (e, b) => this.swatches(b, WD_MARKS, (c) => this.exec('hiliteColor', c || 'transparent')));
    sep();
    ui.align = btn(ICONS.alignLeft, 'Выравнивание', (e, b) => this.menuAt(b, [
      ['По левому краю', () => this.setAlign('left'), 'Ctrl+L'],
      ['По центру', () => this.setAlign('center'), 'Ctrl+E'],
      ['По правому краю', () => this.setAlign('right'), 'Ctrl+R'],
      ['По ширине', () => this.setAlign('justify'), 'Ctrl+J'],
    ]));
    ui.bullets = btn(ICONS.listBullet, 'Маркированный список', () => this.toggleList('bullet'));
    ui.numbers = btn(ICONS.listNum, 'Нумерованный список', () => this.toggleList('number'));
    sep();
    btn(ICONS.image, 'Картинка', () => this.pickImage());
    btn(ICONS.table, 'Таблица', (e, b) => this.tablePicker(b));
    sep();
    btn(ICONS.hide, 'Скрыть панель · Ctrl+.', () => window.setDockHidden && window.setDockHidden(true));
    const handle = el('button', 'dock-handle');
    handle.dataset.tip = 'Показать панель · Ctrl+.';
    handle.addEventListener('click', () => window.setDockHidden && window.setDockHidden(false));
    this.root.append(dock, handle);
    this.dock = dock;
    this.dockMover = makeDockMovable(this.root, dock, null, 'docx');
  }

  // Меню открываем в сторону документа: пенал может стоять у любого края.
  menuAt(b, items) {
    const r = b.getBoundingClientRect();
    const edge = this.dockMover ? this.dockMover.edge() : 'bottom';
    const h = items.length * 31 + 12;
    if (edge === 'top') showMenu(r.left, r.bottom + 6, items);
    else if (edge === 'left') showMenu(r.right + 8, r.top, items);
    else if (edge === 'right') showMenu(r.left - 8 - 220, r.top, items);
    else showMenu(r.left, r.top - 8 - h + 12, items);
  }

  popover(anchor, content) {
    this.pop && this.pop.remove();
    const p = el('div', 'ctx wd-pop');
    p.append(content);
    p.addEventListener('pointerdown', (e) => e.preventDefault());
    this.root.append(p);
    const ar = anchor.getBoundingClientRect(), rr = this.root.getBoundingClientRect();
    const edge = this.dockMover ? this.dockMover.edge() : 'bottom';
    if (edge === 'left' || edge === 'right') {
      p.style.transform = 'translateY(-50%)';
      p.style.top = ar.top + ar.height / 2 - rr.top + 'px';
      if (edge === 'left') p.style.left = ar.right - rr.left + 10 + 'px';
      else { p.style.left = 'auto'; p.style.right = rr.right - ar.left + 10 + 'px'; }
    } else {
      p.style.left = ar.left + ar.width / 2 - rr.left + 'px';
      if (edge === 'top') p.style.top = ar.bottom - rr.top + 10 + 'px';
      else p.style.bottom = rr.bottom - ar.top + 10 + 'px';
    }
    this.pop = p;
    const off = (e) => { if (!p.contains(e.target) && e.target !== anchor && !anchor.contains(e.target)) { p.remove(); window.removeEventListener('pointerdown', off, true); } };
    setTimeout(() => window.addEventListener('pointerdown', off, true));
    return p;
  }

  swatches(anchor, list, fn) {
    const row = el('div', 'wd-sw');
    for (const [name, col] of list) {
      const b = el('button', 'sw', col ? '<i></i>' : '<i class="none"></i>');
      b.style.setProperty('--c', col || 'transparent');
      b.dataset.tip = col ? name : 'Без выделения';
      b.addEventListener('click', () => { this.pop && this.pop.remove(); fn(col); });
      row.append(b);
    }
    this.popover(anchor, row);
  }

  tablePicker(anchor) {
    const box = el('div', 'wd-grid');
    const label = el('div', 'wd-grid-label', 'Таблица');
    const grid = el('div', 'wd-grid-cells');
    for (let r = 1; r <= 8; r++) for (let c = 1; c <= 8; c++) {
      const cell = el('button');
      cell.dataset.r = r;
      cell.dataset.c = c;
      grid.append(cell);
    }
    grid.addEventListener('pointerover', (e) => {
      const t = e.target.closest('button');
      if (!t) return;
      const R = +t.dataset.r, C = +t.dataset.c;
      label.textContent = `${C} × ${R}`;
      for (const b of grid.children) b.classList.toggle('on', +b.dataset.r <= R && +b.dataset.c <= C);
    });
    grid.addEventListener('click', (e) => {
      const t = e.target.closest('button');
      if (!t) return;
      this.pop && this.pop.remove();
      this.insertTable(+t.dataset.r, +t.dataset.c);
    });
    box.append(label, grid);
    this.popover(anchor, box);
  }

  syncTools() {
    if (!this.ui || this.root.hidden) return;
    const sel = getSelection();
    if (!sel.rangeCount || !this.surface.contains(sel.anchorNode)) return;
    let n = sel.anchorNode;
    if (n.nodeType === 3) n = n.parentElement;
    const cs = getComputedStyle(n);
    const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
    this.ui.size.textContent = pt;
    this.ui.font.textContent = firstFam(cs.fontFamily);
    this.ui.font.style.fontFamily = famCss(firstFam(cs.fontFamily));
    this.ui.b.classList.toggle('on', +cs.fontWeight >= 600);
    this.ui.i.classList.toggle('on', cs.fontStyle === 'italic');
    let u = false;
    for (let e = n; e && e !== this.surface; e = e.parentElement) if (getComputedStyle(e).textDecorationLine.includes('underline')) { u = true; break; }
    this.ui.u.classList.toggle('on', u);
    this.ui.color.style.setProperty('--c', cs.color);
    const p = n.closest('p');
    if (p) {
      const st = this.styles[p.dataset.style || this.defPStyle];
      const nm = st ? st.name : 'normal';
      this.ui.style.textContent = STYLE_RU[nm] || (/^heading (\d)$/.test(nm) ? 'Заголовок ' + nm.slice(8) : /^toc (\d)$/.test(nm) ? 'Оглавление ' + nm.slice(4) : nm.charAt(0).toUpperCase() + nm.slice(1));
      const al = getComputedStyle(p).textAlign;
      this.ui.align.innerHTML = { center: ICONS.alignCenter, right: ICONS.alignRight, justify: ICONS.alignJustify }[al] || ICONS.alignLeft;
      this.ui.bullets.classList.toggle('on', this.listKind(p) === 'bullet');
      this.ui.numbers.classList.toggle('on', this.listKind(p) === 'number');
    }
  }

  // ---------- Совместимость с оболочкой ----------
  // ---------- Панель страниц ----------
  toggleSide(force) {
    const open = force != null ? force : !this.root.classList.contains('side-open');
    this.root.classList.toggle('side-open', open);
    if (this.ui && this.ui.pages) this.ui.pages.classList.toggle('on', open);
    if (open) {
      this.buildThumbs();
      setTimeout(() => {
        const t = this.thumbEls[this.curPage || 0];
        if (t) t.scrollIntoView({ block: 'center' });
        this.renderVisibleThumbs();
      }, 220);
    }
  }

  // Листы Word — не картинки, поэтому миниатюра — уменьшенная копия содержимого листа.
  // Рисуем только те, что видны в панели, и перерисовываем после правок.
  buildThumbs() {
    const pages = this.pageList || [];
    const X = PXPT;
    while (this.thumbEls.length > pages.length) this.thumbEls.pop().remove();
    pages.forEach((pg, i) => {
      const g = this.sections[pg.sec].pg;
      const k = 134 / Math.max(g.w * X, g.h * X * 0.78);
      let t = this.thumbEls[i];
      if (!t) {
        t = el('div', 'thumb');
        t.innerHTML = '<div class="t-paper"></div><div class="t-num"></div>';
        this.thumbs.append(t);
        this.thumbEls[i] = t;
      }
      t.dataset.i = i;
      const paper = t.firstChild;
      paper.style.width = Math.round(g.w * X * k) + 'px';
      paper.style.height = Math.round(g.h * X * k) + 'px';
      t.querySelector('.t-num').textContent = i + 1;
      t.dataset.dirty = '1';
    });
    const gap = (i) => {
      let g = this.gapEls[i];
      if (!g) {
        g = el('div', 'gap');
        const b = el('button', '', ICONS.plus);
        b.dataset.tip = 'Вставить пустую страницу';
        g.append(b);
        this.gapEls[i] = g;
      }
      g.dataset.at = i;
      return g;
    };
    const order = [gap(0)];
    this.thumbEls.slice(0, pages.length).forEach((t, i) => order.push(t, gap(i + 1)));
    this.gapEls.length = pages.length + 1;
    this.thumbs.replaceChildren(...order);
    this.markCurThumb(this.curPage || 0);
    this.renderVisibleThumbs();
  }

  // Пустая страница на месте промежутка: предыдущий лист заканчиваем разрывом,
  // на новом — пустой абзац с разрывом после него. Курсор ставим туда.
  insertPageAt(i) {
    const pages = this.pageList || [];
    const blocks = [...this.surface.children];
    const pb = () => { const s = el('span', 'atom pbreak'); s.contentEditable = 'false'; s.dataset.pb = '1'; return s; };
    const ended = (b) => b && ((b.tagName === 'P' && b.querySelector('.atom.pbreak')) || b.dataset.sect != null);
    const B = i < pages.length ? blocks.find((b) => b.offsetTop >= pages[i].top - 1) || null : null;
    const prev = B ? B.previousElementSibling : this.surface.lastElementChild;
    // Новый абзац «с новой страницы» (pageBreakBefore) — предыдущий лист не трогаем,
    // а разрыв после абзаца уводит остальной текст дальше.
    const blank = this.emptyP();
    if (B) {
      blank.append(pb());
      if (B !== blocks[0]) blank.dataset.pbb = '1';
      B.before(blank);
    } else {
      blank.dataset.pbb = '1';
      this.surface.append(blank);
    }
    void prev; void ended;
    this.surface.focus({ preventScroll: true });
    const r = document.createRange();
    r.setStart(blank, 0);
    r.collapse(true);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    this.touch();
    // Показываем лист, где новая страница реально оказалась (таблицу посередине не разрезать).
    setTimeout(() => {
      let at = i;
      (this.pageList || []).forEach((p, k) => { if (p.top <= blank.offsetTop + 1) at = k; });
      this.scrollToPage(at);
    }, 320);
    toast('Добавлена пустая страница');
  }

  renderVisibleThumbs() {
    if (!this.root.classList.contains('side-open')) return;
    const tr = this.thumbs.getBoundingClientRect();
    this.thumbEls.forEach((t, i) => {
      if (!t.dataset.dirty) return;
      const r = t.getBoundingClientRect();
      if (r.bottom > tr.top - 300 && r.top < tr.bottom + 300) this.renderThumb(i);
    });
  }

  renderThumb(i) {
    const t = this.thumbEls[i], pg = this.pageList && this.pageList[i];
    if (!t || !pg) return;
    delete t.dataset.dirty;
    const X = PXPT;
    const g = this.sections[pg.sec].pg;
    const W = g.w * X, H = g.h * X;
    const paper = t.firstChild;
    const k = parseFloat(paper.style.width) / W;
    const box = el('div', 'wd-thumb-page wd-surface');
    Object.assign(box.style, { width: W + 'px', height: H + 'px', transform: `scale(${k})` });
    const sheet = this.sheets.children[i];
    const left0 = sheet ? parseFloat(sheet.style.left) || 0 : 0;
    if (sheet) for (const c of sheet.children) box.append(c.cloneNode(true));
    for (const b of this.surface.children) {
      const top = b.offsetTop - pg.top;
      if (top + b.offsetHeight < 0 || top > H) continue;
      const c = b.cloneNode(true);
      Object.assign(c.style, { position: 'absolute', left: b.offsetLeft - left0 + 'px', top: top + 'px', width: b.offsetWidth + 'px', margin: '0' });
      c.classList.remove('wd-split');
      c.removeAttribute('contenteditable');
      box.append(c);
    }
    paper.replaceChildren(box);
  }

  markCurThumb(i) {
    this.curPage = i;
    this.thumbEls.forEach((t, j) => t.classList.toggle('cur', j === i));
  }

  scrollToPage(i) {
    const pg = this.pageList && this.pageList[i];
    if (!pg) return;
    this.viewer.scrollTo({ top: pg.top * this.zoom + this.col.offsetTop + 28 - 16, behavior: 'smooth' });
    this.markCurThumb(i);
  }

  endEdit() {}
  onKey(e) {
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && (e.key === '=' || e.key === '+')) { this.setZoom(this.zoom * 1.1); return true; }
    if (ctrl && e.key === '-') { this.setZoom(this.zoom / 1.1); return true; }
    if (ctrl && e.key === '0') { this.fitWidth(); return true; }
    return false;
  }

  destroy() {
    document.removeEventListener('selectionchange', this.onSel);
    if (this.onFonts) document.fonts.removeEventListener('loadingdone', this.onFonts);
    if (this.mo) this.mo.disconnect();
    this.ro.disconnect();
    if (this.mediaCache) for (const p of this.mediaCache.values()) if (p) p.then((u) => URL.revokeObjectURL(u));
    this.root.remove();
  }
}

// Встроенные стили, если в документе их нет.
const STYLE_TEMPLATES = {
  title: '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>',
  'heading 1': '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>',
  'heading 2': '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>',
  'heading 3': '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>',
};

// Пустой документ Word: А4, поля как в российских требованиях (30/15/20/20 мм).
async function makeBlankDocx() {
  const zip = new JSZip();
  const W = NS.w;
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${STYLE_TEMPLATES.title}${STYLE_TEMPLATES['heading 1']}${STYLE_TEMPLATES['heading 2']}${STYLE_TEMPLATES['heading 3']}<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
  return zip.generateAsync({ type: 'uint8array' });
}

(() => {
  const s = document.createElement('style');
  s.textContent = WD_CSS;
  document.head.append(s);
})();

// Панель страниц у Word ведёт себя так же, как у PDF: зеркально в двух окнах и перетаскивается вбок.
DocxEditor.prototype.setSideRight = PdfEditor.prototype.setSideRight;
DocxEditor.prototype.bindSideGrip = PdfEditor.prototype.bindSideGrip;
DocxEditor.prototype.isShown = PdfEditor.prototype.isShown;
