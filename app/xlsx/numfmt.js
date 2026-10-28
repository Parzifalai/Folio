// Форматы чисел Excel: 0,00 · # ##0 · 0% · даты · денежный. Понимаем основные приёмы,
// редкие (дроби, условия) показываем как обычное число.

const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  12: '# ?/?', 13: '# ??/??', 14: 'dd.mm.yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'dd.mm.yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
};
const MONTHS = {
  ru: ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'],
  rug: ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'],
  en: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
};
const DAYS = {
  ru: ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'],
  en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
};
const COLOR_NAMES = { red: '#ff0000', blue: '#0000ff', green: '#00a000', black: '#000000', white: '#ffffff', magenta: '#ff00ff', yellow: '#c0a000', cyan: '#00a0a0' };

const fmtCache = new Map();

function splitSections(fmt) {
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') q = !q;
    if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

// Разбор одной секции: литералы, плейсхолдеры, признаки даты.
function compileSection(sec) {
  let color = null, cond = null, currency = null;
  sec = sec.replace(/\[([^\]]*)\]/g, (m, inner) => {
    const lo = inner.toLowerCase();
    if (COLOR_NAMES[lo]) { color = COLOR_NAMES[lo]; return ''; }
    if (/^color\d+$/.test(lo)) return '';
    if (/^[<>=]/.test(inner)) { cond = inner; return ''; }
    if (inner.startsWith('$')) { const sym = inner.slice(1).split('-')[0]; currency = sym; return sym ? '"' + sym + '"' : ''; }
    if (/^h+$|^m+$|^s+$/i.test(inner)) return '\u0000' + inner.toLowerCase() + '\u0000'; // [h] — прошедшие часы
    return '';
  });
  const toks = [];
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); toks.push({ lit: sec.slice(i + 1, j < 0 ? undefined : j) }); i = j < 0 ? sec.length : j; continue; }
    if (ch === '\\') { toks.push({ lit: sec[i + 1] || '' }); i++; continue; }
    if (ch === '_') { toks.push({ lit: ' ' }); i++; continue; }
    if (ch === '*') { i++; continue; }
    if (ch === '\u0000') { const j = sec.indexOf('\u0000', i + 1); toks.push({ d: '[' + sec.slice(i + 1, j) + ']' }); i = j; continue; }
    const rest = sec.slice(i);
    let m;
    if ((m = /^(AM\/PM|A\/P)/i.exec(rest))) { toks.push({ d: 'ampm' }); i += m[0].length - 1; continue; }
    if ((m = /^(y+|m+|d+|h+|s+)/i.exec(rest))) { toks.push({ d: m[0].toLowerCase() }); i += m[0].length - 1; continue; }
    if ((m = /^[eE][+-]/.exec(rest))) { toks.push({ e: m[0][1] }); i += 1; continue; }
    if ('0#?'.includes(ch)) { toks.push({ p: ch }); continue; }
    if (ch === '.') { toks.push({ dot: true }); continue; }
    if (ch === ',') { toks.push({ comma: true }); continue; }
    if (ch === '%') { toks.push({ pct: true, lit: '%' }); continue; }
    if (ch === '@') { toks.push({ at: true }); continue; }
    if (ch === '/' && toks.some((t) => t.p)) { toks.push({ lit: '/' }); continue; }
    toks.push({ lit: ch });
  }
  const isDate = toks.some((t) => t.d);
  // «m» рядом с часами или секундами — минуты.
  if (isDate) {
    toks.forEach((t, i) => {
      if (!t.d || t.d[0] !== 'm' || t.d.length > 2) return;
      const prev = toks.slice(0, i).reverse().find((x) => x.d), next = toks.slice(i + 1).find((x) => x.d);
      if ((prev && /^h|^\[h/.test(prev.d)) || (next && /^s/.test(next.d))) t.min = true;
    });
  }
  return { toks, color, cond, isDate, isText: toks.some((t) => t.at), currency };
}

function compileFmt(fmt) {
  let c = fmtCache.get(fmt);
  if (!c) {
    const secs = splitSections(fmt).map(compileSection);
    c = { secs, isDate: secs[0] && secs[0].isDate, general: /^general$/i.test(fmt.trim()) || fmt === '' };
    fmtCache.set(fmt, c);
  }
  return c;
}

const fmtIsDate = (fmt) => !!fmt && compileFmt(fmt).isDate;

function groupInt(s, grp) {
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, grp);
}

function numSection(v, sec, dec, grp) {
  const toks = sec.toks;
  if (toks.some((t) => t.pct)) v *= 100 * toks.filter((t) => t.pct).length;
  const firstP = toks.findIndex((t) => t.p || t.dot);
  const lastP = toks.length - 1 - [...toks].reverse().findIndex((t) => t.p || t.dot || (t.comma && false));
  if (firstP < 0) return toks.map((t) => t.lit || '').join('');
  const body = toks.slice(firstP, lastP + 1);
  const pre = toks.slice(0, firstP).map((t) => t.lit || '').join('');
  const post = toks.slice(lastP + 1).filter((t) => !t.comma).map((t) => t.lit || '').join('');
  // Запятые сразу после последнего плейсхолдера — деление на 1000.
  let tail = 0;
  for (let k = lastP + 1; k < toks.length && toks[k].comma; k++) tail++;
  v /= Math.pow(1000, tail);
  const ei = body.findIndex((t) => t.e);
  const intPart = [], fracPart = [];
  let seenDot = false, grouping = false;
  for (const t of ei >= 0 ? body.slice(0, ei) : body) {
    if (t.dot) { seenDot = true; continue; }
    if (t.comma) { if (!seenDot) grouping = true; continue; }
    if (t.p) (seenDot ? fracPart : intPart).push(t.p);
  }
  const maxDec = fracPart.length, minDec = fracPart.filter((p) => p === '0').length;
  const minInt = intPart.filter((p) => p === '0').length;
  let s;
  if (ei >= 0) {
    const expDigits = body.slice(ei + 1).filter((t) => t.p).length || 2;
    let [m, e] = Math.abs(v).toExponential(maxDec).split('e');
    let ex = parseInt(e, 10);
    const sign = ex < 0 ? '-' : body[ei].e === '+' ? '+' : '';
    s = m.replace('.', dec) + 'E' + sign + String(Math.abs(ex)).padStart(expDigits, '0');
  } else {
    let fixed = Math.abs(v).toFixed(maxDec);
    let [ip, fp = ''] = fixed.split('.');
    while (fp.length > minDec && fp.endsWith('0')) fp = fp.slice(0, -1);
    if (ip === '0' && minInt === 0) ip = '';
    ip = ip.padStart(minInt, '0');
    if (grouping) ip = groupInt(ip, grp);
    s = ip + (fp || (maxDec && minDec) ? dec + fp : fracPart.length && !fp ? '' : '');
    if (fracPart.length && !fp && seenDot && minDec === 0) s = ip;
    if (!s) s = '0';
  }
  return pre + s + post;
}

function dateSection(v, sec, dec, ru) {
  const D = serialToDate(v);
  let secs = D.secs;
  const hasAmpm = sec.toks.some((t) => t.d === 'ampm');
  const fracTok = sec.toks.findIndex((t) => t.dot);
  let out = '';
  sec.toks.forEach((t, i) => {
    if (t.lit != null && !t.d) { out += t.lit; return; }
    if (t.dot && !(sec.toks[i + 1] && sec.toks[i + 1].p === '0')) { out += '.'; return; }
    if (t.dot) {
      const zs = [];
      for (let k = i + 1; k < sec.toks.length && sec.toks[k].p === '0'; k++) zs.push(0);
      const frac = (v * 86400) % 1;
      out += dec + frac.toFixed(zs.length).slice(2);
      return;
    }
    if (t.p) return;
    if (!t.d) return;
    const H = Math.floor(secs / 3600), M = Math.floor(secs / 60) % 60, S = secs % 60;
    const d = t.d;
    if (d === 'ampm') { out += H < 12 ? 'AM' : 'PM'; return; }
    if (d.startsWith('[')) {
      const tot = v * 86400;
      out += d[1] === 'h' ? Math.floor(tot / 3600) : d[1] === 'm' ? Math.floor(tot / 60) : Math.floor(tot);
      return;
    }
    if (d[0] === 'y') { out += d.length <= 2 ? String(D.y % 100).padStart(2, '0') : D.y; return; }
    if (d[0] === 'd') {
      if (d.length <= 2) out += d.length === 2 ? String(D.d).padStart(2, '0') : D.d;
      else { const n = DAYS[ru ? 'ru' : 'en'][D.wd]; out += d.length === 3 ? n.slice(0, ru ? 2 : 3) : n; }
      return;
    }
    if (d[0] === 'm' && t.min) { out += d.length === 2 ? String(M).padStart(2, '0') : M; return; }
    if (d[0] === 'm') {
      if (d.length <= 2) { out += d.length === 2 ? String(D.m).padStart(2, '0') : D.m; return; }
      const names = ru ? (sec.toks.some((x) => x.d && x.d[0] === 'd' && x.d.length <= 2) ? MONTHS.rug : MONTHS.ru) : MONTHS.en;
      const n = names[D.m - 1];
      out += d.length === 3 ? n.slice(0, 3) : d.length === 5 ? n[0] : n;
      return;
    }
    if (d[0] === 'h') { const h = hasAmpm ? ((H + 11) % 12) + 1 : H; out += d.length === 2 ? String(h).padStart(2, '0') : h; return; }
    if (d[0] === 's') { out += d.length === 2 ? String(fracTok >= 0 ? Math.floor(S) : S).padStart(2, '0') : S; return; }
  });
  void secs;
  return out;
}

// Значение → текст ячейки. dec/grp — разделители по языку интерфейса.
function formatValue(v, fmt, dec = ',', grp = ' ') {
  if (v == null) return { text: '' };
  if (isErr(v)) return { text: v.err, err: true };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', bool: true };
  const c = compileFmt(fmt || 'General');
  if (typeof v === 'string') {
    const ts = c.secs[3] || c.secs.find((s) => s.isText);
    if (ts) return { text: ts.toks.map((t) => (t.at ? v : t.lit || '')).join(''), color: ts.color };
    return { text: v };
  }
  if (c.general) return { text: fmtGeneral(v, dec), num: true };
  let sec = c.secs[0], val = v, sign = '';
  const conds = c.secs.filter((s) => s.cond);
  if (conds.length) {
    const ok = (s) => { const m = /^(<=|>=|<>|=|<|>)(-?[\d.]+)$/.exec(s.cond); if (!m) return false; const n = +m[2]; return { '<': v < n, '>': v > n, '=': v === n, '<=': v <= n, '>=': v >= n, '<>': v !== n }[m[1]]; };
    sec = c.secs.find((s) => s.cond && ok(s)) || c.secs.find((s) => !s.cond) || sec;
    val = Math.abs(v);
    if (v < 0 && sec === c.secs[0] && !sec.cond) sign = '-';
  } else if (v < 0 && c.secs[1] && c.secs[1].toks.length) { sec = c.secs[1]; val = -v; }
  else if (v === 0 && c.secs[2] && c.secs[2].toks.length) sec = c.secs[2];
  else if (v < 0) { val = -v; sign = '-'; }
  if (sec.isText && !sec.isDate && !sec.toks.some((t) => t.p)) return { text: sec.toks.map((t) => (t.at ? fmtGeneral(v, dec) : t.lit || '')).join(''), color: sec.color };
  if (sec.isDate) {
    if (v < 0 || v > 2958465) return { text: '#'.repeat(8), num: true };
    return { text: dateSection(v, sec, dec, dec === ','), color: sec.color, num: true };
  }
  if (!sec.toks.some((t) => t.p)) return { text: sign + sec.toks.map((t) => t.lit || '').join(''), color: sec.color, num: true };
  const s = numSection(val, sec, dec, grp);
  return { text: (sign && /[1-9]/.test(s) ? sign : '') + s, color: sec.color, num: true };
}

// Для поля ввода: число по-человечески (без разделителей разрядов, все знаки).
function editText(v, fmt, dec) {
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v.err;
  if (typeof v === 'number') {
    if (fmtIsDate(fmt)) {
      const D = serialToDate(v);
      const date = `${String(D.d).padStart(2, '0')}.${String(D.m).padStart(2, '0')}.${D.y}`;
      if (D.secs) return date + ' ' + Math.floor(D.secs / 3600) + ':' + String(Math.floor(D.secs / 60) % 60).padStart(2, '0') + (D.secs % 60 ? ':' + String(D.secs % 60).padStart(2, '0') : '');
      return date;
    }
    if (fmt && /%/.test(fmt.split(';')[0].replace(/"[^"]*"/g, ''))) return fmtGeneral(v * 100, dec) + '%';
    const s = String(+v.toPrecision(15));
    return dec === '.' ? s : s.replace('.', dec);
  }
  return String(v);
}
