// Формулы Excel: разбор, вычисление, сдвиг ссылок. Внутри — английские имена функций и запятые,
// как в самом файле .xlsx; ввод понимает и русские имена (СУММ, ЕСЛИ…) и точку с запятой.

const XL_MAXR = 1048576, XL_MAXC = 16384;

class XErr {
  constructor(code) { this.err = code; }
  toString() { return this.err; }
}
const E = {
  div0: new XErr('#DIV/0!'), value: new XErr('#VALUE!'), ref: new XErr('#REF!'),
  name: new XErr('#NAME?'), num: new XErr('#NUM!'), na: new XErr('#N/A'), nul: new XErr('#NULL!'),
};
const errOf = (code) => Object.values(E).find((e) => e.err === code) || new XErr(code);
const isErr = (v) => v instanceof XErr;

// ---------- Адреса ----------
function colName(c) {
  let s = '';
  for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}
function colIndex(s) {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}
function addr(r, c) { return colName(c) + (r + 1); }
function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
  return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null;
}
function parseRange(s) {
  const [a, b] = s.split(':');
  const A = parseAddr(a), B = b ? parseAddr(b) : A;
  if (!A || !B) return null;
  return { r1: Math.min(A.r, B.r), c1: Math.min(A.c, B.c), r2: Math.max(A.r, B.r), c2: Math.max(A.c, B.c) };
}
function rangeStr(g) {
  const a = addr(g.r1, g.c1), b = addr(g.r2, g.c2);
  return a === b ? a : a + ':' + b;
}
const quoteSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");

// ---------- Русские имена функций ----------
const RU_FN = {
  СУММ: 'SUM', СРЗНАЧ: 'AVERAGE', МИН: 'MIN', МАКС: 'MAX', СЧЁТ: 'COUNT', СЧЕТ: 'COUNT', СЧЁТЗ: 'COUNTA', СЧЕТЗ: 'COUNTA',
  СЧИТАТЬПУСТОТЫ: 'COUNTBLANK', ЕСЛИ: 'IF', ЕСЛИОШИБКА: 'IFERROR', И: 'AND', ИЛИ: 'OR', НЕ: 'NOT', ОКРУГЛ: 'ROUND',
  ОКРУГЛВВЕРХ: 'ROUNDUP', ОКРУГЛВНИЗ: 'ROUNDDOWN', ABS: 'ABS', ЦЕЛОЕ: 'INT', ОСТАТ: 'MOD', СТЕПЕНЬ: 'POWER', КОРЕНЬ: 'SQRT',
  ПРОИЗВЕД: 'PRODUCT', СУММЕСЛИ: 'SUMIF', СЧЁТЕСЛИ: 'COUNTIF', СЧЕТЕСЛИ: 'COUNTIF', СРЗНАЧЕСЛИ: 'AVERAGEIF',
  СУММЕСЛИМН: 'SUMIFS', СЧЁТЕСЛИМН: 'COUNTIFS', СЧЕТЕСЛИМН: 'COUNTIFS', ВПР: 'VLOOKUP', ГПР: 'HLOOKUP', ИНДЕКС: 'INDEX',
  ПОИСКПОЗ: 'MATCH', СЦЕП: 'CONCAT', СЦЕПИТЬ: 'CONCATENATE', ДЛСТР: 'LEN', ПРОПИСН: 'UPPER', СТРОЧН: 'LOWER',
  СЖПРОБЕЛЫ: 'TRIM', ЛЕВСИМВ: 'LEFT', ПРАВСИМВ: 'RIGHT', ПСТР: 'MID', ТЕКСТ: 'TEXT', СЕГОДНЯ: 'TODAY', ТДАТА: 'NOW',
  ГОД: 'YEAR', МЕСЯЦ: 'MONTH', ДЕНЬ: 'DAY', ДАТА: 'DATE', ЕПУСТО: 'ISBLANK', ЕЧИСЛО: 'ISNUMBER', ЕТЕКСТ: 'ISTEXT',
  ЕОШИБКА: 'ISERROR', ИСТИНА: 'TRUE', ЛОЖЬ: 'FALSE', ЗНАЧЕН: 'VALUE', СУММПРОИЗВ: 'SUMPRODUCT', МЕДИАНА: 'MEDIAN',
  РАНГ: 'RANK', ОКРУГЛТ: 'MROUND', ЗНАК: 'SIGN', ПИ: 'PI', СЛЧИС: 'RAND', ЕСЛИМН: 'IFS', МАКСЕСЛИ: 'MAXIFS', МИНЕСЛИ: 'MINIFS',
  СТРОКА: 'ROW', СТОЛБЕЦ: 'COLUMN', ЗАМЕНИТЬ: 'REPLACE', ПОДСТАВИТЬ: 'SUBSTITUTE', НАЙТИ: 'FIND', ПОИСК: 'SEARCH',
  ЧПС: 'NPV', ВСД: 'IRR', ПЛТ: 'PMT', БС: 'FV', ПС: 'PV', ДНЕЙ: 'DAYS', ЧАС: 'HOUR', МИНУТЫ: 'MINUTE', СЕКУНДЫ: 'SECOND', ДЕНЬНЕД: 'WEEKDAY', ВЫБОР: 'CHOOSE', СТАНДОТКЛОН: 'STDEV',
};

// ---------- Разбор ----------
const RE_SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\u0400-\u04FF][\w.\u0400-\u04FF]*)!`;
const RE_CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const RE_REF = new RegExp(String.raw`^(${RE_SHEET})?(${RE_CELL}(?::${RE_CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(\u0400-\u04FF])`);
const RE_ERR = /^(#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A|#GETTING_DATA)/i;

function tokenize(src) {
  const out = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    const rest = src.slice(i);
    let m;
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= src.length) throw new Error('quote');
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      out.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    if ((m = RE_ERR.exec(rest))) { out.push({ t: 'err', v: m[1].toUpperCase() }); i += m[1].length; continue; }
    if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?|^\.\d+([eE][+-]?\d+)?/.exec(rest)) && !RE_REF.test(rest)) {
      out.push({ t: 'num', v: parseFloat(m[0]), raw: m[0] });
      i += m[0].length;
      continue;
    }
    if ((m = RE_REF.exec(rest))) { out.push({ t: 'ref', v: m[0] }); i += m[0].length; continue; }
    if ((m = /^(?:_xlfn\.|_xlws\.)?[A-Za-z_\u0400-\u04FF][\w.\u0400-\u04FF]*(?=\s*\()/.exec(rest))) { out.push({ t: 'fn', v: m[0] }); i += m[0].length; continue; }
    if ((m = /^[A-Za-z_\\\u0400-\u04FF][\w.\u0400-\u04FF]*/.exec(rest))) {
      const u = m[0].toUpperCase();
      if (u === 'TRUE' || u === 'ИСТИНА') out.push({ t: 'bool', v: true });
      else if (u === 'FALSE' || u === 'ЛОЖЬ') out.push({ t: 'bool', v: false });
      else out.push({ t: 'name', v: m[0] });
      i += m[0].length;
      continue;
    }
    if ((m = /^(<>|<=|>=)/.exec(rest))) { out.push({ t: 'op', v: m[0] }); i += 2; continue; }
    if ('+-*/^&=<>%:'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
    if (ch === '(' || ch === ')') { out.push({ t: ch }); i++; continue; }
    if (ch === ',' || ch === ';') { out.push({ t: ',' }); i++; continue; }
    if (ch === '{' || ch === '}') throw new Error('array');
    throw new Error('char ' + ch);
  }
  return out;
}

function parseRefText(text) {
  let sheet = null, body = text;
  const bang = text.lastIndexOf('!');
  if (bang >= 0) {
    sheet = text.slice(0, bang);
    if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    body = text.slice(bang + 1);
  }
  const part = (p) => {
    const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(p);
    return { ac: !!m[1] && !!m[2], c: m[2] ? colIndex(m[2]) : null, ar: !!m[3] && !!m[4] || (!m[2] && !!m[1]), r: m[4] ? +m[4] - 1 : null };
  };
  const [a, b] = body.split(':');
  if (b == null) return { t: 'ref', sheet, a: part(a) };
  return { t: 'range', sheet, a: part(a), b: part(b) };
}

const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

function parseFormula(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  function primary() {
    const t = next();
    if (!t) throw new Error('end');
    if (t.t === 'num') return { t: 'num', v: t.v, raw: t.raw };
    if (t.t === 'str') return { t: 'str', v: t.v };
    if (t.t === 'bool') return { t: 'bool', v: t.v };
    if (t.t === 'err') return { t: 'err', v: t.v };
    if (t.t === 'ref') return parseRefText(t.v);
    if (t.t === 'name') return { t: 'name', v: t.v };
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) return { t: 'un', op: t.v, a: unary() };
    if (t.t === '(') {
      const e = expr(0);
      if (!peek() || peek().t !== ')') throw new Error(')');
      next();
      return { t: 'paren', a: e };
    }
    if (t.t === 'fn') {
      next(); // (
      const args = [];
      if (peek() && peek().t === ')') { next(); }
      else {
        for (;;) {
          if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ t: 'missing' });
          else args.push(expr(0));
          const s = next();
          if (!s) throw new Error(')');
          if (s.t === ')') break;
          if (s.t !== ',') throw new Error(',');
        }
      }
      let name = t.v;
      const bare = name.replace(/^_xl(fn|ws)\./, '').toUpperCase();
      if (RU_FN[bare]) name = RU_FN[bare];
      return { t: 'fn', name, args };
    }
    throw new Error('unexpected');
  }
  function unary() {
    let e = primary();
    while (peek() && peek().t === 'op' && peek().v === '%') { next(); e = { t: 'pct', a: e }; }
    return e;
  }
  function expr(min) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || !(t.v in PREC) || PREC[t.v] <= min) break;
      next();
      const right = expr(PREC[t.v]);
      left = { t: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }
  const ast = expr(0);
  if (p < toks.length) throw new Error('rest');
  return ast;
}

// ---------- Обратно в текст ----------
function partStr(x) {
  return (x.c != null ? (x.ac ? '$' : '') + colName(x.c) : '') + (x.r != null ? (x.ar ? '$' : '') + (x.r + 1) : '');
}
function fmtNumLit(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(+n.toPrecision(15));
}
function astStr(n) {
  switch (n.t) {
    case 'num': return n.raw || fmtNumLit(n.v);
    case 'str': return '"' + n.v.replace(/"/g, '""') + '"';
    case 'bool': return n.v ? 'TRUE' : 'FALSE';
    case 'err': return n.v;
    case 'name': return n.v;
    case 'missing': return '';
    case 'ref': return n.bad ? (n.sheet ? quoteSheet(n.sheet) + '!' : '') + '#REF!' : (n.sheet ? quoteSheet(n.sheet) + '!' : '') + partStr(n.a);
    case 'range': return n.bad ? (n.sheet ? quoteSheet(n.sheet) + '!' : '') + '#REF!' : (n.sheet ? quoteSheet(n.sheet) + '!' : '') + partStr(n.a) + ':' + partStr(n.b);
    case 'un': return n.op + astStr(n.a);
    case 'pct': return astStr(n.a) + '%';
    case 'paren': return '(' + astStr(n.a) + ')';
    case 'bin': return astStr(n.a) + n.op + astStr(n.b);
    case 'fn': return n.name + '(' + n.args.map(astStr).join(',') + ')';
  }
  return '';
}

// Обход всех ссылок формулы.
function eachRef(n, fn) {
  if (!n) return;
  if (n.t === 'ref' || n.t === 'range') fn(n);
  else if (n.t === 'fn') n.args.forEach((a) => eachRef(a, fn));
  else if (n.a) { eachRef(n.a, fn); if (n.b && n.t === 'bin') eachRef(n.b, fn); }
}
const cloneAst = (n) => JSON.parse(JSON.stringify(n));

// Копирование формулы на dr строк и dc столбцов: относительные части едут, $-части стоят.
function shiftFormula(text, dr, dc) {
  let ast;
  try { ast = parseFormula(text); } catch { return text; }
  eachRef(ast, (n) => {
    for (const x of n.t === 'range' ? [n.a, n.b] : [n.a]) {
      if (x.r != null && !x.ar) { x.r += dr; if (x.r < 0 || x.r >= XL_MAXR) n.bad = true; }
      if (x.c != null && !x.ac) { x.c += dc; if (x.c < 0 || x.c >= XL_MAXC) n.bad = true; }
    }
  });
  return astStr(ast);
}

// Вставка/удаление строк или столбцов: сдвигаем ссылки на этот лист (и абсолютные тоже).
// axis 'r' | 'c', at — первый индекс, n > 0 вставка, n < 0 удаление |n| штук.
function adjustAst(ast, onSheet, axis, at, n) {
  let changed = false;
  eachRef(ast, (ref) => {
    if (!onSheet(ref.sheet) || ref.bad) return;
    const k = axis;
    const pts = ref.t === 'range' ? [ref.a, ref.b] : [ref.a];
    if (pts.some((x) => x[k] == null)) return; // A:A при вставке строк не меняется
    if (n > 0) {
      for (const x of pts) if (x[k] >= at) { x[k] += n; changed = true; }
    } else {
      const d = -n, end = at + d - 1;
      if (ref.t === 'ref') {
        const v = ref.a[k];
        if (v >= at && v <= end) { ref.bad = true; changed = true; }
        else if (v > end) { ref.a[k] -= d; changed = true; }
      } else {
        let lo = Math.min(ref.a[k], ref.b[k]), hi = Math.max(ref.a[k], ref.b[k]);
        if (lo >= at && hi <= end) { ref.bad = true; changed = true; return; }
        const nlo = lo > end ? lo - d : lo >= at ? at : lo;
        const nhi = hi > end ? hi - d : hi >= at ? at - 1 : hi;
        if (nlo !== lo || nhi !== hi) changed = true;
        const aLo = ref.a[k] <= ref.b[k];
        (aLo ? ref.a : ref.b)[k] = nlo;
        (aLo ? ref.b : ref.a)[k] = nhi;
      }
    }
  });
  return changed;
}

// Переименование листа в ссылках.
function renameSheetAst(ast, from, to) {
  let changed = false;
  eachRef(ast, (ref) => { if (ref.sheet && ref.sheet.toLowerCase() === from.toLowerCase()) { ref.sheet = to; changed = true; } });
  return changed;
}

// ---------- Вычисление ----------
// ctx: { cell(sheet, r, c) → значение, sheet — текущий лист, dims(sheet) → {r, c} — заполненная область, name(n) → ast|null }
class Matrix {
  constructor(rows) { this.rows = rows; }
  *values() { for (const row of this.rows) for (const v of row) yield v; }
  get h() { return this.rows.length; }
  get w() { return this.rows[0] ? this.rows[0].length : 0; }
}

function toNum(v) {
  if (v instanceof Matrix) v = v.rows[0] && v.rows[0][0];
  if (isErr(v)) throw v;
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const s = String(v).trim().replace(/\s/g, '').replace(',', '.');
  if (/%$/.test(s) && isFinite(parseFloat(s))) return parseFloat(s) / 100;
  const d = parseDateText(String(v));
  if (d != null) return d;
  if (s !== '' && isFinite(+s)) return +s;
  throw E.value;
}
function toStr(v) {
  if (v instanceof Matrix) v = v.rows[0] && v.rows[0][0];
  if (isErr(v)) throw v;
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return fmtGeneral(v, '.');
  return String(v);
}
function toBool(v) {
  if (v instanceof Matrix) v = v.rows[0] && v.rows[0][0];
  if (isErr(v)) throw v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null || v === '') return false;
  const u = String(v).toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  throw E.value;
}
const scalar = (v) => (v instanceof Matrix ? (v.rows[0] ? v.rows[0][0] : null) : v);

// Excel «Общий»: до ~11 значащих цифр.
function fmtGeneral(n, dec = '.') {
  if (!isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const a = Math.abs(n);
  let s;
  if (a >= 1e11 || a < 1e-9) {
    s = n.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E').replace(/E\+?(-?)(\d)$/, 'E+$1' + '0$2').replace('E+-', 'E-');
  } else {
    const digits = Math.max(0, 10 - Math.floor(Math.log10(a)));
    s = String(+n.toFixed(Math.min(digits, 15)));
    if (/e/.test(s)) s = n.toFixed(Math.min(digits, 15)).replace(/\.?0+$/, '');
  }
  return dec === '.' ? s : s.replace('.', dec);
}

// Даты: серийный номер Excel (1900, с «лишним» 29.02.1900).
function dateToSerial(y, m, d) {
  const ms = Date.UTC(y, m - 1, d);
  let s = ms / 86400000 + 25569;
  if (s < 61) s -= 1;
  return s;
}
function serialToDate(s) {
  let n = Math.floor(s);
  if (n < 61) n += 1;
  const d = new Date((n - 25569) * 86400000);
  const frac = s - Math.floor(s);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay(), secs: Math.round(frac * 86400) };
}
function parseDateText(s) {
  s = String(s).trim();
  let m = /^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  let y, mo, d;
  if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) { y = +m[1]; mo = +m[2]; d = +m[3]; m = [m[0], 0, 0, 0, m[4], m[5], m[6]]; }
  else return null;
  if (y < 100) y += y < 30 ? 2000 : 1900;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  let v = dateToSerial(y, mo, d);
  if (m[4] != null) v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400;
  return v;
}

function cmp(a, b) {
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 0);
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return a < b ? -1 : a > b ? 1 : 0;
}

// Условие для СУММЕСЛИ/СЧЁТЕСЛИ: ">5", "<>0", "яблоко", "а*".
function criteria(c) {
  c = scalar(c);
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => cmp(v, c) === 0 && v != null;
  const s = String(c ?? '');
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(s);
  const op = m[1] || '=', raw = m[2];
  const num = raw.trim() !== '' && isFinite(+raw.replace(',', '.')) ? +raw.replace(',', '.') : null;
  if (num != null) {
    return (v) => {
      if (typeof v !== 'number') return op === '<>' ? true : false;
      return { '=': v === num, '<>': v !== num, '<': v < num, '>': v > num, '<=': v <= num, '>=': v >= num }[op];
    };
  }
  if (/[*?]/.test(raw) && (op === '=' || op === '<>')) {
    const re = new RegExp('^' + raw.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'is');
    return (v) => (op === '=') === re.test(v == null ? '' : String(v));
  }
  if (raw === '' && op === '=') return (v) => v == null || v === '';
  if (raw === '' && op === '<>') return (v) => v != null && v !== '';
  return (v) => {
    const r = cmp(typeof v === 'number' ? String(v) : v, raw);
    return { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[op];
  };
}

function flat(args, cb) {
  for (const a of args) {
    if (a instanceof Matrix) { for (const v of a.values()) cb(v, true); }
    else cb(a, false);
  }
}
function nums(args) {
  const out = [];
  flat(args, (v, inRange) => {
    if (isErr(v)) throw v;
    if (inRange) { if (typeof v === 'number') out.push(v); }
    else if (v != null) out.push(toNum(v));
  });
  return out;
}
const asMatrix = (v) => (v instanceof Matrix ? v : new Matrix([[v]]));
const roundTo = (x, d, mode) => {
  const k = Math.pow(10, d);
  const y = x * k;
  const r = mode === 'up' ? Math.sign(y) * Math.ceil(Math.abs(y) - 1e-9) : mode === 'down' ? Math.sign(y) * Math.floor(Math.abs(y) + 1e-9) : Math.sign(y) * Math.round(Math.abs(y) + 1e-9);
  return r / k;
};

function lookupIdx(list, key, approx) {
  if (!approx) {
    const crit = typeof key === 'string' && /[*?]/.test(key) ? criteria(key) : null;
    for (let i = 0; i < list.length; i++) if (crit ? crit(list[i]) : list[i] != null && cmp(list[i], key) === 0) return i;
    return -1;
  }
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    if (list[i] == null) continue;
    if (cmp(list[i], key) <= 0) best = i; else break;
  }
  return best;
}

const FN = {
  SUM: (a) => nums(a).reduce((s, x) => s + x, 0),
  AVERAGE: (a) => { const n = nums(a); if (!n.length) throw E.div0; return n.reduce((s, x) => s + x, 0) / n.length; },
  MIN: (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; },
  COUNT: (a) => { let k = 0; flat(a, (v, inR) => { if (typeof v === 'number' || (!inR && v != null && !isErr(v) && (() => { try { toNum(v); return true; } catch { return false; } })())) k++; }); return k; },
  COUNTA: (a) => { let k = 0; flat(a, (v) => { if (v != null && v !== '') k++; }); return k; },
  COUNTBLANK: (a) => { let k = 0; flat(a, (v) => { if (v == null || v === '') k++; }); return k; },
  PRODUCT: (a) => nums(a).reduce((s, x) => s * x, 1),
  MEDIAN: (a) => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw E.num; const h = n.length >> 1; return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2; },
  STDEV: (a) => { const n = nums(a); if (n.length < 2) throw E.div0; const m = n.reduce((s, x) => s + x, 0) / n.length; return Math.sqrt(n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1)); },
  AND: (a) => { let r = true; flat(a, (v, inR) => { if (isErr(v)) throw v; if (inR && typeof v !== 'boolean' && typeof v !== 'number') return; if (v != null && !toBool(v)) r = false; }); return r; },
  OR: (a) => { let r = false; flat(a, (v, inR) => { if (isErr(v)) throw v; if (inR && typeof v !== 'boolean' && typeof v !== 'number') return; if (v != null && toBool(v)) r = true; }); return r; },
  NOT: (a) => !toBool(a[0]),
  ROUND: (a) => roundTo(toNum(a[0]), toNum(a[1] ?? 0)),
  ROUNDUP: (a) => roundTo(toNum(a[0]), toNum(a[1] ?? 0), 'up'),
  ROUNDDOWN: (a) => roundTo(toNum(a[0]), toNum(a[1] ?? 0), 'down'),
  MROUND: (a) => { const x = toNum(a[0]), m = toNum(a[1]); if (m === 0) return 0; if (x * m < 0) throw E.num; return Math.round(x / m) * m; },
  ABS: (a) => Math.abs(toNum(a[0])),
  SIGN: (a) => Math.sign(toNum(a[0])),
  INT: (a) => Math.floor(toNum(a[0])),
  MOD: (a) => { const x = toNum(a[0]), d = toNum(a[1]); if (d === 0) throw E.div0; return x - d * Math.floor(x / d); },
  POWER: (a) => Math.pow(toNum(a[0]), toNum(a[1])),
  SQRT: (a) => { const x = toNum(a[0]); if (x < 0) throw E.num; return Math.sqrt(x); },
  PI: () => Math.PI,
  RAND: () => Math.random(),
  CONCAT: (a) => { let s = ''; flat(a, (v) => (s += toStr(v))); return s; },
  CONCATENATE: (a) => a.map(toStr).join(''),
  LEN: (a) => toStr(a[0]).length,
  UPPER: (a) => toStr(a[0]).toUpperCase(),
  LOWER: (a) => toStr(a[0]).toLowerCase(),
  TRIM: (a) => toStr(a[0]).trim().replace(/ +/g, ' '),
  LEFT: (a) => toStr(a[0]).slice(0, a.length > 1 ? toNum(a[1]) : 1),
  RIGHT: (a) => { const s = toStr(a[0]), n = a.length > 1 ? toNum(a[1]) : 1; return n ? s.slice(-n) : ''; },
  MID: (a) => toStr(a[0]).substr(toNum(a[1]) - 1, toNum(a[2])),
  REPLACE: (a) => { const s = toStr(a[0]), st = toNum(a[1]) - 1, n = toNum(a[2]); return s.slice(0, st) + toStr(a[3]) + s.slice(st + n); },
  SUBSTITUTE: (a) => { const s = toStr(a[0]), f = toStr(a[1]), t = toStr(a[2]); if (!f) return s; if (a.length > 3) { const k = toNum(a[3]); let i = -1; for (let j = 0; j < k; j++) { i = s.indexOf(f, i + 1); if (i < 0) return s; } return s.slice(0, i) + t + s.slice(i + f.length); } return s.split(f).join(t); },
  FIND: (a) => { const i = toStr(a[1]).indexOf(toStr(a[0]), (a.length > 2 ? toNum(a[2]) : 1) - 1); if (i < 0) throw E.value; return i + 1; },
  SEARCH: (a) => { const i = toStr(a[1]).toLowerCase().indexOf(toStr(a[0]).toLowerCase(), (a.length > 2 ? toNum(a[2]) : 1) - 1); if (i < 0) throw E.value; return i + 1; },
  VALUE: (a) => toNum(a[0]),
  TEXT: (a) => formatValue(scalar(a[0]), toStr(a[1]), '.').text,
  TODAY: () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); },
  NOW: () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
  DATE: (a) => { const y = toNum(a[0]), m = toNum(a[1]), d = toNum(a[2]); const dt = new Date(Date.UTC(y < 1900 ? y + 1900 : y, m - 1, d)); return dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); },
  YEAR: (a) => serialToDate(toNum(a[0])).y,
  MONTH: (a) => serialToDate(toNum(a[0])).m,
  DAY: (a) => serialToDate(toNum(a[0])).d,
  HOUR: (a) => Math.floor(serialToDate(toNum(a[0])).secs / 3600),
  MINUTE: (a) => Math.floor(serialToDate(toNum(a[0])).secs / 60) % 60,
  SECOND: (a) => serialToDate(toNum(a[0])).secs % 60,
  WEEKDAY: (a) => { const w = serialToDate(toNum(a[0])).wd; const t = a.length > 1 ? toNum(a[1]) : 1; return t === 2 ? ((w + 6) % 7) + 1 : t === 3 ? (w + 6) % 7 : w + 1; },
  DAYS: (a) => Math.floor(toNum(a[0])) - Math.floor(toNum(a[1])),
  ISBLANK: (a) => scalar(a[0]) == null,
  ISNUMBER: (a) => typeof scalar(a[0]) === 'number',
  ISTEXT: (a) => typeof scalar(a[0]) === 'string',
  ISERROR: (a) => isErr(scalar(a[0])),
  CHOOSE: (a) => { const i = toNum(a[0]); if (i < 1 || i >= a.length) throw E.value; return a[Math.floor(i)]; },
  SUMIF: (a) => {
    const rng = asMatrix(a[0]), f = criteria(a[1]), sum = a.length > 2 ? asMatrix(a[2]) : rng;
    let s = 0;
    rng.rows.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = sum.rows[i] && sum.rows[i][j]; if (typeof x === 'number') s += x; } }));
    return s;
  },
  COUNTIF: (a) => { const f = criteria(a[1]); let k = 0; for (const v of asMatrix(a[0]).values()) if (f(v)) k++; return k; },
  AVERAGEIF: (a) => {
    const rng = asMatrix(a[0]), f = criteria(a[1]), avg = a.length > 2 ? asMatrix(a[2]) : rng;
    let s = 0, k = 0;
    rng.rows.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = avg.rows[i] && avg.rows[i][j]; if (typeof x === 'number') { s += x; k++; } } }));
    if (!k) throw E.div0;
    return s / k;
  },
  SUMIFS: (a) => ifs(a.slice(1), asMatrix(a[0]), (xs) => xs.reduce((s, x) => s + x, 0)),
  COUNTIFS: (a) => ifs(a, null, (xs) => xs.length),
  MAXIFS: (a) => ifs(a.slice(1), asMatrix(a[0]), (xs) => (xs.length ? Math.max(...xs) : 0)),
  MINIFS: (a) => ifs(a.slice(1), asMatrix(a[0]), (xs) => (xs.length ? Math.min(...xs) : 0)),
  SUMPRODUCT: (a) => {
    const ms = a.map(asMatrix);
    let s = 0;
    if (ms.some((m) => m.h !== ms[0].h || m.w !== ms[0].w)) throw E.value;
    ms[0].rows.forEach((row, i) => row.forEach((_, j) => { let p = 1; for (const m of ms) { const v = m.rows[i] && m.rows[i][j]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : 0; } s += p; }));
    return s;
  },
  VLOOKUP: (a) => {
    const key = scalar(a[0]), t = asMatrix(a[1]), col = toNum(a[2]), approx = a.length > 3 ? toBool(a[3]) : true;
    if (col < 1 || col > t.w) throw E.ref;
    const i = lookupIdx(t.rows.map((r) => r[0]), key, approx);
    if (i < 0) throw E.na;
    return t.rows[i][col - 1];
  },
  HLOOKUP: (a) => {
    const key = scalar(a[0]), t = asMatrix(a[1]), row = toNum(a[2]), approx = a.length > 3 ? toBool(a[3]) : true;
    if (row < 1 || row > t.h) throw E.ref;
    const i = lookupIdx(t.rows[0], key, approx);
    if (i < 0) throw E.na;
    return t.rows[row - 1][i];
  },
  INDEX: (a) => {
    const t = asMatrix(a[0]);
    let r = a.length > 1 ? toNum(a[1]) : 1, c = a.length > 2 ? toNum(a[2]) : 1;
    if (t.h === 1 && a.length === 2) { c = r; r = 1; }
    // 0 — вся строка или весь столбец
    if (r === 0 && c === 0) return t;
    if (r === 0) { if (c < 1 || c > t.w) throw E.ref; return new Matrix(t.rows.map((row) => [row[c - 1]])); }
    if (c === 0) { if (r < 1 || r > t.h) throw E.ref; return new Matrix([t.rows[r - 1].slice()]); }
    if (r < 1 || c < 1 || r > t.h || c > t.w) throw E.ref;
    return t.rows[r - 1][c - 1];
  },
  MATCH: (a) => {
    const key = scalar(a[0]), t = asMatrix(a[1]), type = a.length > 2 ? toNum(a[2]) : 1;
    const list = t.h === 1 ? t.rows[0] : t.rows.map((r) => r[0]);
    let i;
    if (type === 0) i = lookupIdx(list, key, false);
    else if (type > 0) i = lookupIdx(list, key, true);
    else { i = -1; for (let k = 0; k < list.length; k++) if (list[k] != null && cmp(list[k], key) >= 0) i = k; else break; }
    if (i < 0) throw E.na;
    return i + 1;
  },
  NPV: (a) => { const r = toNum(a[0]); let s = 0, k = 1; for (const x of nums(a.slice(1))) s += x / Math.pow(1 + r, k++); return s; },
  IRR: (a) => {
    const cf = nums([a[0]]);
    let r = a.length > 1 ? toNum(a[1]) : 0.1;
    for (let it = 0; it < 100; it++) {
      let f = 0, d = 0;
      cf.forEach((x, k) => { f += x / Math.pow(1 + r, k); d -= (k * x) / Math.pow(1 + r, k + 1); });
      if (Math.abs(d) < 1e-14) break;
      const nr = r - f / d;
      if (!isFinite(nr)) throw E.num;
      if (Math.abs(nr - r) < 1e-12) return nr;
      r = nr;
    }
    throw E.num;
  },
  PMT: (a) => { const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0, t = a.length > 4 ? toNum(a[4]) : 0; if (r === 0) return -(pv + fv) / n; const q = Math.pow(1 + r, n); return -(r * (pv * q + fv)) / ((1 + r * t) * (q - 1)); },
  FV: (a) => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), pv = a.length > 3 ? toNum(a[3]) : 0, t = a.length > 4 ? toNum(a[4]) : 0; if (r === 0) return -(pv + pmt * n); const q = Math.pow(1 + r, n); return -(pv * q + pmt * (1 + r * t) * (q - 1) / r); },
  PV: (a) => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0, t = a.length > 4 ? toNum(a[4]) : 0; if (r === 0) return -(fv + pmt * n); const q = Math.pow(1 + r, n); return -(fv + pmt * (1 + r * t) * (q - 1) / r) / q; },
  RANK: (a) => { const x = toNum(a[0]), n = nums([a[1]]), desc = !(a.length > 2 && toNum(a[2])); if (!n.includes(x)) throw E.na; return 1 + n.filter((v) => (desc ? v > x : v < x)).length; },
};
FN.AVERAGEA = FN.AVERAGE;

function ifs(pairs, target, agg) {
  const conds = [];
  for (let i = 0; i + 1 < pairs.length; i += 2) conds.push([asMatrix(pairs[i]), criteria(pairs[i + 1])]);
  const base = target || conds[0][0];
  const xs = [];
  base.rows.forEach((row, i) => row.forEach((v, j) => {
    if (conds.every(([m, f]) => f(m.rows[i] ? m.rows[i][j] : null))) {
      if (!target) xs.push(1);
      else if (typeof v === 'number') xs.push(v);
    }
  }));
  return agg(xs);
}

const SUPPORTED = new Set([...Object.keys(FN), 'IF', 'IFERROR', 'IFS', 'IFNA', 'ROW', 'COLUMN']);

// Поддерживается ли формула целиком (иначе оставляем значение, сохранённое в файле).
function formulaSupported(ast) {
  let ok = true;
  (function walk(n) {
    if (!n || !ok) return;
    if (n.t === 'fn') {
      if (!SUPPORTED.has(n.name.replace(/^_xl(fn|ws)\./, '').toUpperCase())) ok = false;
      n.args.forEach(walk);
    } else if (n.t === 'name') ok = false;
    else if (n.a) { walk(n.a); if (n.t === 'bin') walk(n.b); }
  })(ast);
  return ok;
}

function evalAst(n, ctx) {
  switch (n.t) {
    case 'num': return n.v;
    case 'str': return n.v;
    case 'bool': return n.v;
    case 'err': return errOf(n.v);
    case 'missing': return null;
    case 'paren': return evalAst(n.a, ctx);
    case 'name': throw E.name;
    case 'ref': {
      if (n.bad) throw E.ref;
      const sh = n.sheet || ctx.sheet;
      if (n.a.r == null || n.a.c == null) throw E.value;
      const v = ctx.cell(sh, n.a.r, n.a.c);
      if (v === undefined) throw E.ref;
      return v;
    }
    case 'range': {
      if (n.bad) throw E.ref;
      const sh = n.sheet || ctx.sheet;
      const dim = ctx.dims(sh);
      if (!dim) throw E.ref;
      let r1 = n.a.r, r2 = n.b.r, c1 = n.a.c, c2 = n.b.c;
      if (r1 == null) { r1 = 0; r2 = dim.r; }
      if (c1 == null) { c1 = 0; c2 = dim.c; }
      if (r1 > r2) [r1, r2] = [r2, r1];
      if (c1 > c2) [c1, c2] = [c2, c1];
      r2 = Math.min(r2, Math.max(dim.r, r1));
      c2 = Math.min(c2, Math.max(dim.c, c1));
      const rows = [];
      for (let r = r1; r <= r2; r++) {
        const row = [];
        for (let c = c1; c <= c2; c++) row.push(ctx.cell(sh, r, c));
        rows.push(row);
      }
      return new Matrix(rows);
    }
    case 'un': {
      const x = evalAst(n.a, ctx);
      if (x instanceof Matrix) return mapM(x, null, (v) => { const k = toNum(v); return n.op === '-' ? -k : k; });
      const v = toNum(x);
      return n.op === '-' ? -v : v;
    }
    case 'pct': { const x = evalAst(n.a, ctx); return x instanceof Matrix ? mapM(x, null, (v) => toNum(v) / 100) : toNum(x) / 100; }
    case 'bin': {
      const A = evalAst(n.a, ctx), B = evalAst(n.b, ctx);
      // Действие над диапазонами — поэлементно (как в Excel 365 и внутри СУММПРОИЗВ)
      if (A instanceof Matrix || B instanceof Matrix) return mapM(A, B, (a, b) => binOp(n.op, a, b));
      return binOp(n.op, scalar(A), scalar(B));
    }
    case 'fn': {
      const name = n.name.replace(/^_xl(fn|ws)\./, '').toUpperCase();
      const arg = (i) => (n.args[i] ? evalAst(n.args[i], ctx) : null);
      if (name === 'IF') {
        const cond = arg(0);
        if (cond instanceof Matrix) {
          const yes = n.args.length > 1 ? arg(1) : true, no = n.args.length > 2 ? arg(2) : false;
          return mapM(cond, null, (v) => v).rows.length && new Matrix(cond.rows.map((row, i) => row.map((v, j) => {
            const pick = (x) => (x instanceof Matrix ? (x.rows[x.h === 1 ? 0 : i] || [])[x.w === 1 ? 0 : j] : x);
            try { return toBool(v) ? pick(yes) : pick(no); } catch (e) { return e; }
          })));
        }
        const c = toBool(scalar(cond));
        if (c) return n.args.length > 1 ? (n.args[1].t === 'missing' ? 0 : arg(1)) : true;
        return n.args.length > 2 ? (n.args[2].t === 'missing' ? 0 : arg(2)) : false;
      }
      if (name === 'IFS') {
        for (let i = 0; i + 1 < n.args.length; i += 2) if (toBool(scalar(arg(i)))) return arg(i + 1);
        throw E.na;
      }
      if (name === 'IFERROR' || name === 'IFNA') {
        try {
          const v = scalar(arg(0));
          if (isErr(v) && (name === 'IFERROR' || v.err === '#N/A')) return arg(1);
          return v;
        } catch (e) {
          if (isErr(e) && (name === 'IFERROR' || e.err === '#N/A')) return arg(1);
          throw e;
        }
      }
      if (name === 'ROW' || name === 'COLUMN') {
        const a = n.args[0];
        if (!a) return name === 'ROW' ? ctx.r + 1 : ctx.c + 1;
        if (a.t !== 'ref' && a.t !== 'range') throw E.value;
        return name === 'ROW' ? a.a.r + 1 : a.a.c + 1;
      }
      const f = FN[name];
      if (!f) throw E.name;
      const args = n.args.map((x) => evalAst(x, ctx));
      const v = f(args);
      if (typeof v === 'number' && !isFinite(v)) throw E.num;
      return v;
    }
  }
  throw E.value;
}

function binOp(op, a, b) {
  if (isErr(a)) throw a;
  if (isErr(b)) throw b;
  switch (op) {
    case '+': return toNum(a) + toNum(b);
    case '-': return toNum(a) - toNum(b);
    case '*': return toNum(a) * toNum(b);
    case '/': { const d = toNum(b); if (d === 0) throw E.div0; return toNum(a) / d; }
    case '^': return Math.pow(toNum(a), toNum(b));
    case '&': return toStr(a) + toStr(b);
    case '=': return cmp(a, b) === 0;
    case '<>': return cmp(a, b) !== 0;
    case '<': return cmp(a, b) < 0;
    case '>': return cmp(a, b) > 0;
    case '<=': return cmp(a, b) <= 0;
    case '>=': return cmp(a, b) >= 0;
  }
  throw E.value;
}

// Поэлементно; одиночное значение растягивается на весь диапазон, ошибки остаются в своих клетках.
function mapM(A, B, fn) {
  const a = A instanceof Matrix ? A : null, b = B instanceof Matrix ? B : null;
  const h = Math.max(a ? a.h : 1, b ? b.h : 1), w = Math.max(a ? a.w : 1, b ? b.w : 1);
  const at = (m, s, i, j) => (m ? (m.rows[m.h === 1 ? 0 : i] || [])[m.w === 1 ? 0 : j] : s);
  const rows = [];
  for (let i = 0; i < h; i++) {
    const row = [];
    for (let j = 0; j < w; j++) {
      try { row.push(fn(at(a, A, i, j), at(b, B, i, j))); } catch (e) { if (isErr(e)) row.push(e); else throw e; }
    }
    rows.push(row);
  }
  return new Matrix(rows);
}

function evaluate(ast, ctx) {
  try {
    let v = scalar(evalAst(ast, ctx));
    if (v == null) v = 0;
    if (typeof v === 'number') v = +v.toPrecision(15);
    return v;
  } catch (e) {
    if (isErr(e)) return e;
    console.error(e);
    return E.value;
  }
}

// Текст, набранный пользователем: русские имена и «;» → как в файле.
function normalizeInput(text) {
  // Русская запись: «;» между аргументами и запятая в числах (1,5) → 1.5 и «,».
  let out = '', q = false;
  const ru = /;/.test(text.replace(/"[^"]*"/g, ''));
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') q = !q;
    if (!q && ru && ch === ',' && /\d/.test(text[i - 1] || '') && /\d/.test(text[i + 1] || '')) { out += '.'; continue; }
    out += ch;
  }
  const ast = parseFormula(out);
  // Пустые скобки там, где нужны аргументы, Excel не примет: СУММ() — ошибка
  const NOARGS = new Set(['TODAY', 'NOW', 'PI', 'RAND', 'ROW', 'COLUMN', 'NA', 'TRUE', 'FALSE']);
  (function walk(n) {
    if (!n) return;
    if (n.t === 'fn') {
      if (!n.args.length && !NOARGS.has(n.name.replace(/^_xl(fn|ws)\./, '').toUpperCase())) throw new Error('args');
      n.args.forEach(walk);
    } else if (n.a) { walk(n.a); if (n.t === 'bin') walk(n.b); }
  })(ast);
  return astStr(ast);
}
