// Folio — оболочка: вкладки, две половины окна, открытие файлов, горячие клавиши.
const $ = (s) => document.querySelector(s);
const tabsEl = $('#tabs');
const panesEl = $('#panes');
const emptyEl = $('#empty');
const saveBtn = $('#btn-save');
const splitBtn = $('#btn-split');
const showBtn = $('#btn-show');
const tabs = [];
let active = null;
let tabSeq = 1;
let mru = []; // вкладки в порядке последнего использования

// Две половины окна: в каждой — своя вкладка. focus — половина, с которой сейчас работают.
const panes = [
  { el: $('#pane0'), tab: null },
  { el: $('#pane1'), tab: null },
];
let split = false;
let focus = 0;
let full = false;

$('#btn-open').innerHTML = ICONS.plus;
saveBtn.innerHTML = ICONS.save;
splitBtn.innerHTML = ICONS.split;
showBtn.innerHTML = ICONS.play;
$('#btn-menu').innerHTML = ICONS.more;
document.querySelectorAll('.drop-icon').forEach((n) => (n.innerHTML = ICONS.open));
$('.nd-icon').innerHTML = ICONS.newdoc;

const baseName = (p) => p.split(/[\\/]/).pop();
const dirName = (p) => p.split(/[\\/]/).slice(0, -1).join('\\');
const isTyping = (t) => t && t.closest && t.closest('input, textarea, [contenteditable="true"], [contenteditable="plaintext-only"]');

// ---------- Сообщения и меню ----------
function toast(text) {
  const t = $('#toast');
  t.textContent = text;
  t.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove('show'), 2200);
}

function showMenu(x, y, items) {
  document.querySelector('.menu')?.remove();
  const m = document.createElement('div');
  m.className = 'menu';
  for (const it of items) {
    if (!it) { m.append(document.createElement('hr')); continue; }
    const b = document.createElement('button');
    if (it[3] != null) b.classList.add(it[3] ? 'checked' : 'unchecked');
    b.innerHTML = `<span></span>${it[2] ? `<kbd>${it[2]}</kbd>` : ''}`;
    b.firstChild.textContent = it[0];
    b.addEventListener('click', () => { m.remove(); it[1](); });
    m.append(b);
  }
  document.body.append(m);
  const r = m.getBoundingClientRect();
  m.style.left = Math.max(8, Math.min(x, innerWidth - r.width - 8)) + 'px';
  m.style.top = Math.min(y, innerHeight - r.height - 8) + 'px';
  const off = (e) => {
    if (m.contains(e.target)) return;
    m.remove();
    window.removeEventListener('pointerdown', off, true);
  };
  setTimeout(() => window.addEventListener('pointerdown', off, true));
}

// Меню над кнопкой (для кнопок у нижнего края).
function menuAbove(btn, items) {
  const r = btn.getBoundingClientRect();
  showMenu(r.left, r.top - 8 - items.filter(Boolean).length * 32 - 10, items);
}

function menuUnder(btn, items) {
  const r = btn.getBoundingClientRect();
  showMenu(r.left, r.bottom + 6, items);
}

// Подсказки с задержкой — чтобы интерфейс оставался чистым.
(() => {
  const tip = $('#tip');
  let timer, cur;
  document.addEventListener('pointerover', (e) => {
    const t = e.target.closest('[data-tip]');
    if (t === cur) return;
    cur = t;
    clearTimeout(timer);
    tip.classList.remove('show');
    if (!t) return;
    timer = setTimeout(() => {
      tip.textContent = t.dataset.tip;
      const r = t.getBoundingClientRect();
      const tr = tip.getBoundingClientRect();
      const below = r.top < 60;
      tip.style.left = Math.max(6, Math.min(innerWidth - tr.width - 6, r.left + r.width / 2 - tr.width / 2)) + 'px';
      tip.style.top = (below ? r.bottom + 8 : r.top - tr.height - 8) + 'px';
      tip.classList.add('show');
    }, 500);
  });
  document.addEventListener('pointerdown', () => { clearTimeout(timer); tip.classList.remove('show'); }, true);
})();

// ---------- Недавние файлы ----------
function recentList() {
  try { return JSON.parse(localStorage.getItem('folio.recent') || '[]'); } catch { return []; }
}
function pushRecent(p) {
  if (!p) return;
  const l = recentList().filter((x) => x.toLowerCase() !== p.toLowerCase());
  l.unshift(p);
  try { localStorage.setItem('folio.recent', JSON.stringify(l.slice(0, 7))); } catch {}
}
function dropRecent(p) {
  try { localStorage.setItem('folio.recent', JSON.stringify(recentList().filter((x) => x !== p))); } catch {}
}
function drawRecent() {
  const ul = $('#recent');
  ul.textContent = '';
  for (const p of recentList()) {
    const li = document.createElement('li');
    const b = document.createElement('button');
    b.innerHTML = `${ICONS.file}<span class="r-name"></span><span class="r-dir"></span>`;
    b.querySelector('.r-name').textContent = baseName(p);
    b.querySelector('.r-dir').textContent = dirName(p);
    b.addEventListener('click', () => openPaths([p]));
    li.append(b);
    ul.append(li);
  }
}

// ---------- Половины окна ----------
const shownPanes = () => (split ? [0, 1] : [0]);
const paneOf = (t) => shownPanes().find((i) => panes[i].tab === t);

// Где панель страниц: в двух половинах по умолчанию зеркально (правая — у правого края).
function sideKey(i) { return 'folio.side.' + (split ? 'split' + i : 'single'); }
function sideRight(i) {
  let v = null;
  try { v = localStorage.getItem(sideKey(i)); } catch {}
  return v ? v === 'right' : split && i === 1;
}
window.folioSidePref = (ed, right) => {
  const t = tabs.find((x) => x.ed === ed);
  const i = t ? paneOf(t) : null;
  if (i == null) return;
  try { localStorage.setItem(sideKey(i), right ? 'right' : 'left'); } catch {}
};

function placeEditors() {
  panesEl.classList.toggle('split', split);
  for (const t of tabs) {
    const i = paneOf(t);
    const root = t.ed.root;
    if (i == null) {
      if (!root.hidden) t.ed.endEdit();
      root.hidden = true;
      continue;
    }
    if (root.parentNode !== panes[i].el) {
      // При переносе в другую половину сохраняем место, где читали.
      const top = t.ed.viewer.scrollTop;
      panes[i].el.append(root);
      setTimeout(() => (t.ed.viewer.scrollTop = top), 0);
    }
    root.hidden = false;
    if (t.ed.setSideRight) t.ed.setSideRight(sideRight(i), false);
  }
  $('#pane1-empty').hidden = !(split && !panes[1].tab);
  panes.forEach((p, i) => p.el.classList.toggle('focused', split && i === focus));
}

// Документы в обеих половинах подгоняем по ширине — каждый виден целиком по ширине.
function fitShown() {
  setTimeout(() => {
    for (const i of shownPanes()) {
      const t = panes[i].tab;
      if (t && t.ed.fitWidth) t.ed.fitWidth();
    }
  }, 60);
}

function setFocus(i) {
  focus = i;
  active = panes[i].tab;
  if (active) mru = [active, ...mru.filter((x) => x !== active)];
  placeEditors();
  render();
}

function toggleSplit(force) {
  const on = force != null ? force : !split;
  if (on === split) return;
  if (on) {
    split = true;
    if (!panes[0].tab) panes[0].tab = active;
    panes[1].tab = mru.find((t) => t !== panes[0].tab) || null;
    focus = 0;
  } else {
    const keep = panes[focus].tab || panes[0].tab;
    split = false;
    panes[0].tab = keep;
    panes[1].tab = null;
    focus = 0;
  }
  setFocus(focus);
  fitShown();
}

// Пустая половина: второй документ откроется именно туда.
$('#drop2').addEventListener('click', () => { focus = 1; openDialog(); });

// Разделитель половин можно тянуть; двойной щелчок — поровну.
(() => {
  const div = $('#divider');
  div.addEventListener('pointerdown', (e) => {
    div.setPointerCapture(e.pointerId);
    div.classList.add('dragging');
    const r = panesEl.getBoundingClientRect();
    const move = (ev) => {
      const k = Math.min(0.8, Math.max(0.2, (ev.clientX - r.left) / r.width));
      panes[0].el.style.flex = `${k} 1 0`;
      panes[1].el.style.flex = `${1 - k} 1 0`;
    };
    const up = () => {
      div.classList.remove('dragging');
      div.removeEventListener('pointermove', move);
      div.removeEventListener('pointerup', up);
    };
    div.addEventListener('pointermove', move);
    div.addEventListener('pointerup', up);
  });
  div.addEventListener('dblclick', () => { panes.forEach((p) => (p.el.style.flex = '')); fitShown(); });
})();

// Щелчок в половине делает её активной.
panes.forEach((p, i) => {
  p.el.addEventListener('pointerdown', () => { if (split && focus !== i) setFocus(i); }, true);
});

// ---------- Вкладки ----------
function render() {
  tabsEl.textContent = '';
  for (const t of tabs) {
    const d = document.createElement('div');
    const other = split && paneOf(t) != null && t !== active;
    d.className = 'tab' + (t === active ? ' on' : '') + (other ? ' on2' : '') + (t.ed.dirty ? ' dirty' : '');
    d.innerHTML = `<span class="tab-name"></span><button class="tab-x">${ICONS.close}</button>`;
    d.querySelector('.tab-name').textContent = t.name;
    d.title = t.ed.path || t.name;
    d.addEventListener('pointerdown', (e) => { if (e.button === 0 && !e.target.closest('.tab-x')) activate(t); });
    d.addEventListener('auxclick', (e) => { if (e.button === 1) closeTab(t); });
    d.addEventListener('contextmenu', (e) => { e.preventDefault(); tabMenu(t, e.clientX, e.clientY); });
    d.querySelector('.tab-x').addEventListener('click', () => closeTab(t));
    tabsEl.append(d);
  }
  document.body.classList.toggle('no-docs', !tabs.length);
  const home = tabs.find((t) => t.kind === 'home' && paneOf(t) != null);
  if (home) {
    if (emptyEl.parentNode !== home.ed.root) home.ed.root.append(emptyEl);
    emptyEl.hidden = false;
  } else {
    if (emptyEl.parentNode !== $('#stage')) $('#stage').append(emptyEl);
    emptyEl.hidden = !!tabs.length;
  }
  const isHome = !!active && active.kind === 'home';
  saveBtn.hidden = !(active && active.ed.dirty);
  splitBtn.hidden = !tabs.length;
  showBtn.hidden = !tabs.length || isHome;
  if (!tabs.length || home) drawRecent();
  window.folio.setTitle(active && !isHome ? `${active.name} — Folio` : 'Folio');
}

function tabMenu(t, x, y) {
  if (t.kind === 'home') { showMenu(x, y, [['Закрыть', () => closeTab(t), 'Ctrl+W']]); return; }
  const where = paneOf(t);
  const items = [
    ['Сохранить', () => saveTab(t), 'Ctrl+S'],
    ['Сохранить как…', () => saveTab(t, true), 'Ctrl+Shift+S'],
  ];
  if (t.ed.path) items.push(['Показать в папке', () => window.folio.showInFolder(t.ed.path)]);
  items.push(null);
  if (split) items.push([where === 1 ? 'Перенести в левую половину' : 'Перенести в правую половину', () => moveToPane(t, where === 1 ? 0 : 1)]);
  else if (tabs.length > 1) items.push(['Открыть рядом', () => openBeside(t)]);
  items.push(['Закрыть', () => closeTab(t), 'Ctrl+W']);
  showMenu(x, y, items);
}

// Включаем два окна с выбором: что поставить во вторую половину.
function splitChooser() {
  if (split) { toggleSplit(false); return; }
  if (!active) return;
  const second = (fn) => () => {
    split = true;
    panes[0].tab = active;
    panes[1].tab = null;
    setFocus(1);
    fitShown();
    if (fn) fn();
  };
  const others = [...mru, ...tabs].filter((t, i, list) => t !== active && list.indexOf(t) === i);
  const items = others.map((t) => [t.name, () => openBeside(t), KIND_LABEL[t.kind] || '']);
  if (items.length) items.push(null);
  items.push(['Открыть файл…', second(openDialog), 'Ctrl+O']);
  items.push(['Новый документ Word', second(newWord)]);
  items.push(['Новая таблица Excel', second(newExcel)]);
  items.push(['Новая презентация', second(newSlides)]);
  items.push(['Пустая половина', second(() => setFocus(0))]);
  menuUnder(splitBtn.hidden ? $('#btn-menu') : splitBtn, items);
}

function openBeside(t) {
  split = true;
  panes[0].tab = active && active !== t ? active : mru.find((x) => x !== t) || null;
  panes[1].tab = t;
  setFocus(1);
  fitShown();
}

function moveToPane(t, i) {
  const from = paneOf(t);
  if (from === i) return;
  const there = panes[i].tab;
  panes[i].tab = t;
  if (from != null) panes[from].tab = there || null;
  setFocus(i);
  fitShown();
}

function activate(t) {
  if (!t) { setFocus(focus); return; }
  const i = paneOf(t);
  if (i != null) setFocus(i);
  else { panes[focus].tab = t; setFocus(focus); }
  t.ed.viewer.focus({ preventScroll: true });
}

function newTab(name, path, kind = 'pdf') {
  const t = { id: tabSeq++, name, path, kind };
  const Editor = { docx: DocxEditor, xlsx: typeof XlsxEditor !== 'undefined' && XlsxEditor, pptx: typeof PptxEditor !== 'undefined' && PptxEditor }[kind] || PdfEditor;
  t.ed = new Editor(panes[focus].el, {
    onChange: () => {
      if (t.ed.path && t.ed.path !== t.path) { t.path = t.ed.path; t.name = baseName(t.ed.path); pushRecent(t.path); }
      render();
    },
  });
  t.ed.path = path;
  t.ed.root.hidden = true;
  // Документ, открытый с вкладки «Главная», встаёт на её место — как новая вкладка в браузере.
  const home = active && active.kind === 'home' && paneOf(active) === focus ? active : null;
  if (home) {
    tabs.splice(tabs.indexOf(home), 1, t);
    mru = mru.filter((x) => x !== home);
    panes[focus].tab = t;
    home.ed.destroy();
  } else tabs.push(t);
  activate(t);
  return t;
}

// ---------- Вкладка «Главная» ----------
// «+» открывает начальный экран отдельной вкладкой. Экран один на всё окно —
// он переезжает в ту вкладку «Главная», которая сейчас видна.
function homeView() {
  const root = document.createElement('div');
  root.className = 'doc home';
  root.tabIndex = -1;
  return {
    root, viewer: root, dirty: false, path: null,
    endEdit() {}, onKey: () => false,
    destroy() {
      if (root.contains(emptyEl)) $('#stage').append(emptyEl);
      root.remove();
    },
  };
}
function openHome() {
  const h = tabs.find((t) => t.kind === 'home');
  if (h) { activate(h); return; }
  const t = { id: tabSeq++, name: T('Главная'), path: null, kind: 'home', ed: homeView() };
  t.ed.root.hidden = true;
  panes[focus].el.append(t.ed.root);
  tabs.push(t);
  activate(t);
}

async function closeTab(t) {
  if (t.ed.dirty) {
    activate(t);
    const r = await window.folio.ask({
      message: T(`Сохранить изменения в «${t.name}»?`),
      detail: T('Если не сохранить, правки пропадут.'),
      buttons: [T('Сохранить'), T('Не сохранять'), T('Отмена')],
      cancelId: 2,
    });
    if (r === 2) return false;
    if (r === 0 && !(await saveTab(t))) return false;
  }
  tabs.splice(tabs.indexOf(t), 1);
  mru = mru.filter((x) => x !== t);
  const i = paneOf(t);
  if (i != null) {
    const other = split ? panes[1 - i].tab : null;
    panes[i].tab = mru.find((x) => x !== other) || null;
    if (split && !panes[0].tab && panes[1].tab) {
      panes[0].tab = panes[1].tab;
      panes[1].tab = null;
      if (focus === 1) focus = 0;
    }
  }
  t.ed.destroy();
  setFocus(panes[focus].tab ? focus : 0);
  return true;
}

async function saveTab(t, asNew) {
  if (!t || t.kind === 'home') return false;
  try {
    const ok = await t.ed.save(asNew);
    if (ok) toast(asNew ? 'Сохранено как новый файл' : 'Сохранено');
    return ok;
  } catch (err) {
    console.error(err);
    const locked = /EBUSY|EPERM|EACCES/.test(String(err));
    toast(locked ? 'Файл открыт в другой программе — закройте её или сохраните как новый файл' : 'Не получилось сохранить: ' + (err.message || err));
    return false;
  }
}

// ---------- Открытие и создание ----------
const isPdf = (p) => /\.pdf$/i.test(p);
const isDocx = (p) => /\.docx$/i.test(p);
const isXlsx = (p) => /\.xlsx$/i.test(p);
const isPptx = (p) => /\.pptx$/i.test(p);
const kindOf = (p) => (isDocx(p) ? 'docx' : isXlsx(p) ? 'xlsx' : isPptx(p) ? 'pptx' : 'pdf');
const KIND_LABEL = { pdf: 'PDF', docx: 'Word', xlsx: 'Excel', pptx: 'PowerPoint' };
const isImg = (p) => /\.(png|jpe?g|webp|gif|bmp)$/i.test(p);

async function openPaths(paths) {
  const docs = paths.filter((p) => isPdf(p) || isDocx(p) || isXlsx(p) || isPptx(p));
  const imgs = paths.filter(isImg);
  for (const p of docs) {
    const open = tabs.find((t) => t.path && t.path.toLowerCase() === p.toLowerCase());
    if (open) { activate(open); continue; }
    const t = newTab(baseName(p), p, kindOf(p));
    try {
      const bytes = await window.folio.readFile(p);
      await t.ed.open([{ name: p, bytes }]);
      pushRecent(p);
    } catch (err) {
      console.error(err);
      await closeTab(t);
      dropRecent(p);
      const msg = String(err && (err.name || err.message));
      toast(/Password|encrypt/i.test(msg) ? 'Файл защищён паролем — такие пока не открываются'
        : /ENOENT/.test(String(err)) ? 'Файл не найден — возможно, его переместили'
        : 'Не получилось открыть файл');
      render();
    }
  }
  if (imgs.length) {
    const t = newTab(imgs.length === 1 ? baseName(imgs[0]).replace(/\.[^.]+$/, '.pdf') : T('Картинки.pdf'), null);
    const files = [];
    for (const p of imgs) files.push({ name: p, bytes: await window.folio.readFile(p) });
    await t.ed.open(files);
    t.ed.dirty = true;
    render();
  }
}

async function openDialog() {
  const paths = await window.folio.openDialog();
  if (paths.length) openPaths(paths);
}

let blankSeq = 1;
function newBlank() {
  const t = newTab(T(blankSeq === 1 ? 'Новый документ.pdf' : `Новый документ ${blankSeq}.pdf`), null);
  blankSeq++;
  t.ed.blank();
  t.ed.dirty = true;
  render();
}

let xlSeq = 1;
async function newExcel() {
  const t = newTab(T(xlSeq === 1 ? 'Таблица.xlsx' : `Таблица ${xlSeq}.xlsx`), null, 'xlsx');
  xlSeq++;
  await t.ed.blank();
  t.ed.dirty = true;
  render();
}

let ppSeq = 1;
async function newSlides() {
  const t = newTab(T(ppSeq === 1 ? 'Презентация.pptx' : `Презентация ${ppSeq}.pptx`), null, 'pptx');
  ppSeq++;
  await t.ed.blank();
  t.ed.dirty = true;
  render();
}

// Что можно создать — для кнопки на начальном экране.
const NEW_ITEMS = () => [
  ['Документ Word', newWord, 'Ctrl+N'],
  ['Таблица Excel', newExcel],
  ...(typeof newSlides === 'function' ? [['Презентация PowerPoint', newSlides]] : []),
  ['PDF', newBlank, 'Ctrl+Shift+N'],
];

let wordSeq = 1;
async function newWord() {
  const t = newTab(T(wordSeq === 1 ? 'Документ.docx' : `Документ ${wordSeq}.docx`), null, 'docx');
  wordSeq++;
  await t.ed.blank();
  t.ed.dirty = true;
  render();
}

// ---------- Показ на весь экран ----------
let showing = false;
async function present() {
  if (!active || active.kind === 'home' || showing) return;
  showing = true;
  try {
    const ed = active.ed;
    let bytes, start = 0;
    if (active.kind === 'pdf') {
      ed.endEdit();
      const s0 = ed.sources[0];
      const untouched = !ed.dirty && ed.sources.length === 1 && s0 && ed.pages.length === s0.count &&
        ed.pages.every((p, i) => p.src === s0.id && p.index === i && !p.rot && !p.objs.some((o) => !o.pristine));
      bytes = untouched ? s0.bytes.slice() : await buildPdf(ed);
      start = Math.max(0, ed.pages.findIndex((p) => p.id === ed.cur));
    } else {
      toast('Готовлю показ…');
      bytes = await window.folio.htmlToPdf(await ed.buildPrintHtml());
      if (ed.kind === 'pptx') start = ed.curSlide || 0;
      else if (ed.pageList) {
        const y = (ed.viewer.scrollTop + ed.viewer.clientHeight * 0.4) / ed.zoom;
        ed.pageList.forEach((p, i) => { if (p.top <= y) start = i; });
      }
    }
    await startShow(bytes, start, () => { showing = false; });
  } catch (err) {
    console.error(err);
    toast('Не получилось начать показ');
    showing = false;
  }
}

// ---------- Тема и язык ----------
const THEMES = [['system', 'Как в системе', 'auto'], ['light', 'Светлая', 'sun'], ['dark', 'Тёмная', 'moon']];
const LANGS = [['ru', 'Русский'], ['en', 'English'], ['kk', 'Қазақша']];
let themePref = (window.folio.prefs && window.folio.prefs.theme) || 'system';

function drawPrefs() {
  const th = THEMES.find((x) => x[0] === themePref) || THEMES[0];
  $('#btn-theme').innerHTML = `${ICONS[th[2]]}<span></span>`;
  $('#btn-theme span').textContent = th[1];
  $('#btn-lang').innerHTML = `${ICONS.globe}<span></span>`;
  $('#btn-lang span').textContent = (LANGS.find((x) => x[0] === LANG) || LANGS[0])[1];
}

function setTheme(v) {
  themePref = v;
  window.folio.setPref('theme', v);
  drawPrefs();
  setTimeout(syncTheme, 50);
}

function setLanguage(v) {
  window.folio.setPref('lang', v);
  setLang(v);
  drawPrefs();
  render();
}

function checkItems(list, cur, fn) {
  return list.map(([v, label]) => [label, () => fn(v), '', v === cur]);
}

function themeMenu(btn, above) { (above ? menuAbove : menuUnder)(btn, checkItems(THEMES, themePref, setTheme)); }
function langMenu(btn, above) { (above ? menuAbove : menuUnder)(btn, checkItems(LANGS, LANG, setLanguage)); }

$('#btn-theme').addEventListener('click', (e) => themeMenu(e.currentTarget, true));
$('#btn-lang').addEventListener('click', (e) => langMenu(e.currentTarget, true));

// ---------- Панель инструментов и полный экран ----------
window.setDockHidden = (v) => {
  document.body.classList.toggle('dock-hidden', v);
  try { localStorage.setItem('folio.dockHidden', v ? '1' : ''); } catch {}
};
try { if (localStorage.getItem('folio.dockHidden')) document.body.classList.add('dock-hidden'); } catch {}
const dockHidden = () => document.body.classList.contains('dock-hidden');

function toggleFull(v) { window.folio.setFullscreen(v == null ? !full : v); }
window.folio.onFullscreen((v) => {
  full = v;
  document.body.classList.toggle('full', v);
  if (v) toast('Полный экран · выйти — Esc или F11');
});

function mainMenu() {
  const items = [
    ['Новый документ Word', newWord, 'Ctrl+N'],
    ['Новая таблица Excel', newExcel],
    ['Новая презентация', newSlides],
    ['Новый PDF', newBlank, 'Ctrl+Shift+N'],
    ['Открыть…', openDialog, 'Ctrl+O'],
    null,
  ];
  const doc = active && active.kind !== 'home' ? active : null;
  if (doc) {
    items.push(['Сохранить', () => saveTab(active), 'Ctrl+S']);
    items.push(['Сохранить как…', () => saveTab(active, true), 'Ctrl+Shift+S']);
    if (active.ed.path) items.push(['Показать в папке', () => window.folio.showInFolder(active.ed.path)]);
    items.push(null);
  }
  if (doc) items.push(['Показ на весь экран', present, 'F5']);
  items.push([split ? 'Один документ' : 'Два документа рядом', () => splitChooser(), 'Ctrl+\\']);
  items.push([dockHidden() ? 'Показать панель инструментов' : 'Скрыть панель инструментов', () => window.setDockHidden(!dockHidden()), 'Ctrl+.']);
  items.push([full ? 'Выйти из полного экрана' : 'Во весь экран', () => toggleFull(), 'F11']);
  items.push(null);
  items.push(['Тема', () => themeMenu($('#btn-menu')), '›']);
  items.push(['Язык', () => langMenu($('#btn-menu')), '›']);
  menuUnder($('#btn-menu'), items);
}

$('#btn-open').addEventListener('click', openHome);
$('#btn-menu').addEventListener('click', mainMenu);
$('#drop').addEventListener('click', openDialog);
$('#btn-new').addEventListener('click', (e) => menuUnder(e.currentTarget, NEW_ITEMS()));
saveBtn.addEventListener('click', () => active && saveTab(active));
splitBtn.addEventListener('click', () => splitChooser());
showBtn.addEventListener('click', () => present());
window.folio.onOpenFiles((files) => openPaths(files));

// Перетаскивание файлов из проводника.
(() => {
  const dz = $('#dropzone');
  let depth = 0;
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  window.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    depth++;
    dz.hidden = false;
  });
  window.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; dz.hidden = true; } });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('drop', async (e) => {
    e.preventDefault();
    depth = 0;
    dz.hidden = true;
    const paths = [...e.dataTransfer.files].map((f) => f.path || window.folio.pathOf(f)).filter((p) => isPdf(p) || isDocx(p) || isXlsx(p) || isPptx(p) || isImg(p));
    if (!paths.length) { toast('Открываются PDF, Word, Excel, PowerPoint и картинки'); return; }
    // Файл брошен в правую или левую половину — откроется там.
    if (split) {
      const i = panes.findIndex((p) => p.el.contains(e.target));
      if (i >= 0) focus = i;
    }
    // Брошено на панель страниц — добавляем страницы в этот документ.
    const target = panes[focus].tab;
    if (target && target.kind === 'pdf' && e.target.closest && e.target.closest('.side')) {
      const files = [];
      for (const p of paths) files.push({ name: p, bytes: await window.folio.readFile(p) });
      target.ed.addFiles(files);
      return;
    }
    openPaths(paths);
  });
})();

// ---------- Клавиатура ----------
document.addEventListener('keydown', (e) => {
  if (document.querySelector('.modal-bg')) return;
  const k = e.key.toLowerCase();
  const ctrl = e.ctrlKey || e.metaKey;
  if (e.key === 'F11') { e.preventDefault(); toggleFull(); return; }
  if (e.key === 'F5') { e.preventDefault(); present(); return; }
  if (ctrl && (k === 'o' || k === 'щ')) { e.preventDefault(); openDialog(); return; }
  if (ctrl && (k === 's' || k === 'ы')) { e.preventDefault(); if (active) saveTab(active, e.shiftKey); return; }
  if (ctrl && (k === 'n' || k === 'т')) { e.preventDefault(); if (e.shiftKey) newBlank(); else newWord(); return; }
  if (ctrl && (k === 't' || k === 'е')) { e.preventDefault(); openHome(); return; }
  if (ctrl && (k === 'w' || k === 'ц')) { e.preventDefault(); if (active) closeTab(active); return; }
  if (ctrl && e.code === 'Backslash') { e.preventDefault(); splitChooser(); return; }
  if (ctrl && e.code === 'Period') { e.preventDefault(); window.setDockHidden(!dockHidden()); return; }
  if (ctrl && e.key === 'Tab' && tabs.length > 1) {
    e.preventDefault();
    const i = tabs.indexOf(active);
    activate(tabs[(i + (e.shiftKey ? -1 : 1) + tabs.length) % tabs.length]);
    return;
  }
  if (isTyping(e.target)) return;
  if (active && active.ed.onKey(e)) { e.preventDefault(); return; }
  if (e.key === 'Escape' && full) toggleFull(false);
});

// Копирование объектов: в системный буфер кладём метку, сами объекты — внутри приложения.
document.addEventListener('copy', (e) => {
  if (isTyping(e.target) || !active || !active.ed.copy || active.ed.editing) return;
  if (active.ed.copy()) { e.preventDefault(); e.clipboardData.setData('text/plain', CLIP_MARK); }
});
document.addEventListener('cut', (e) => {
  if (isTyping(e.target) || !active || !active.ed.copy || active.ed.editing) return;
  if (active.ed.copy()) { e.preventDefault(); e.clipboardData.setData('text/plain', CLIP_MARK); active.ed.deleteSel(); }
});
document.addEventListener('paste', (e) => {
  if (isTyping(e.target) || !active || !active.ed.paste || active.ed.editing) return;
  e.preventDefault();
  active.ed.paste(e);
});

// ---------- Закрытие окна ----------
window.folio.onCloseRequest(async () => {
  for (const t of [...tabs]) {
    if (!t.ed.dirty) continue;
    if (!(await closeTab(t))) return;
  }
  window.folio.confirmClose();
});

// ---------- Тема окна ----------
function syncTheme() {
  const cs = getComputedStyle(document.documentElement);
  window.folio.setTheme({ bg: cs.getPropertyValue('--bar').trim(), fg: cs.getPropertyValue('--ink').trim() });
}
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', syncTheme);
syncTheme();
drawPrefs();
placeEditors();
render();
