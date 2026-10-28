// Folio — главный процесс: окно, локальный протокол app://, работа с файлами.
const { app, BrowserWindow, ipcMain, dialog, protocol, net, Menu, nativeTheme, shell } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const fss = require('fs');
const { pathToFileURL } = require('url');

const APP_DIR = path.join(__dirname, 'app');
const DOC_RE = /\.(pdf|docx|xlsx|pptx|png|jpe?g|webp|gif|bmp)$/i;

protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } },
]);

let win = null;

// Настройки (тема и язык) — в файле, чтобы тема применялась ещё до появления окна.
const prefsFile = () => path.join(app.getPath('userData'), 'prefs.json');
let prefs = { theme: 'system', lang: 'ru' };
function loadPrefs() {
  try { prefs = { ...prefs, ...JSON.parse(fss.readFileSync(prefsFile(), 'utf8')) }; } catch {}
  nativeTheme.themeSource = ['light', 'dark'].includes(prefs.theme) ? prefs.theme : 'system';
}
const FILTER_NAMES = {
  ru: { docs: 'Документы', all: 'Все файлы' },
  en: { docs: 'Documents', all: 'All files' },
  kk: { docs: 'Құжаттар', all: 'Барлық файлдар' },
};
let allowClose = false;
const queued = [];

function docsFromArgv(argv) {
  return argv.filter((a) => DOC_RE.test(a) && fss.existsSync(a)).map((a) => path.resolve(a));
}

function sendFiles(files) {
  if (!files.length) return;
  if (win && win.webContents && !win.webContents.isLoading()) win.webContents.send('open-files', files);
  else queued.push(...files);
}

// Версия из исходников живёт отдельно от установленной и не мешает ей.
if (!app.isPackaged) app.setPath('userData', app.getPath('userData') + '-dev');

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  queued.push(...docsFromArgv(process.argv));

  app.on('second-instance', (_e, argv) => {
    sendFiles(docsFromArgv(argv));
    if (win) {
      if (win.isMinimized()) win.restore();
      win.focus();
    }
  });

  app.whenReady().then(() => {
    loadPrefs();
    protocol.handle('app', (req) => {
      const rel = decodeURIComponent(new URL(req.url).pathname).replace(/^\/+/, '');
      const file = path.normalize(path.join(APP_DIR, rel));
      if (!file.startsWith(APP_DIR)) return new Response('Forbidden', { status: 403 });
      return net.fetch(pathToFileURL(file).toString());
    });
    Menu.setApplicationMenu(null);
    createWindow();
  });

  app.on('window-all-closed', () => app.quit());
}

function createWindow() {
  const dark = nativeTheme.shouldUseDarkColors;
  win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 720,
    minHeight: 480,
    show: true,
    title: 'Folio',
    icon: path.join(__dirname, 'build', 'icon.png'),
    backgroundColor: dark ? '#131416' : '#dfe1e4',
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: dark ? '#131416' : '#dfe1e4', symbolColor: dark ? '#e6e7e9' : '#1d2125', height: 40 },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  win.loadURL('app://folio/index.html');

  win.webContents.on('did-finish-load', () => {
    if (queued.length) win.webContents.send('open-files', queued.splice(0));
  });

  // Ссылки из PDF — во внешний браузер, не внутри приложения.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:|^mailto:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  win.on('close', (e) => {
    if (allowClose) return;
    e.preventDefault();
    win.webContents.send('close-request');
  });
  win.on('closed', () => (win = null));
  win.on('enter-full-screen', () => win.webContents.send('fullscreen', true));
  win.on('leave-full-screen', () => win.webContents.send('fullscreen', false));
}

// ---------- IPC ----------
ipcMain.handle('open-dialog', async (_e, opts = {}) => {
  const r = await dialog.showOpenDialog(win, {
    properties: ['openFile', 'multiSelections'],
    filters: opts.filters || [
      { name: (FILTER_NAMES[prefs.lang] || FILTER_NAMES.ru).docs, extensions: ['pdf', 'docx', 'xlsx', 'pptx', 'png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'] },
      { name: (FILTER_NAMES[prefs.lang] || FILTER_NAMES.ru).all, extensions: ['*'] },
    ],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle('save-dialog', async (_e, defaultPath, filters) => {
  const r = await dialog.showSaveDialog(win, {
    defaultPath,
    filters: filters || [{ name: 'PDF', extensions: ['pdf'] }],
  });
  return r.canceled ? null : r.filePath;
});

// Печать HTML-документа в PDF через невидимое окно (Word → PDF).
ipcMain.handle('html-to-pdf', async (_e, html) => {
  const tmp = path.join(app.getPath('temp'), `folio-print-${Date.now()}.html`);
  await fs.writeFile(tmp, html, 'utf8');
  const w = new BrowserWindow({ show: false, webPreferences: { javascript: false } });
  try {
    await w.loadFile(tmp);
    const pdf = await w.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true });
    return new Uint8Array(pdf);
  } finally {
    w.destroy();
    fs.unlink(tmp).catch(() => {});
  }
});

ipcMain.handle('read-file', async (_e, p) => {
  const buf = await fs.readFile(p);
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
});

ipcMain.handle('write-file', async (_e, p, data) => {
  // Пишем во временный файл и подменяем — если что-то пойдёт не так, оригинал цел.
  const tmp = p + '.folio-tmp';
  await fs.writeFile(tmp, Buffer.from(data));
  await fs.rename(tmp, p);
  return true;
});

ipcMain.handle('font', async (_e, name) => {
  if (!/^[a-z]+$/i.test(name)) throw new Error('bad font');
  const dir = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
  const buf = await fs.readFile(path.join(dir, name + '.ttf'));
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
});

ipcMain.handle('ask', async (_e, { message, detail, buttons, cancelId }) => {
  const r = await dialog.showMessageBox(win, {
    type: 'none',
    message,
    detail,
    buttons,
    cancelId,
    defaultId: 0,
    noLink: true,
  });
  return r.response;
});

ipcMain.on('close-confirmed', () => {
  allowClose = true;
  if (win) win.close();
});

ipcMain.on('theme', (_e, { bg, fg }) => {
  if (win) {
    win.setTitleBarOverlay({ color: bg, symbolColor: fg, height: 40 });
    win.setBackgroundColor(bg);
  }
});

ipcMain.on('fullscreen', (_e, v) => win && win.setFullScreen(v == null ? !win.isFullScreen() : !!v));

ipcMain.on('get-prefs', (e) => { e.returnValue = prefs; });
ipcMain.on('set-pref', (_e, key, value) => {
  prefs[key] = value;
  if (key === 'theme') nativeTheme.themeSource = ['light', 'dark'].includes(value) ? value : 'system';
  try { fss.writeFileSync(prefsFile(), JSON.stringify(prefs)); } catch {}
});

ipcMain.on('title', (_e, t) => win && win.setTitle(t));

ipcMain.handle('show-in-folder', (_e, p) => shell.showItemInFolder(p));
