const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('folio', {
  openDialog: (opts) => ipcRenderer.invoke('open-dialog', opts),
  saveDialog: (defaultPath, filters) => ipcRenderer.invoke('save-dialog', defaultPath, filters),
  htmlToPdf: (html) => ipcRenderer.invoke('html-to-pdf', html),
  readFile: (p) => ipcRenderer.invoke('read-file', p),
  writeFile: (p, data) => ipcRenderer.invoke('write-file', p, data),
  font: (name) => ipcRenderer.invoke('font', name),
  ask: (opts) => ipcRenderer.invoke('ask', opts),
  showInFolder: (p) => ipcRenderer.invoke('show-in-folder', p),
  pathOf: (file) => (webUtils && webUtils.getPathForFile ? webUtils.getPathForFile(file) : file.path),
  setTheme: (colors) => ipcRenderer.send('theme', colors),
  setTitle: (t) => ipcRenderer.send('title', t),
  prefs: ipcRenderer.sendSync('get-prefs'),
  setPref: (key, value) => ipcRenderer.send('set-pref', key, value),
  confirmClose: () => ipcRenderer.send('close-confirmed'),
  setFullscreen: (v) => ipcRenderer.send('fullscreen', v),
  onFullscreen: (cb) => ipcRenderer.on('fullscreen', (_e, v) => cb(v)),
  onOpenFiles: (cb) => ipcRenderer.on('open-files', (_e, files) => cb(files)),
  onCloseRequest: (cb) => ipcRenderer.on('close-request', () => cb()),
});
