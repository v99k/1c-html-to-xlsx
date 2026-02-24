const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  chooseHtmlFiles: () => ipcRenderer.invoke('choose-html-files'),
  chooseSaveDir: () => ipcRenderer.invoke('choose-save-dir'),
  convert: (paths) => ipcRenderer.invoke('convert', paths),
  saveResults: (results, dir) => ipcRenderer.invoke('save-results', results, dir),
  onLog: (cb) => { ipcRenderer.on('log', (_, line) => cb(line)); },
  onProgress: (cb) => { ipcRenderer.on('progress', (_, data) => cb(data)); },
});
