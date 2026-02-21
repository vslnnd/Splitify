const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('splitify', {
  // Core
  getProfiles: () => ipcRenderer.invoke('get-profiles'),
  saveProfiles: (profiles) => ipcRenderer.invoke('save-profiles', profiles),
  pickFiles: () => ipcRenderer.invoke('pick-files'),
  pickFolder: () => ipcRenderer.invoke('pick-folder'),
  getFilePath: (file) => webUtils.getPathForFile(file),
  readFileColumns: (filePath) => ipcRenderer.invoke('read-file-columns', filePath),
  splitFile: (opts) => ipcRenderer.invoke('split-file', opts),
  openFolder: (folderPath) => ipcRenderer.invoke('open-folder', folderPath),

  // Auto-updater
  installUpdate: () => ipcRenderer.invoke('install-update'),
  onUpdateAvailable: (cb) => ipcRenderer.on('update-available', (_, version) => cb(version)),
  onUpdateProgress: (cb) => ipcRenderer.on('update-progress', (_, percent) => cb(percent)),
  onUpdateDownloaded: (cb) => ipcRenderer.on('update-downloaded', () => cb())
});
