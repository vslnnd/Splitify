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
  openFile: (filePath) => ipcRenderer.invoke('open-file', filePath),
  showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),

  // Auto-updater
  installUpdate: () => ipcRenderer.invoke('install-update'),
  onUpdateAvailable: (cb) => ipcRenderer.on('update-available', (_, version) => cb(version)),
  onUpdateProgress: (cb) => ipcRenderer.on('update-progress', (_, progress) => cb(progress)),
  onUpdateDownloaded: (cb) => ipcRenderer.on('update-downloaded', () => cb()),
  onUpdateChecking: (cb) => ipcRenderer.on('update-checking', () => cb()),
  onUpdateNotAvailable: (cb) => ipcRenderer.on('update-not-available', (_, v) => cb(v)),
  onUpdateError: (cb) => ipcRenderer.on('update-error', (_, msg) => cb(msg)),
  onAppVersion: (cb) => ipcRenderer.on('app-version', (_, v) => cb(v)),

  // Settings
  getSettings: () => ipcRenderer.invoke('get-settings'),
  saveSettings: (s) => ipcRenderer.invoke('save-settings', s),
  checkForUpdates: () => ipcRenderer.invoke('check-for-updates'),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),

  // History
  getHistory: () => ipcRenderer.invoke('get-history'),
  addHistoryEntry: (entry) => ipcRenderer.invoke('add-history-entry', entry),
  clearHistory: () => ipcRenderer.invoke('clear-history'),
});
