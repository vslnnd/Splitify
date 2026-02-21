const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('splitify', {
  // Files
  getFilePath:   (file)          => webUtils.getPathForFile(file),
  pickFiles:     ()              => ipcRenderer.invoke('pick-files'),
  pickFolder:    ()              => ipcRenderer.invoke('pick-folder'),
  readFileColumns: (p)           => ipcRenderer.invoke('read-file-columns', p),
  splitFile:     (opts)          => ipcRenderer.invoke('split-file', opts),
  openFolder:    (p)             => ipcRenderer.invoke('open-folder', p),
  openExternal:  (url)           => ipcRenderer.invoke('open-external', url),

  // Profiles
  getProfiles:   ()              => ipcRenderer.invoke('get-profiles'),
  saveProfiles:  (profiles)      => ipcRenderer.invoke('save-profiles', profiles),

  // Settings
  getSettings:   ()              => ipcRenderer.invoke('get-settings'),
  saveSettings:  (s)             => ipcRenderer.invoke('save-settings', s),
  markTutorialShown: ()          => ipcRenderer.invoke('mark-tutorial-shown'),

  // History
  getHistory:       ()           => ipcRenderer.invoke('get-history'),
  addHistoryEntry:  (entry)      => ipcRenderer.invoke('add-history-entry', entry),
  clearHistory:     ()           => ipcRenderer.invoke('clear-history'),

  // Updater
  installUpdate:    ()           => ipcRenderer.invoke('install-update'),
  checkForUpdates:  ()           => ipcRenderer.invoke('check-for-updates'),

  // Events
  onAppVersion:           (cb) => ipcRenderer.on('app-version',              (_, v)    => cb(v)),
  onUpdateChecking:       (cb) => ipcRenderer.on('update-checking',          ()        => cb()),
  onUpdateAvailable:      (cb) => ipcRenderer.on('update-available',         (_, data) => cb(data)),
  onUpdateNotAvailable:   (cb) => ipcRenderer.on('update-not-available',     (_, v)    => cb(v)),
  onUpdateProgress:       (cb) => ipcRenderer.on('update-progress',          (_, pct)  => cb(pct)),
  onUpdateDownloaded:     (cb) => ipcRenderer.on('update-downloaded',        ()        => cb()),
  onUpdateError:          (cb) => ipcRenderer.on('update-error',             (_, msg)  => cb(msg)),
  onFirstLaunchAfterUpdate:(cb) => ipcRenderer.on('first-launch-after-update',(_, v)   => cb(v)),
  onShowTutorial:         (cb) => ipcRenderer.on('show-tutorial',            ()        => cb()),
});
