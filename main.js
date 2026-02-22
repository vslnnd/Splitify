const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const PROFILES_PATH  = path.join(app.getPath('userData'), 'splitify_profiles.json');
const SETTINGS_PATH  = path.join(app.getPath('userData'), 'splitify_settings.json');

const DEFAULT_SETTINGS = {
  theme: 'dark',
  favoriteProfileId: null,
  updateIntervalHours: 4,
  checkUpdatesOnStartup: true
};

function loadSettings() {
  try {
    if (fs.existsSync(SETTINGS_PATH)) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8')) };
    }
  } catch(e) {}
  return { ...DEFAULT_SETTINGS };
}

function saveSettings(settings) {
  try { fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2)); return true; }
  catch(e) { return false; }
}

// ─── Default SElectric Profile ────────────────────────────────────────────────
const DEFAULT_PROFILES = [
  {
    id: 'selectric_' + Date.now(),
    name: 'SElectric',
    description: 'Schneider Electric cost center profile',
    parameters: [
      { value: '01',        label: 'FR',     payable: true  },
      { value: '01|100',    label: 'FR',     payable: true  },
      { value: '05',        label: 'Eurotherm', payable: true },
      { value: '06',        label: 'UK',     payable: true  },
      { value: '07',        label: 'DK',     payable: true  },
      { value: '08',        label: 'AU',     payable: true  },
      { value: '30',        label: 'NAM',    payable: true  },
      { value: '31',        label: 'SOLAR',  payable: true  },
      { value: '32',        label: 'NAM',    payable: true  },
      { value: '33',        label: 'NAM',    payable: true  },
      { value: '34',        label: 'NAM',    payable: true  },
      { value: '50',        label: 'CN',     payable: true  },
      { value: '51',        label: 'CN',     payable: true  },
      { value: '52',        label: 'CN',     payable: true  },
      { value: '53',        label: 'CN',     payable: true  },
      { value: '54',        label: 'CN',     payable: true  },
      { value: '70',        label: 'JP',     payable: true  },
      { value: '130',       label: 'AU',     payable: true  },
      { value: '131',       label: '',       payable: false },
      { value: '132',       label: '',       payable: false },
      { value: '150',       label: 'CN',     payable: false },
      { value: 'DEFAULT',   label: '',       payable: false },
      { value: 'DEFAULT|100', label: '',     payable: false }
    ],
    createdAt: new Date().toISOString()
  }
];

// ─── Profile Storage ──────────────────────────────────────────────────────────
function loadProfiles() {
  try {
    if (fs.existsSync(PROFILES_PATH)) {
      const data = JSON.parse(fs.readFileSync(PROFILES_PATH, 'utf8'));
      return Array.isArray(data) ? data : DEFAULT_PROFILES;
    }
  } catch (e) {}
  return DEFAULT_PROFILES;
}

function saveProfiles(profiles) {
  try {
    fs.writeFileSync(PROFILES_PATH, JSON.stringify(profiles, null, 2));
    return true;
  } catch (e) {
    return false;
  }
}

// ─── Window ───────────────────────────────────────────────────────────────────
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 580,
    height: 700,
    minWidth: 520,
    minHeight: 600,
    resizable: true,
    frame: true,
    backgroundColor: '#0d0d0d',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    icon: path.join(__dirname, 'assets', 'icon.png')
  });

  mainWindow.loadFile('index.html');
  mainWindow.setMenu(null);
  mainWindow.webContents.on('did-finish-load', () => {
    mainWindow.webContents.send('app-version', app.getVersion());
  });
}

app.whenReady().then(() => {
  createWindow();

  // ── Auto-updater setup ────────────────────────────────────────────────────
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = require('electron-log');
  autoUpdater.logger.transports.file.level = 'info';

  const settings = loadSettings();

  function checkForUpdates() {
    autoUpdater.checkForUpdates().catch(err => {
      if (mainWindow) mainWindow.webContents.send('update-error', err.message);
    });
  }

  // Startup check
  if (settings.checkUpdatesOnStartup) {
    setTimeout(checkForUpdates, 3000);
  }

  // Periodic check
  const intervalMs = (settings.updateIntervalHours || 4) * 60 * 60 * 1000;
  setInterval(checkForUpdates, intervalMs);

  autoUpdater.on('checking-for-update', () => {
    if (mainWindow) mainWindow.webContents.send('update-checking');
  });

  autoUpdater.on('update-available', (info) => {
    if (mainWindow) mainWindow.webContents.send('update-available', info.version);
  });

  autoUpdater.on('update-not-available', (info) => {
    if (mainWindow) mainWindow.webContents.send('update-not-available', info.version);
  });

  autoUpdater.on('download-progress', (progress) => {
    if (mainWindow) mainWindow.webContents.send('update-progress', Math.round(progress.percent));
  });

  autoUpdater.on('update-downloaded', () => {
    if (mainWindow) mainWindow.webContents.send('update-downloaded');
  });

  autoUpdater.on('error', (err) => {
    if (mainWindow) mainWindow.webContents.send('update-error', err.message);
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ─── IPC Handlers ─────────────────────────────────────────────────────────────

ipcMain.handle('get-profiles', () => loadProfiles());

ipcMain.handle('save-profiles', (_, profiles) => saveProfiles(profiles));

ipcMain.handle('pick-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Excel Files',
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] },
      { name: 'All Files', extensions: ['*'] }
    ],
    properties: ['openFile', 'multiSelections']
  });
  return result.canceled ? [] : result.filePaths;
});

ipcMain.handle('pick-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Output Folder',
    properties: ['openDirectory', 'createDirectory']
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('read-file-columns', async (_, filePath) => {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!data || data.length === 0) return { error: 'File is empty' };

    const headers = data[0].map((h, i) => ({ index: i, name: String(h || `Column ${i + 1}`) }));
    const totalRows = data.length - 1;
    const sheetNames = workbook.SheetNames;

    // Sample values for each column (first 5 non-empty)
    const samples = {};
    headers.forEach(h => {
      const vals = [];
      for (let r = 1; r < data.length && vals.length < 5; r++) {
        const v = String(data[r][h.index] ?? '');
        if (v && !vals.includes(v)) vals.push(v);
      }
      samples[h.index] = vals;
    });

    return { headers, totalRows, sheetNames, samples };
  } catch (e) {
    return { error: e.message };
  }
});

ipcMain.handle('split-file', async (_, { filePath, columnIndex, parameters, outputDir, keepNonMatching, outputPrefix }) => {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (!data || data.length < 2) return { error: 'File has no data rows' };

    const headers = data[0];
    const rows = data.slice(1);

    // Build parameter map: normalized value -> param config
    const paramMap = new Map();
    parameters.forEach(p => {
      paramMap.set(String(p.value).trim().toLowerCase(), p);
    });

    // ── Date prefix: YYYYMMDD ──────────────────────────────────────────────────
    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const fileBase = outputPrefix || path.basename(filePath, path.extname(filePath));

    // ── Group rows by REGION LABEL (param.label) ──────────────────────────────
    // Key = label string (e.g. "NAM", "FR", "CN"). Empty label → use value itself.
    const regionGroups = new Map();   // label -> { rows[], costCenters: Set }
    const excludedRows = [];          // payable=false rows
    const nonMatchingRows = [];       // not in profile at all

    rows.forEach(row => {
      const cellVal = String(row[columnIndex] ?? '').trim();
      const key = cellVal.toLowerCase();
      const match = paramMap.get(key);

      if (match) {
        if (match.payable) {
          // Use label as region key; fall back to the value itself if label is empty
          const regionKey = (match.label && match.label.trim()) ? match.label.trim() : String(match.value);
          if (!regionGroups.has(regionKey)) {
            regionGroups.set(regionKey, { rows: [], costCenters: new Set() });
          }
          regionGroups.get(regionKey).rows.push(row);
          regionGroups.get(regionKey).costCenters.add(String(match.value));
        } else {
          // Explicitly excluded by profile
          excludedRows.push(row);
        }
      } else {
        // Not in profile at all
        nonMatchingRows.push(row);
      }
    });

    const created = [];
    const skipped = [];

    // ── Write one file per region ──────────────────────────────────────────────
    for (const [region, group] of regionGroups) {
      if (group.rows.length === 0) { skipped.push(region); continue; }

      const safeRegion = region.replace(/[/\\?%*:|"<>]/g, '-');
      // Format: YYYYMMDD_[fileBase]_[REGION].xlsx
      const fileName = `${dateStr}_${fileBase}_${safeRegion}.xlsx`;
      const outPath = path.join(outputDir, fileName);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, ...group.rows]);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, outPath);

      created.push({
        file: fileName,
        rows: group.rows.length,
        region,
        costCenters: Array.from(group.costCenters).sort()
      });
    }

    // ── Non-matching / excluded file (optional) ────────────────────────────────
    const extraRows = [
      ...(keepNonMatching ? nonMatchingRows : []),
      ...(keepNonMatching ? excludedRows : [])
    ];
    if (extraRows.length > 0) {
      const fileName = `${dateStr}_${fileBase}_NON_MATCHING.xlsx`;
      const outPath = path.join(outputDir, fileName);
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, ...extraRows]);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, outPath);
      created.push({ file: fileName, rows: extraRows.length, region: 'NON_MATCHING', costCenters: [] });
    }

    return { success: true, created, skipped, outputDir };
  } catch (e) {
    return { error: e.message };
  }
});

ipcMain.handle('open-folder', (_, folderPath) => {
  shell.openPath(folderPath);
});

ipcMain.handle('install-update', () => {
  autoUpdater.quitAndInstall();
});

ipcMain.handle('get-settings', () => loadSettings());
ipcMain.handle('save-settings', (_, settings) => saveSettings(settings));
ipcMain.handle('check-for-updates', () => {
  autoUpdater.checkForUpdates().catch(err => {
    if (mainWindow) mainWindow.webContents.send('update-error', err.message);
  });
});
ipcMain.handle('open-external', (_, url) => shell.openExternal(url));
