const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const { autoUpdater, CancellationToken } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const https = require('https');
const ExcelJS = require('exceljs');

const PROFILES_PATH   = path.join(app.getPath('userData'), 'splitify_profiles.json');
const SETTINGS_PATH   = path.join(app.getPath('userData'), 'splitify_settings.json');
const HISTORY_PATH    = path.join(app.getPath('userData'), 'splitify_history.json');
const FAVFOLDERS_PATH = path.join(app.getPath('userData'), 'splitify_favfolders.json');

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
  } catch (e) { }
  return { ...DEFAULT_SETTINGS };
}

function saveSettings(settings) {
  try { fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2)); return true; }
  catch (e) { return false; }
}

// ─── History Storage ──────────────────────────────────────────────────────────
function loadHistory() {
  try {
    if (fs.existsSync(HISTORY_PATH)) return JSON.parse(fs.readFileSync(HISTORY_PATH, 'utf8'));
  } catch (e) { }
  return [];
}
function saveHistory(history) {
  try { fs.writeFileSync(HISTORY_PATH, JSON.stringify(history, null, 2)); return true; }
  catch (e) { return false; }
}

// ─── Fetch release notes from GitHub ─────────────────────────────────────────
function fetchReleaseNotes(version, cb) {
  const options = {
    hostname: 'api.github.com',
    path: `/repos/vslnnd/Splitify/releases/tags/v${version}`,
    method: 'GET',
    headers: {
      'User-Agent': 'Splitify-App',
      'Accept': 'application/vnd.github.v3+json'
    }
  };
  const req = https.request(options, (res) => {
    let raw = '';
    res.on('data', chunk => raw += chunk);
    res.on('end', () => {
      try {
        const data = JSON.parse(raw);
        cb(data.body || '');
      } catch(e) { cb(''); }
    });
  });
  req.on('error', () => cb(''));
  req.end();
}
const DEFAULT_PROFILES = [
  {
    id: 'selectric_' + Date.now(),
    name: 'SElectric',
    description: 'Schneider Electric cost center profile',
    parameters: [
      { value: '01', label: 'FR', payable: true },
      { value: '01|100', label: 'FR', payable: true },
      { value: '05', label: 'Eurotherm', payable: true },
      { value: '06', label: 'UK', payable: true },
      { value: '07', label: 'DK', payable: true },
      { value: '08', label: 'AU', payable: true },
      { value: '30', label: 'NAM', payable: true },
      { value: '31', label: 'SOLAR', payable: true },
      { value: '32', label: 'NAM', payable: true },
      { value: '33', label: 'NAM', payable: true },
      { value: '34', label: 'NAM', payable: true },
      { value: '50', label: 'CN', payable: true },
      { value: '51', label: 'CN', payable: true },
      { value: '52', label: 'CN', payable: true },
      { value: '53', label: 'CN', payable: true },
      { value: '54', label: 'CN', payable: true },
      { value: '70', label: 'JP', payable: true },
      { value: '130', label: 'AU', payable: true },
      { value: '131', label: '', payable: false },
      { value: '132', label: '', payable: false },
      { value: '150', label: 'CN', payable: false },
      { value: 'DEFAULT', label: '', payable: false },
      { value: 'DEFAULT|100', label: '', payable: false }
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
  } catch (e) { }
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
    height: 882,
    minWidth: 520,
    minHeight: 882,
    resizable: true,
    frame: false,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#0d0d0d',
      symbolColor: '#848484',
      height: 46,
    },
    backgroundColor: '#0d0d0d',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    icon: path.join(__dirname, 'assets', 'icon.png')
  });

  mainWindow.loadFile('index.html');
  mainWindow.setMenu(null);
  mainWindow.webContents.on('did-finish-load', () => {
    const currentVersion = app.getVersion();
    mainWindow.webContents.send('app-version', currentVersion);

    // Check if this is the first launch after a genuine upgrade
    const s = loadSettings();
    const semverGt = (a, b) => {
      const pa = a.split('.').map(Number);
      const pb = b.split('.').map(Number);
      for (let i = 0; i < 3; i++) {
        if (pa[i] > pb[i]) return true;
        if (pa[i] < pb[i]) return false;
      }
      return false;
    };
    if (s.lastSeenVersion && semverGt(currentVersion, s.lastSeenVersion)) {
      const prevVersion = s.lastSeenVersion;
      fetchReleaseNotes(currentVersion, (notes) => {
        setTimeout(() => {
          if (mainWindow) mainWindow.webContents.send('first-launch-after-update', {
            prevVersion,
            newVersion: currentVersion,
            notes: notes || ''
          });
        }, 800);
      });
    }
    if (s.lastSeenVersion !== currentVersion) {
      s.lastSeenVersion = currentVersion;
      saveSettings(s);
    }
  });
}

let downloadCancellationToken = null;

app.whenReady().then(() => {
  createWindow();

  // ── Auto-updater setup ────────────────────────────────────────────────────
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = require('electron-log');
  autoUpdater.logger.transports.file.level = 'info';


  const settings = loadSettings();

  function checkForUpdates() {
    // Safety timeout: if no update event fires within 15s, surface an error
    // This prevents the UI from freezing forever on network hangs or silent failures
    const timeout = setTimeout(() => {
      if (mainWindow) mainWindow.webContents.send('update-error', 'Update check timed out. Check your internet connection.');
    }, 15000);

    autoUpdater.checkForUpdates()
      .then(() => clearTimeout(timeout))
      .catch(err => {
        clearTimeout(timeout);
        if (mainWindow) mainWindow.webContents.send('update-error', err.message);
      });
  }

  // Startup check
  if (settings.checkUpdatesOnStartup) {
    setTimeout(checkForUpdates, 3000);
  }

  // Periodic check — re-reads settings each tick so runtime changes take effect
  let periodicTimer = null;
  function scheduleNextCheck() {
    const currentSettings = loadSettings();
    if (!currentSettings.checkUpdatesOnStartup) {
      periodicTimer = setTimeout(scheduleNextCheck, 60 * 60 * 1000); // retry in 1h
      return;
    }
    const intervalMs = (currentSettings.updateIntervalHours || 4) * 60 * 60 * 1000;
    periodicTimer = setTimeout(() => {
      checkForUpdates();
      scheduleNextCheck();
    }, intervalMs);
  }
  scheduleNextCheck();

  autoUpdater.on('checking-for-update', () => {
    if (mainWindow) mainWindow.webContents.send('update-checking');
  });

  autoUpdater.on('update-available', (info) => {
    if (mainWindow) mainWindow.webContents.send('update-available', info);
  });

  autoUpdater.on('update-not-available', (info) => {
    if (mainWindow) mainWindow.webContents.send('update-not-available', info.version);
  });

  autoUpdater.on('download-progress', (progress) => {
    if (mainWindow) mainWindow.webContents.send('update-progress', {
      percent: Math.round(progress.percent),
      bytesPerSecond: progress.bytesPerSecond,
      transferred: progress.transferred,
      total: progress.total
    });
  });

  autoUpdater.on('update-downloaded', (info) => {
    if (mainWindow) mainWindow.webContents.send('update-downloaded', info);
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

ipcMain.handle('show-message-box', async (_, options) => {
  const result = await dialog.showMessageBox(mainWindow, options);
  return result.response;
});

ipcMain.handle('read-file-columns', async (_, filePath) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) return { error: 'File is empty' };

    const sheetNames = workbook.worksheets.map(ws => ws.name);

    // Read header row
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
      const name = cell.value !== null && cell.value !== undefined
        ? (cell.value.richText ? cell.value.richText.map(r => r.text).join('') : String(cell.value))
        : `Column ${colNum}`;
      headers.push({ index: colNum - 1, name: name.trim() || `Column ${colNum}` });
    });

    if (headers.length === 0) return { error: 'File is empty' };

    const totalRows = worksheet.rowCount - 1;

    // Sample up to 5 unique non-empty values per column
    const samples = {};
    headers.forEach(h => { samples[h.index] = []; });

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      headers.forEach(h => {
        if (samples[h.index].length >= 5) return;
        const cell = row.getCell(h.index + 1);
        if (cell.value === null || cell.value === undefined) return;
        const v = cell.value.richText
          ? cell.value.richText.map(r => r.text).join('')
          : String(cell.value);
        const trimmed = v.trim();
        if (trimmed && !samples[h.index].includes(trimmed)) {
          samples[h.index].push(trimmed);
        }
      });
    });

    return { headers, totalRows, sheetNames, samples };
  } catch (e) {
    return { error: e.message };
  }
});



ipcMain.handle('split-file', async (_, { filePath, columnIndex, parameters, outputDir, keepNonMatching, outputPrefix }) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) return { error: 'File has no sheets' };
    const sheetName = worksheet.name;

    // Build parameter map
    const paramMap = new Map();
    parameters.forEach(p => {
      paramMap.set(String(p.value).trim().toLowerCase(), p);
    });

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const fileBase = outputPrefix || path.basename(filePath, path.extname(filePath));

    // Collect header row and group data rows by region
    const headerRow = worksheet.getRow(1);
    const regionGroups = new Map();
    const excludedRows = [];
    const nonMatchingRows = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const cell = row.getCell(columnIndex + 1); // ExcelJS is 1-indexed
      let cellVal = '';
      if (cell.value !== null && cell.value !== undefined) {
        // Handle rich text objects
        cellVal = (cell.value && typeof cell.value === 'object' && cell.value.richText)
          ? cell.value.richText.map(r => r.text).join('')
          : String(cell.value);
        cellVal = cellVal.trim();
      }
      const match = paramMap.get(cellVal.toLowerCase());

      if (match) {
        if (match.payable) {
          const regionKey = (match.label && match.label.trim()) ? match.label.trim() : String(match.value);
          if (!regionGroups.has(regionKey)) regionGroups.set(regionKey, { rows: [], rowNumbers: [], costCenters: new Set() });
          regionGroups.get(regionKey).rows.push(row);
          regionGroups.get(regionKey).rowNumbers.push(rowNumber);
          regionGroups.get(regionKey).costCenters.add(String(match.value));
        } else {
          excludedRows.push({ row, rowNumber });
        }
      } else {
        nonMatchingRows.push({ row, rowNumber });
      }
    });

    // Deep-copy all style properties from one cell to another
    function copyStyle(srcCell, destCell) {
      try {
        if (srcCell.font)       destCell.font       = JSON.parse(JSON.stringify(srcCell.font));
        if (srcCell.fill)       destCell.fill       = JSON.parse(JSON.stringify(srcCell.fill));
        if (srcCell.border)     destCell.border     = JSON.parse(JSON.stringify(srcCell.border));
        if (srcCell.alignment)  destCell.alignment  = JSON.parse(JSON.stringify(srcCell.alignment));
        if (srcCell.numFmt)     destCell.numFmt     = srcCell.numFmt;
        if (srcCell.protection) destCell.protection = JSON.parse(JSON.stringify(srcCell.protection));
      } catch(e) {}
    }

    // Copy a source row into a destination worksheet at a given row number
    function copyRow(srcRow, destWs, destRowNum) {
      const destRow = destWs.getRow(destRowNum);
      if (srcRow.height) destRow.height = srcRow.height;

      srcRow.eachCell({ includeEmpty: true }, (srcCell, colNum) => {
        const destCell = destRow.getCell(colNum);

        // Copy value — handle formulas, rich text, dates, etc.
        if (srcCell.type === ExcelJS.ValueType.Formula) {
          destCell.value = { formula: srcCell.formula, result: srcCell.result };
        } else if (srcCell.type === ExcelJS.ValueType.RichText) {
          destCell.value = JSON.parse(JSON.stringify(srcCell.value));
        } else if (srcCell.type === ExcelJS.ValueType.Hyperlink) {
          destCell.value = JSON.parse(JSON.stringify(srcCell.value));
        } else {
          destCell.value = srcCell.value;
        }

        copyStyle(srcCell, destCell);
      });

      destRow.commit();
    }

    // Build output workbook with header + given rows, remapping merges correctly
    async function writeRegionFile(outPath, srcRowEntries) {
      const outWb = new ExcelJS.Workbook();
      outWb.creator  = workbook.creator || 'Splitify';
      outWb.modified = new Date();

      const outWs = outWb.addWorksheet(sheetName);

      // Copy column widths and styles — guard against files with no column definitions
      try {
        if (worksheet.columns && worksheet.columns.length) {
          worksheet.columns.forEach((col, i) => {
            if (!col) return;
            const outCol = outWs.getColumn(i + 1);
            if (col.width)  outCol.width  = col.width;
            if (col.hidden) outCol.hidden = col.hidden;
            if (col.style)  outCol.style  = JSON.parse(JSON.stringify(col.style));
          });
        }
      } catch(e) {}

      // Build mapping: source row number -> destination row number
      // Header (row 1 in source) -> row 1 in dest
      // Data rows start at dest row 2
      const srcToDestRow = new Map();
      srcToDestRow.set(1, 1);
      srcRowEntries.forEach(({ rowNumber }, idx) => {
        srcToDestRow.set(rowNumber, idx + 2);
      });

      // Remap merges: only include merges where both start and end rows are in the output
      const merges = worksheet.model && worksheet.model.merges;
      if (merges && merges.length) {
        merges.forEach(mergeRange => {
          try {
            const match = mergeRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
            if (!match) return;
            const srcStartRow = parseInt(match[2]);
            const srcEndRow   = parseInt(match[4]);
            const startCol = match[1];
            const endCol   = match[3];
            if (srcToDestRow.has(srcStartRow) && srcToDestRow.has(srcEndRow)) {
              const destStart = srcToDestRow.get(srcStartRow);
              const destEnd   = srcToDestRow.get(srcEndRow);
              outWs.mergeCells(`${startCol}${destStart}:${endCol}${destEnd}`);
            }
          } catch(e) {}
        });
      }

      // Write header row
      copyRow(headerRow, outWs, 1);

      // Write data rows
      srcRowEntries.forEach(({ row }, idx) => {
        copyRow(row, outWs, idx + 2);
      });

      await outWb.xlsx.writeFile(outPath);
    }

    const created = [];
    const skipped = [];

    for (const [region, group] of regionGroups) {
      if (group.rows.length === 0) { skipped.push(region); continue; }

      const safeRegion = region.replace(/[/\\?%*:|"<>]/g, '-');
      const fileName = `${dateStr}_${fileBase}_${safeRegion}.xlsx`;
      const outPath = path.join(outputDir, fileName);

      const entries = group.rows.map((row, i) => ({ row, rowNumber: group.rowNumbers[i] }));
      await writeRegionFile(outPath, entries);

      created.push({
        file: fileName,
        rows: group.rows.length,
        region,
        costCenters: Array.from(group.costCenters).sort()
      });
    }

    const extraEntries = [
      ...(keepNonMatching ? nonMatchingRows : []),
      ...(keepNonMatching ? excludedRows   : [])
    ];
    if (extraEntries.length > 0) {
      const fileName = `${dateStr}_${fileBase}_NON_MATCHING.xlsx`;
      const outPath = path.join(outputDir, fileName);
      await writeRegionFile(outPath, extraEntries);
      created.push({ file: fileName, rows: extraEntries.length, region: 'NON_MATCHING', costCenters: [] });
    }

    return { success: true, created, skipped, outputDir };
  } catch (e) {
    return { error: e.message };
  }
});


ipcMain.handle('open-folder', (_, folderPath) => {
  if (typeof folderPath === 'string' && path.isAbsolute(folderPath)) {
    shell.openPath(folderPath);
  }
});

ipcMain.handle('open-file', (_, filePath) => {
  if (typeof filePath === 'string' && path.isAbsolute(filePath)) {
    shell.openPath(filePath);
  }
});

ipcMain.handle('approve-download', () => {
  downloadCancellationToken = new CancellationToken();
  autoUpdater.downloadUpdate(downloadCancellationToken).catch(err => {
    if (mainWindow) mainWindow.webContents.send('update-error', err.message);
  });
});

ipcMain.handle('cancel-download', () => {
  if (downloadCancellationToken) {
    downloadCancellationToken.cancel();
    downloadCancellationToken = null;
    return true;
  }
  return false;
});

ipcMain.handle('install-update', () => {
  autoUpdater.quitAndInstall();
});

ipcMain.handle('set-titlebar-overlay', (_, opts) => {
  if (mainWindow && typeof opts === 'object') {
    try { mainWindow.setTitleBarOverlay(opts); } catch (_) {}
  }
});

ipcMain.handle('get-settings', () => loadSettings());
ipcMain.handle('save-settings', (_, settings) => saveSettings(settings));
ipcMain.handle('check-for-updates', () => {
  autoUpdater.checkForUpdates().catch(err => {
    if (mainWindow) mainWindow.webContents.send('update-error', err.message);
  });
});
ipcMain.handle('open-external', (_, url) => {
  try {
    const parsed = new URL(url);
    if (['https:', 'http:', 'mailto:'].includes(parsed.protocol)) {
      shell.openExternal(url);
    }
  } catch (_) {} // invalid URL, do nothing
});

ipcMain.handle('get-history', () => loadHistory());
ipcMain.handle('add-history-entry', (_, entry) => {
  const h = loadHistory();
  h.unshift(entry);
  if (h.length > 100) h.splice(100);
  return saveHistory(h);
});
ipcMain.handle('clear-history', () => saveHistory([]));

ipcMain.handle('fetch-release-notes', async (_, version) => {
  if (typeof version !== 'string' || !/^\d+\.\d+\.\d+$/.test(version)) return '';
  return new Promise(resolve => fetchReleaseNotes(version, resolve));
});

// ─── Favorite Folders ─────────────────────────────────────────────────────────
function loadFavFolders() {
  try {
    if (fs.existsSync(FAVFOLDERS_PATH)) return JSON.parse(fs.readFileSync(FAVFOLDERS_PATH, 'utf8'));
  } catch(e) {}
  return [];
}
function saveFavFolders(folders) {
  try { fs.writeFileSync(FAVFOLDERS_PATH, JSON.stringify(folders, null, 2)); return true; }
  catch(e) { return false; }
}

ipcMain.handle('get-fav-folders',   ()          => loadFavFolders());
ipcMain.handle('add-fav-folder',    (_, folder) => {
  const list = loadFavFolders();
  if (!list.includes(folder)) list.unshift(folder);
  return saveFavFolders(list);
});
ipcMain.handle('remove-fav-folder', (_, folder) => {
  return saveFavFolders(loadFavFolders().filter(f => f !== folder));
});
