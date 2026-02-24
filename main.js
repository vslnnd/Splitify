const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const PROFILES_PATH = path.join(app.getPath('userData'), 'splitify_profiles.json');
const SETTINGS_PATH = path.join(app.getPath('userData'), 'splitify_settings.json');
const HISTORY_PATH = path.join(app.getPath('userData'), 'splitify_history.json');

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

// ─── Default SElectric Profile ────────────────────────────────────────────────
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
    if (mainWindow) mainWindow.webContents.send('update-progress', {
      percent: Math.round(progress.percent),
      bytesPerSecond: progress.bytesPerSecond,
      transferred: progress.transferred,
      total: progress.total
    });
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

ipcMain.handle('show-message-box', async (_, options) => {
  const result = await dialog.showMessageBox(mainWindow, options);
  return result.response;
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

      // Copy column widths and styles
      worksheet.columns.forEach((col, i) => {
        if (!col) return;
        const outCol = outWs.getColumn(i + 1);
        if (col.width)  outCol.width  = col.width;
        if (col.hidden) outCol.hidden = col.hidden;
        if (col.style)  outCol.style  = JSON.parse(JSON.stringify(col.style));
      });

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
            // mergeRange is like "A1:C1"
            const decoded = ExcelJS.utils
              ? null // not available this way
              : null;
            // Parse manually
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
  shell.openPath(folderPath);
});

ipcMain.handle('open-file', (_, filePath) => {
  shell.openPath(filePath);
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

ipcMain.handle('get-history', () => loadHistory());
ipcMain.handle('add-history-entry', (_, entry) => {
  const h = loadHistory();
  h.unshift(entry);
  if (h.length > 100) h.splice(100);
  return saveHistory(h);
});
ipcMain.handle('clear-history', () => saveHistory([]));
