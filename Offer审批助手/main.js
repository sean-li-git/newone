// ========== EPIPE 防护（打包后无终端时必须） ==========
const _safeWrite = function() { return true; };
if (process.stdout) {
  process.stdout.write = _safeWrite;
  if (typeof process.stdout.on === 'function') process.stdout.on('error', () => {});
}
if (process.stderr) {
  process.stderr.write = _safeWrite;
  if (typeof process.stderr.on === 'function') process.stderr.on('error', () => {});
}
const _noop = () => {};
console.log = _noop;
console.error = _noop;
console.warn = _noop;
console.info = _noop;
console.debug = _noop;
process.on('uncaughtException', (err) => {
  if (err && (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED')) return;
});

const { app, BrowserWindow, ipcMain, dialog, session } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// ========== 完全禁止网络（核心安全措施） ==========
app.on('ready', () => {
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    if (details.url.startsWith('file://') || details.url.startsWith('devtools://')) {
      callback({ cancel: false });
    } else {
      callback({ cancel: true });
    }
  });
  session.defaultSession.setPermissionRequestHandler((webContents, permission, callback) => {
    callback(false);
  });
});

let mainWindow = null;

// ========== 创建主窗口 ==========
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 1024,
    minHeight: 680,
    show: true,
    center: true,
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: { x: 12, y: 18 },
    resizable: true,
    backgroundColor: '#F8FAFC',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: true,
      allowRunningInsecureContent: false,
    },
  });

  mainWindow.loadFile('index.html');

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ========== 启动 ==========
app.whenReady().then(() => {
  createMainWindow();
});

app.on('window-all-closed', () => {
  app.quit();
});

app.on('activate', () => {
  if (!mainWindow) createMainWindow();
});

// ========== IPC: 选择Excel文件 ==========
ipcMain.handle('select-excel-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择Excel文件',
    filters: [{ name: 'Excel文件', extensions: ['xlsx', 'xls', 'csv'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

// ========== 全局：保存原始文件路径 ==========
let lastReadFilePath = '';

// ========== IPC: 读取Excel文件 ==========
ipcMain.handle('read-excel', async (event, filePath) => {
  try {
    lastReadFilePath = filePath;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheetNames = [];
    const sheets = {};

    workbook.eachSheet((worksheet) => {
      const sheetName = worksheet.name;
      sheetNames.push(sheetName);

      const rowCount = worksheet.rowCount;
      const colCount = worksheet.columnCount;
      const jsonData = [];
      const formulas = {};

      for (let r = 1; r <= rowCount; r++) {
        const row = worksheet.getRow(r);
        const rowData = [];
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c);
          if (cell.formula || cell.sharedFormula) {
            const formulaStr = cell.formula || cell.sharedFormula;
            const rowIdx = r - 1;
            const colIdx = c - 1;
            if (!formulas[rowIdx]) formulas[rowIdx] = {};
            formulas[rowIdx][colIdx] = formulaStr;
          }
          let val = cell.value;
          if (val === null || val === undefined) {
            val = '';
          } else if (typeof val === 'object') {
            if (val instanceof Date) {
              val = val.toISOString().split('T')[0];
            } else if (val.formula) {
              const rowIdx = r - 1;
              const colIdx = c - 1;
              if (!formulas[rowIdx]) formulas[rowIdx] = {};
              formulas[rowIdx][colIdx] = val.formula;
              val = val.result !== undefined && val.result !== null ? val.result : '';
            } else if (val.sharedFormula) {
              const rowIdx = r - 1;
              const colIdx = c - 1;
              if (!formulas[rowIdx]) formulas[rowIdx] = {};
              formulas[rowIdx][colIdx] = val.sharedFormula;
              val = val.result !== undefined && val.result !== null ? val.result : '';
            } else if (val.richText) {
              val = val.richText.map(rt => rt.text).join('');
            } else if (val.text) {
              val = val.text;
            } else if (val.error) {
              val = val.error;
            } else {
              val = String(val);
            }
          }
          rowData.push(val);
        }
        jsonData.push(rowData);
      }

      let maxColWithData = 0;
      for (let row = 0; row < jsonData.length; row++) {
        for (let col = jsonData[row].length - 1; col >= 0; col--) {
          const val = jsonData[row][col];
          if (val !== undefined && val !== null && val !== '') {
            if (col + 1 > maxColWithData) maxColWithData = col + 1;
            break;
          }
        }
      }
      for (const [rStr, cols] of Object.entries(formulas)) {
        for (const cStr of Object.keys(cols)) {
          const c = parseInt(cStr);
          if (c + 1 > maxColWithData) maxColWithData = c + 1;
        }
      }
      if (maxColWithData === 0) maxColWithData = colCount;

      const trimmedData = jsonData.map(row => {
        const trimmed = row.slice(0, maxColWithData);
        while (trimmed.length < maxColWithData) trimmed.push('');
        return trimmed;
      });

      sheets[sheetName] = { data: trimmedData, formulas };
    });

    return { sheetNames, sheets };
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导出Excel ==========
ipcMain.handle('export-excel', async (event, { sheetNames, sheets, defaultName, sourceFilePath }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '导出Excel文件',
    defaultPath: defaultName || '导出数据.xlsx',
    filters: [{ name: 'Excel文件', extensions: ['xlsx'] }],
  });
  if (result.canceled) return null;

  try {
    const srcPath = sourceFilePath || lastReadFilePath;

    if (srcPath && fs.existsSync(srcPath)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(srcPath);

      for (const sheetName of sheetNames) {
        const sheetInfo = sheets[sheetName];
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet || !sheetInfo) continue;

        const data = sheetInfo.data;
        const formulas = sheetInfo.formulas || {};

        for (let r = 0; r < data.length; r++) {
          const row = worksheet.getRow(r + 1);
          for (let c = 0; c < data[r].length; c++) {
            const cell = row.getCell(c + 1);
            if (formulas[r] && formulas[r][c]) {
              cell.value = { formula: formulas[r][c] };
            } else {
              const newVal = data[r][c];
              cell.value = newVal !== undefined && newVal !== null ? newVal : '';
            }
          }
        }
        for (let r = 1; r <= data.length; r++) {
          worksheet.getRow(r).commit();
        }
      }

      await workbook.xlsx.writeFile(result.filePath);
    } else {
      const workbook = new ExcelJS.Workbook();
      for (const sheetName of sheetNames) {
        const sheetInfo = sheets[sheetName];
        const worksheet = workbook.addWorksheet(sheetName);
        const data = sheetInfo.data;
        const formulas = sheetInfo.formulas || {};

        for (let r = 0; r < data.length; r++) {
          const row = worksheet.getRow(r + 1);
          for (let c = 0; c < data[r].length; c++) {
            const cell = row.getCell(c + 1);
            if (formulas[r] && formulas[r][c]) {
              cell.value = { formula: formulas[r][c] };
            } else {
              cell.value = data[r][c] !== undefined && data[r][c] !== null ? data[r][c] : '';
            }
          }
          row.commit();
        }
      }
      await workbook.xlsx.writeFile(result.filePath);
    }

    return result.filePath;
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导出 JSON ==========
ipcMain.handle('export-json', async (event, { data, defaultName }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存文件',
    defaultPath: defaultName || 'data.json',
    filters: [{ name: 'JSON文件', extensions: ['json'] }],
  });
  if (result.canceled) return null;

  try {
    fs.writeFileSync(result.filePath, JSON.stringify(data, null, 2), 'utf-8');
    return result.filePath;
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导入 JSON ==========
ipcMain.handle('import-json', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择 JSON 文件',
    filters: [{ name: 'JSON文件', extensions: ['json'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;

  try {
    const data = fs.readFileSync(result.filePaths[0], 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return { error: e.message };
  }
});
