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
  // 拦截所有网络请求，全部阻止
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    // 只允许 file:// 协议（本地文件）和 devtools://
    if (details.url.startsWith('file://') || details.url.startsWith('devtools://')) {
      callback({ cancel: false });
    } else {
      callback({ cancel: true });
    }
  });

  // 禁用所有权限请求（摄像头、麦克风、地理位置等）
  session.defaultSession.setPermissionRequestHandler((webContents, permission, callback) => {
    callback(false);
  });
});

let mainWindow = null;

// ========== 创建主窗口 ==========
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 750,
    minWidth: 900,
    minHeight: 600,
    show: true,
    center: true,
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: { x: 12, y: 12 },
    resizable: true,
    backgroundColor: '#0f0f1a',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      // 禁用远程内容
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

// ========== 全局：保存原始文件路径（用于导出时保留格式） ==========
let lastReadFilePath = '';

// ========== IPC: 读取Excel文件 ==========
ipcMain.handle('read-excel', async (event, filePath) => {
  try {
    // 保存原始文件路径，导出时基于它修改以保留格式
    lastReadFilePath = filePath;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheetNames = [];
    const sheets = {};

    workbook.eachSheet((worksheet) => {
      const sheetName = worksheet.name;
      sheetNames.push(sheetName);

      // 获取实际有数据的范围
      const rowCount = worksheet.rowCount;
      const colCount = worksheet.columnCount;

      // 构造二维数组数据（类似 xlsx 的 sheet_to_json header:1 模式）
      const jsonData = [];
      const formulas = {};

      for (let r = 1; r <= rowCount; r++) {
        const row = worksheet.getRow(r);
        const rowData = [];
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c);
          // 提取公式
          if (cell.formula || cell.sharedFormula) {
            const formulaStr = cell.formula || cell.sharedFormula;
            const rowIdx = r - 1; // 转为 0-based
            const colIdx = c - 1;
            if (!formulas[rowIdx]) formulas[rowIdx] = {};
            formulas[rowIdx][colIdx] = formulaStr;
          }
          // 提取值
          let val = cell.value;
          // exceljs 特殊类型处理
          if (val === null || val === undefined) {
            val = '';
          } else if (typeof val === 'object') {
            if (val instanceof Date) {
              // 日期类型：转为 ISO 字符串
              val = val.toISOString().split('T')[0];
            } else if (val.formula) {
              // 公式对象：取 result
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
              // 富文本：拼接纯文本
              val = val.richText.map(rt => rt.text).join('');
            } else if (val.text) {
              // 超链接等
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

      // 计算实际有内容的最大列数（裁掉尾部空列）
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
      // 也检查公式列
      for (const [rStr, cols] of Object.entries(formulas)) {
        for (const cStr of Object.keys(cols)) {
          const c = parseInt(cStr);
          if (c + 1 > maxColWithData) maxColWithData = c + 1;
        }
      }

      if (maxColWithData === 0) maxColWithData = colCount;

      // 裁剪每行数据
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

// ========== 数据类型自动识别 ==========
function detectColumnType(header, values) {
  const headerLower = (header || '').toString().toLowerCase();
  const headerStr = header || '';

  // 基于列名关键词匹配
  const nameKeywords = ['姓名', '名字', '联系人', '负责人', '经办人', '员工姓名', 'name'];
  const deptKeywords = ['部门', '团队', '组织', '区域', '分公司', '事业部', 'dept', 'department'];
  const idCardKeywords = ['身份证', '证件号', '身份号', 'id card', 'identity'];
  const bankKeywords = ['银行卡', '卡号', '账号', 'bank'];
  const phoneKeywords = ['手机', '电话', '联系方式', 'phone', 'mobile', 'tel'];
  const salaryKeywords = ['薪资', '工资', '薪酬', '底薪', '基本工资', '月薪', '年薪', 'salary', 'wage', 'pay'];
  const amountKeywords = ['金额', '奖金', '提成', '补贴', '津贴', '社保', '公积金', '扣款', '实发', '应发', '税', 'amount', 'bonus'];
  const perfKeywords = ['业绩', '营收', '收入', '产值', '销售额', '利润', '产量', '指标', 'performance', 'revenue', 'sales'];
  const dateKeywords = ['日期', '时间', '入职', '生日', '出生', '离职', 'date', 'time'];
  const seqKeywords = ['序号', '编号', '行号', '#', 'no', 'seq', 'index'];
  const jobKeywords = ['职级', '职位', '岗位', '职称', 'level', 'rank', 'title', 'position'];
  const empIdKeywords = ['工号', '员工编号', '人员编号', 'emp id', 'employee id', 'staff id'];

  const matchKeyword = (keywords) => keywords.some(k => headerStr.includes(k) || headerLower.includes(k.toLowerCase()));

  if (matchKeyword(nameKeywords)) return { type: 'name', label: '人名', strategy: 'mapping', defaultPrefix: 'P' };
  if (matchKeyword(idCardKeywords)) return { type: 'idcard', label: '证件号', strategy: 'format-replace', defaultPrefix: '' };
  if (matchKeyword(bankKeywords)) return { type: 'bankcard', label: '银行卡', strategy: 'format-replace', defaultPrefix: '' };
  if (matchKeyword(phoneKeywords)) return { type: 'phone', label: '手机号', strategy: 'format-replace', defaultPrefix: '' };
  if (matchKeyword(empIdKeywords)) return { type: 'empid', label: '员工编号', strategy: 'mapping', defaultPrefix: 'E' };
  if (matchKeyword(deptKeywords)) return { type: 'department', label: '组织/部门', strategy: 'mapping', defaultPrefix: 'D' };
  if (matchKeyword(salaryKeywords)) return { type: 'salary', label: '薪资金额', strategy: 'scale', defaultPrefix: '' };
  if (matchKeyword(amountKeywords)) return { type: 'amount', label: '金额', strategy: 'scale', defaultPrefix: '' };
  if (matchKeyword(perfKeywords)) return { type: 'performance', label: '业绩数据', strategy: 'scale', defaultPrefix: '' };
  if (matchKeyword(dateKeywords)) return { type: 'date', label: '日期', strategy: 'date-offset', defaultPrefix: '' };
  if (matchKeyword(seqKeywords)) return { type: 'sequence', label: '序号', strategy: 'none', defaultPrefix: '' };
  if (matchKeyword(jobKeywords)) return { type: 'category', label: '分类/枚举', strategy: 'mapping', defaultPrefix: 'L' };

  // 基于数据内容判断
  const sampleValues = values.filter(v => v !== '' && v !== null && v !== undefined).slice(0, 50);
  if (sampleValues.length === 0) return { type: 'unknown', label: '未知', strategy: 'none', defaultPrefix: '' };

  // 检查是否为身份证号
  const idCardPattern = /^\d{17}[\dXx]$/;
  if (sampleValues.length > 0 && sampleValues.every(v => idCardPattern.test(String(v).trim()))) {
    return { type: 'idcard', label: '证件号', strategy: 'format-replace', defaultPrefix: '' };
  }

  // 检查是否为手机号
  const phonePattern = /^1[3-9]\d{9}$/;
  if (sampleValues.length > 0 && sampleValues.every(v => phonePattern.test(String(v).trim()))) {
    return { type: 'phone', label: '手机号', strategy: 'format-replace', defaultPrefix: '' };
  }

  // 检查是否为纯数字
  const numValues = sampleValues.filter(v => !isNaN(Number(v)));
  if (numValues.length > sampleValues.length * 0.8) {
    const nums = numValues.map(Number);
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    // 序号特征：连续整数，从1开始
    if (min >= 0 && max <= sampleValues.length * 2 && nums.every(n => Number.isInteger(n))) {
      return { type: 'sequence', label: '序号', strategy: 'none', defaultPrefix: '' };
    }
    // 大数值，可能是金额
    if (max > 1000) {
      return { type: 'amount', label: '数值(可能是金额)', strategy: 'scale', defaultPrefix: '' };
    }
    return { type: 'number', label: '数值', strategy: 'scale', defaultPrefix: '' };
  }

  // 检查是否为日期
  const dateValues = sampleValues.filter(v => {
    const d = new Date(v);
    return !isNaN(d.getTime()) && String(v).match(/[\-\/年月日]/);
  });
  if (dateValues.length > sampleValues.length * 0.8) {
    return { type: 'date', label: '日期', strategy: 'date-offset', defaultPrefix: '' };
  }

  // 检查是否为短文本（可能是人名）
  const avgLen = sampleValues.reduce((s, v) => s + String(v).length, 0) / sampleValues.length;
  const hasChinese = sampleValues.some(v => /[\u4e00-\u9fff]/.test(String(v)));
  if (hasChinese && avgLen <= 4) {
    return { type: 'name', label: '文本(可能是人名)', strategy: 'mapping', defaultPrefix: 'P' };
  }

  // 检查是否为分类数据（重复值多）
  const uniqueCount = new Set(sampleValues.map(String)).size;
  if (uniqueCount < sampleValues.length * 0.3) {
    return { type: 'category', label: '分类/枚举', strategy: 'mapping', defaultPrefix: 'C' };
  }

  // 默认：文本
  return { type: 'text', label: '文本', strategy: 'mapping', defaultPrefix: 'T' };
}

// ========== IPC: 自动识别列类型 ==========
ipcMain.handle('detect-columns', async (event, sheetData) => {
  const data = sheetData.data;
  if (!data || data.length < 2) return [];

  const headers = data[0];
  const results = [];

  for (let col = 0; col < headers.length; col++) {
    const header = headers[col];
    const values = [];
    for (let row = 1; row < data.length; row++) {
      if (data[row][col] !== undefined && data[row][col] !== '') {
        values.push(data[row][col]);
      }
    }

    // 跳过没有实际数据的列（表头和数据行都为空，或仅有表头但数据行全空）
    const headerEmpty = (header === undefined || header === null || String(header).trim() === '');
    if (values.length === 0 && headerEmpty) continue;  // 表头和数据都空 → 完全忽略
    if (values.length === 0) {
      // 有表头但数据全空 → 标记为 none，不脱敏
      results.push({
        colIndex: col,
        header: header,
        sample: [],
        type: 'empty',
        label: '空列',
        strategy: 'none',
        defaultPrefix: '',
      });
      continue;
    }

    const detection = detectColumnType(header, values);
    results.push({
      colIndex: col,
      header: header || `列${col + 1}`,
      sample: values.slice(0, 3),
      ...detection,
    });
  }

  // 对同一前缀的多列自动加数字后缀区分（如 P, P2, P3 或 D, D2, D3）
  const prefixCount = {};
  for (const r of results) {
    if (r.defaultPrefix && r.strategy === 'mapping') {
      const p = r.defaultPrefix;
      prefixCount[p] = (prefixCount[p] || 0) + 1;
    }
  }
  // 只对出现 > 1 次的前缀做区分
  const prefixIndex = {};
  for (const r of results) {
    if (r.defaultPrefix && r.strategy === 'mapping') {
      const p = r.defaultPrefix;
      if (prefixCount[p] > 1) {
        prefixIndex[p] = (prefixIndex[p] || 0) + 1;
        // 第1个保持原样，第2个起加数字
        if (prefixIndex[p] > 1) {
          r.defaultPrefix = p + prefixIndex[p];
        }
      }
    }
  }

  return results;
});

// ========== 脱敏引擎 ==========

// 生成随机中文假名
function randomChineseName() {
  const surnames = ['王','李','张','刘','陈','杨','赵','黄','周','吴','徐','孙','胡','朱','高','林','何','郭','马','罗'];
  const chars = ['伟','芳','娜','秀英','敏','静','丽','强','磊','军','洋','勇','艳','杰','娟','涛','明','超','秀兰','霞'];
  return surnames[Math.floor(Math.random() * surnames.length)] + chars[Math.floor(Math.random() * chars.length)] + (Math.random() > 0.5 ? chars[Math.floor(Math.random() * chars.length)] : '');
}

// 生成格式正确的假身份证号
function randomIdCard() {
  const areas = ['110101','310101','440103','330102','320102','510104','420102','610103'];
  const area = areas[Math.floor(Math.random() * areas.length)];
  const year = 1960 + Math.floor(Math.random() * 40);
  const month = String(Math.floor(Math.random() * 12) + 1).padStart(2, '0');
  const day = String(Math.floor(Math.random() * 28) + 1).padStart(2, '0');
  const seq = String(Math.floor(Math.random() * 999) + 1).padStart(3, '0');
  const base = `${area}${year}${month}${day}${seq}`;
  const weights = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2];
  const checkChars = '10X98765432';
  let sum = 0;
  for (let i = 0; i < 17; i++) sum += parseInt(base[i]) * weights[i];
  return base + checkChars[sum % 11];
}

// 生成假银行卡号
function randomBankCard() {
  const prefixes = ['6222','6228','6214','6217','6225'];
  let card = prefixes[Math.floor(Math.random() * prefixes.length)];
  while (card.length < 19) card += Math.floor(Math.random() * 10);
  return card;
}

// 生成假手机号
function randomPhone() {
  const prefixes = ['138','139','136','137','158','159','188','189','135','186'];
  let phone = prefixes[Math.floor(Math.random() * prefixes.length)];
  while (phone.length < 11) phone += Math.floor(Math.random() * 10);
  return phone;
}

// ========== IPC: 执行脱敏 ==========
ipcMain.handle('desensitize', async (event, { sheetNames, sheets, configs, baseKeyData }) => {
  try {
    const keyData = {
      version: '1.0',
      timestamp: new Date().toISOString(),
      sheets: {},
    };
    // 如果有基准密钥，标记到新密钥中
    if (baseKeyData) {
      keyData.basedOn = baseKeyData.timestamp || 'unknown';
    }

    const resultSheets = {};

    for (const sheetName of sheetNames) {
      const sheetInfo = sheets[sheetName];
      const data = sheetInfo.data;
      const formulas = sheetInfo.formulas || {};
      const config = configs[sheetName] || [];

      if (!data || data.length < 2) {
        resultSheets[sheetName] = { data: [...data], formulas };
        continue;
      }

      const resultData = data.map(row => [...row]);
      const sheetKey = { columns: {} };

      for (const colConfig of config) {
        const { colIndex, strategy, params } = colConfig;
        if (strategy === 'none') continue;

        // 整列数据为空（表头以下无任何非空值）→ 跳过不脱敏
        const hasNonEmpty = data.slice(1).some(row => {
          const v = row[colIndex];
          return v !== undefined && v !== null && v !== '';
        });
        if (!hasNonEmpty) continue;

        // 记录列名到密钥中，方便还原时核对
        const headerName = data[0] ? (data[0][colIndex] || `列${colIndex + 1}`) : `列${colIndex + 1}`;
        const colKey = { strategy, header: String(headerName), params: { ...params }, mappings: {} };

        // 获取基准密钥中该列的已有数据（如果有的话）
        const baseColKey = baseKeyData?.sheets?.[sheetName]?.columns?.[colIndex];

        if (strategy === 'mapping') {
          // 映射替换
          const prefix = params?.prefix || 'A';
          const uniqueValues = new Map();
          let counter = 1;
          // 如果有基准密钥的映射表，先载入已有映射
          if (baseColKey?.mappings && (baseColKey.strategy === 'mapping' || baseColKey.strategy === 'fakename' || baseColKey.strategy === 'format-replace')) {
            for (const [orig, masked] of Object.entries(baseColKey.mappings)) {
              uniqueValues.set(orig, masked);
              // 从已有映射中提取最大编号，避免新编号冲突
              const match = masked.match(/(\d+)$/);
              if (match) {
                const num = parseInt(match[1]);
                if (num >= counter) counter = num + 1;
              }
            }
          }
          for (let row = 1; row < data.length; row++) {
            // 跳过有公式的单元格
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (val && !uniqueValues.has(val)) {
              uniqueValues.set(val, `${prefix}${String(counter).padStart(3, '0')}`);
              counter++;
            }
          }
          // 保存映射表（包含旧映射 + 新增映射）
          colKey.mappings = Object.fromEntries(uniqueValues);
          // 应用
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (uniqueValues.has(val)) {
              resultData[row][colIndex] = uniqueValues.get(val);
            }
          }
        } else if (strategy === 'fakename') {
          // 随机假名替换
          const uniqueValues = new Map();
          // 复用基准密钥中已有的假名映射
          if (baseColKey?.mappings) {
            for (const [orig, masked] of Object.entries(baseColKey.mappings)) {
              uniqueValues.set(orig, masked);
            }
          }
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (val && !uniqueValues.has(val)) {
              uniqueValues.set(val, randomChineseName());
            }
          }
          colKey.mappings = Object.fromEntries(uniqueValues);
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (uniqueValues.has(val)) {
              resultData[row][colIndex] = uniqueValues.get(val);
            }
          }
        } else if (strategy === 'scale') {
          // 等比缩放：优先用用户指定值 > 基准密钥值 > 随机
          const factor = params?.factor || baseColKey?.params?.factor || (0.5 + Math.random() * 0.8);
          colKey.params.factor = factor;
          for (let row = 1; row < data.length; row++) {
            // 跳过有公式的单元格
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              resultData[row][colIndex] = Math.round(val * factor * 100) / 100;
            }
          }
        } else if (strategy === 'scale-noise') {
          // 等比缩放 + 随机扰动：优先用用户指定值 > 基准密钥值 > 随机
          const factor = params?.factor || baseColKey?.params?.factor || (0.5 + Math.random() * 0.8);
          const noisePercent = params?.noisePercent || baseColKey?.params?.noisePercent || 2;
          colKey.params.factor = factor;
          colKey.params.noisePercent = noisePercent;
          const noises = {};
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              const noise = val * (noisePercent / 100) * (Math.random() * 2 - 1);
              noises[row] = noise;
              resultData[row][colIndex] = Math.round((val * factor + noise) * 100) / 100;
            }
          }
          colKey.noises = noises;
        } else if (strategy === 'offset') {
          // 固定偏移：优先用用户指定值 > 基准密钥值 > 随机
          const offsetVal = params?.offset || baseColKey?.params?.offset || Math.floor(Math.random() * 10000 - 5000);
          colKey.params.offset = offsetVal;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              resultData[row][colIndex] = Math.round((val + offsetVal) * 100) / 100;
            }
          }
        } else if (strategy === 'range-map') {
          // 区间映射
          const nums = [];
          for (let row = 1; row < data.length; row++) {
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') nums.push(val);
          }
          const origMin = Math.min(...nums);
          const origMax = Math.max(...nums);
          const newMin = params?.newMin || 1000;
          const newMax = params?.newMax || 10000;
          colKey.params = { origMin, origMax, newMin, newMax };
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              const mapped = origMax === origMin
                ? (newMin + newMax) / 2
                : newMin + (val - origMin) / (origMax - origMin) * (newMax - newMin);
              resultData[row][colIndex] = Math.round(mapped * 100) / 100;
            }
          }
        } else if (strategy === 'format-replace') {
          // 格式保留替换（证件号、银行卡、手机号）
          const colType = colConfig.type;
          const uniqueValues = new Map();
          // 复用基准密钥中已有的格式映射
          if (baseColKey?.mappings) {
            for (const [orig, masked] of Object.entries(baseColKey.mappings)) {
              uniqueValues.set(orig, masked);
            }
          }
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (val && !uniqueValues.has(val)) {
              let fake;
              if (colType === 'idcard') fake = randomIdCard();
              else if (colType === 'bankcard') fake = randomBankCard();
              else if (colType === 'phone') fake = randomPhone();
              else fake = `FAKE-${String(uniqueValues.size + 1).padStart(4, '0')}`;
              uniqueValues.set(val, fake);
            }
          }
          colKey.mappings = Object.fromEntries(uniqueValues);
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (uniqueValues.has(val)) {
              resultData[row][colIndex] = uniqueValues.get(val);
            }
          }
        } else if (strategy === 'date-offset') {
          // 日期偏移：优先用用户指定值 > 基准密钥值 > 随机
          const offsetDays = params?.offsetDays || baseColKey?.params?.offsetDays || Math.floor(Math.random() * 365 - 180);
          colKey.params.offsetDays = offsetDays;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = data[row][colIndex];
            if (!val) continue;
            try {
              const d = new Date(val);
              if (!isNaN(d.getTime())) {
                d.setDate(d.getDate() + offsetDays);
                resultData[row][colIndex] = d.toISOString().split('T')[0];
              }
            } catch (e) {}
          }
        }

        sheetKey.columns[colIndex] = colKey;
      }

      keyData.sheets[sheetName] = sheetKey;
      resultSheets[sheetName] = { data: resultData, formulas };
    }

    // ===== 生成操作报告 =====
    const report = { sheets: [] };
    let totalProcessed = 0;
    let totalSkippedFormula = 0;

    for (const sheetName of sheetNames) {
      const sheetInfo = sheets[sheetName];
      const data = sheetInfo.data;
      const formulas = sheetInfo.formulas || {};
      const sheetKey = keyData.sheets[sheetName];
      if (!sheetKey) continue;

      const sheetReport = { name: sheetName, columns: [], totalRows: data.length - 1 };

      for (const [colIndexStr, colKey] of Object.entries(sheetKey.columns)) {
        if (colKey.strategy === 'none') continue;
        const colIndex = parseInt(colIndexStr);

        // 统计该列实际处理的行数和跳过的公式行数
        let processed = 0;
        let skippedFormula = 0;
        for (let row = 1; row < data.length; row++) {
          if (formulas[row] && formulas[row][colIndex]) {
            skippedFormula++;
          } else {
            const val = data[row][colIndex];
            if (val !== undefined && val !== null && val !== '') {
              processed++;
            }
          }
        }

        // 策略中文名映射
        const strategyNames = {
          'mapping': '编号替换', 'fakename': '随机假名', 'format-replace': '格式保留替换',
          'scale': '等比缩放', 'scale-noise': '等比缩放+扰动', 'offset': '固定偏移',
          'range-map': '区间映射', 'date-offset': '日期偏移',
        };

        sheetReport.columns.push({
          header: colKey.header,
          strategy: strategyNames[colKey.strategy] || colKey.strategy,
          processed,
          skippedFormula,
        });
        totalProcessed += processed;
        totalSkippedFormula += skippedFormula;
      }

      if (sheetReport.columns.length > 0) {
        report.sheets.push(sheetReport);
      }
    }
    report.totalProcessed = totalProcessed;
    report.totalSkippedFormula = totalSkippedFormula;

    return { keyData, resultSheets, report };
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 执行还原 ==========
ipcMain.handle('restore', async (event, { sheetNames, sheets, keyData, skipHeaderCheck }) => {
  try {
    const resultSheets = {};

    // 第一阶段：列名校验（可通过 skipHeaderCheck 跳过）
    if (!skipHeaderCheck) {
      const allMismatches = [];
      for (const sheetName of sheetNames) {
        const sheetInfo = sheets[sheetName];
        const data = sheetInfo.data;
        const sheetKey = keyData.sheets[sheetName];
        if (!sheetKey || !data || data.length < 2) continue;
        const headers = data[0] || [];
        for (const [colIndexStr, colKey] of Object.entries(sheetKey.columns)) {
          const colIndex = parseInt(colIndexStr);
          if (colKey.strategy === 'none') continue;
          const currentHeader = String(headers[colIndex] || '').trim();
          const keyHeader = String(colKey.header || '').trim();
          if (keyHeader && currentHeader && keyHeader !== currentHeader) {
            allMismatches.push({ sheet: sheetName, colIndex, keyHeader, currentHeader });
          }
        }
      }
      if (allMismatches.length > 0) {
        return { headerMismatches: allMismatches };
      }
    }

    for (const sheetName of sheetNames) {
      const sheetInfo = sheets[sheetName];
      const data = sheetInfo.data;
      const formulas = sheetInfo.formulas || {};
      const sheetKey = keyData.sheets[sheetName];

      if (!sheetKey || !data || data.length < 2) {
        resultSheets[sheetName] = { data: [...data], formulas };
        continue;
      }

      // 公式单元格在下方各策略循环中会自动跳过（不脱敏也不还原）

      const resultData = data.map(row => [...row]);

      for (const [colIndexStr, colKey] of Object.entries(sheetKey.columns)) {
        const colIndex = parseInt(colIndexStr);
        const { strategy, params, mappings, noises } = colKey;

        if (strategy === 'none') continue;

        if (strategy === 'mapping' || strategy === 'fakename' || strategy === 'format-replace') {
          // 反向映射
          const reverseMap = {};
          for (const [orig, masked] of Object.entries(mappings)) {
            reverseMap[masked] = orig;
          }
          for (let row = 1; row < data.length; row++) {
            // 跳过有公式的单元格
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = String(data[row][colIndex] || '');
            if (reverseMap[val] !== undefined) {
              resultData[row][colIndex] = reverseMap[val];
            }
          }
        } else if (strategy === 'scale') {
          const factor = params.factor;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              resultData[row][colIndex] = Math.round(val / factor * 100) / 100;
            }
          }
        } else if (strategy === 'scale-noise') {
          const factor = params.factor;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              const noise = noises && noises[row] ? noises[row] : 0;
              resultData[row][colIndex] = Math.round((val - noise) / factor * 100) / 100;
            }
          }
        } else if (strategy === 'offset') {
          const offsetVal = params.offset;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              resultData[row][colIndex] = Math.round((val - offsetVal) * 100) / 100;
            }
          }
        } else if (strategy === 'range-map') {
          const { origMin, origMax, newMin, newMax } = params;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = Number(data[row][colIndex]);
            if (!isNaN(val) && data[row][colIndex] !== '') {
              const restored = newMax === newMin
                ? origMin
                : origMin + (val - newMin) / (newMax - newMin) * (origMax - origMin);
              resultData[row][colIndex] = Math.round(restored * 100) / 100;
            }
          }
        } else if (strategy === 'date-offset') {
          const offsetDays = params.offsetDays;
          for (let row = 1; row < data.length; row++) {
            if (formulas[row] && formulas[row][colIndex]) continue;
            const val = data[row][colIndex];
            if (!val) continue;
            try {
              const d = new Date(val);
              if (!isNaN(d.getTime())) {
                d.setDate(d.getDate() - offsetDays);
                resultData[row][colIndex] = d.toISOString().split('T')[0];
              }
            } catch (e) {}
          }
        }
      }

      resultSheets[sheetName] = { data: resultData, formulas };
    }

    // ===== 生成还原操作报告 =====
    const report = { sheets: [] };
    let totalRestored = 0;
    let totalSkippedFormula = 0;

    for (const sheetName of sheetNames) {
      const sheetInfo = sheets[sheetName];
      const data = sheetInfo.data;
      const formulas = sheetInfo.formulas || {};
      const sheetKey = keyData.sheets[sheetName];
      if (!sheetKey) continue;

      const sheetReport = { name: sheetName, columns: [], totalRows: data.length - 1 };

      for (const [colIndexStr, colKey] of Object.entries(sheetKey.columns)) {
        if (colKey.strategy === 'none') continue;
        const colIndex = parseInt(colIndexStr);

        let restored = 0;
        let skippedFormula = 0;
        for (let row = 1; row < data.length; row++) {
          if (formulas[row] && formulas[row][colIndex]) {
            skippedFormula++;
          } else {
            const origVal = String(data[row][colIndex] || '');
            const restoredVal = String(resultSheets[sheetName].data[row][colIndex] || '');
            if (origVal !== restoredVal) {
              restored++;
            }
          }
        }

        const strategyNames = {
          'mapping': '编号替换', 'fakename': '随机假名', 'format-replace': '格式保留替换',
          'scale': '等比缩放', 'scale-noise': '等比缩放+扰动', 'offset': '固定偏移',
          'range-map': '区间映射', 'date-offset': '日期偏移',
        };

        sheetReport.columns.push({
          header: colKey.header,
          strategy: strategyNames[colKey.strategy] || colKey.strategy,
          restored,
          skippedFormula,
        });
        totalRestored += restored;
        totalSkippedFormula += skippedFormula;
      }

      if (sheetReport.columns.length > 0) {
        report.sheets.push(sheetReport);
      }
    }
    report.totalRestored = totalRestored;
    report.totalSkippedFormula = totalSkippedFormula;

    return { resultSheets, report };
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导出脱敏后的Excel（保留原始格式） ==========
ipcMain.handle('export-excel', async (event, { sheetNames, sheets, defaultName, sourceFilePath }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '导出Excel文件',
    defaultPath: defaultName || '脱敏数据.xlsx',
    filters: [{ name: 'Excel文件', extensions: ['xlsx'] }],
  });
  if (result.canceled) return null;

  try {
    // 确定源文件路径：优先使用传入的，其次使用上次读取的
    const srcPath = sourceFilePath || lastReadFilePath;

    if (srcPath && fs.existsSync(srcPath)) {
      // 基于原始文件修改（保留格式）
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(srcPath);

      for (const sheetName of sheetNames) {
        const sheetInfo = sheets[sheetName];
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet || !sheetInfo) continue;

        const data = sheetInfo.data;
        const formulas = sheetInfo.formulas || {};

        for (let r = 0; r < data.length; r++) {
          const row = worksheet.getRow(r + 1); // exceljs 是 1-based
          for (let c = 0; c < data[r].length; c++) {
            const cell = row.getCell(c + 1);
            // 如果有公式，写入公式
            if (formulas[r] && formulas[r][c]) {
              cell.value = { formula: formulas[r][c] };
            } else {
              // 写入值，但保留单元格样式
              const newVal = data[r][c];
              // 保留 style（字体、填充、边框、数字格式等）
              cell.value = newVal !== undefined && newVal !== null ? newVal : '';
            }
          }
        }
        // 确保行被提交
        for (let r = 1; r <= data.length; r++) {
          worksheet.getRow(r).commit();
        }
      }

      await workbook.xlsx.writeFile(result.filePath);
    } else {
      // 没有原始文件时，创建新 workbook（兜底方案，格式会丢失）
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

// ========== IPC: 导出密钥文件 ==========
ipcMain.handle('export-key', async (event, keyData) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存密钥文件（请妥善保管）',
    defaultPath: `脱敏密钥_${new Date().toISOString().slice(0, 10)}.json`,
    filters: [{ name: 'JSON密钥文件', extensions: ['json'] }],
  });
  if (result.canceled) return null;

  try {
    fs.writeFileSync(result.filePath, JSON.stringify(keyData, null, 2), 'utf-8');
    return result.filePath;
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导出 Excel 映射表（方便用户检查） ==========
ipcMain.handle('export-key-excel', async (event, { keyData, sheetHeaders }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存映射表（Excel格式，方便查看）',
    defaultPath: `脱敏映射表_${new Date().toISOString().slice(0, 10)}.xlsx`,
    filters: [{ name: 'Excel文件', extensions: ['xlsx'] }],
  });
  if (result.canceled) return null;

  try {
    const wb = new ExcelJS.Workbook();

    for (const [sheetName, sheetKey] of Object.entries(keyData.sheets)) {
      const headers = sheetHeaders[sheetName] || [];
      const ws = wb.addWorksheet(sheetName.substring(0, 31));

      // 设置列宽
      ws.columns = [
        { width: 16 },
        { width: 16 },
        { width: 24 },
        { width: 24 },
        { width: 32 },
      ];

      // 添加表头
      ws.addRow(['列名', '脱敏策略', '原始值', '脱敏值', '参数']);

      for (const [colIndexStr, colKey] of Object.entries(sheetKey.columns)) {
        const colIndex = parseInt(colIndexStr);
        const colName = headers[colIndex] || `第${colIndex + 1}列`;
        const strategy = colKey.strategy;

        if (strategy === 'none') continue;

        // 有映射关系的策略：逐行列出
        if (colKey.mappings && Object.keys(colKey.mappings).length > 0) {
          let first = true;
          for (const [orig, masked] of Object.entries(colKey.mappings)) {
            ws.addRow([
              first ? colName : '',
              first ? strategy : '',
              orig,
              masked,
              first ? JSON.stringify(colKey.params || {}) : '',
            ]);
            first = false;
          }
        } else {
          // 数值类策略：只显示参数
          const paramStr = Object.entries(colKey.params || {}).map(([k, v]) => `${k}: ${v}`).join(', ');
          ws.addRow([colName, strategy, '(数值类)', '(按参数计算)', paramStr]);
        }
        // 加一个空行分隔
        ws.addRow(['', '', '', '', '']);
      }
    }

    await wb.xlsx.writeFile(result.filePath);
    return result.filePath;
  } catch (e) {
    return { error: e.message };
  }
});

// ========== IPC: 导入密钥文件 ==========
ipcMain.handle('import-key', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择密钥文件',
    filters: [{ name: 'JSON密钥文件', extensions: ['json'] }],
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

// ========== IPC: 通过路径直接导入密钥文件（拖拽用） ==========
ipcMain.handle('import-key-path', async (event, filePath) => {
  try {
    const data = fs.readFileSync(filePath, 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return { error: e.message };
  }
});
