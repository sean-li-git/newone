/* ======================================================================
 * 数据脱敏助手 · 网页版核心逻辑（纯前端，全部本地处理）
 * 使用 SheetJS 读写 Excel；无任何网络请求
 * ====================================================================== */

const APP_VERSION = 'v1.4';
const LS_KEY_PREF = 'desensitizer_prefs_v1';  // 记住「列名 → 策略 + 参数」偏好

// ========== 全局状态 ==========
let dState = { step:1, fileName:'', workbook:null, currentSheet:'', columnConfigs:{}, desensitizedResult:null, keyData:null, baseKeyData:null };
let rState = { step:1, excelName:'', workbook:null, keyData:null, currentSheet:'', restoredResult:null };

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('app-version').textContent = APP_VERSION;
  initDesensitizeUI();
  initRestoreUI();
  initTabs();
});

// ========== 工具 ==========
function showLoading(t){ document.getElementById('loading-text').textContent = t || '处理中...'; document.getElementById('loading').classList.add('show'); }
function hideLoading(){ document.getElementById('loading').classList.remove('show'); }
function downloadBlob(blob, filename){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}
function initTabs(){
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('panel-' + tab).classList.add('active');
    });
  });
}

// ========== Excel 读写 ==========
async function readExcelFile(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:'array', cellDates:false, cellFormula:true });
  const sheetNames = wb.SheetNames;
  const sheets = {};
  for (const name of sheetNames){
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false, blankrows:true });
    const formulas = {};
    if (ws['!ref']){
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = range.s.r; r <= range.e.r; r++){
        for (let c = range.s.c; c <= range.e.c; c++){
          const addr = XLSX.utils.encode_cell({r, c});
          const cell = ws[addr];
          if (cell && cell.f){ if (!formulas[r]) formulas[r] = {}; formulas[r][c] = cell.f; }
        }
      }
    }
    let maxCol = 0;
    for (let r = 0; r < data.length; r++){
      for (let c = data[r].length - 1; c >= 0; c--){
        const v = data[r][c];
        if (v !== undefined && v !== null && v !== ''){ if (c + 1 > maxCol) maxCol = c + 1; break; }
      }
    }
    for (const r of Object.keys(formulas)){
      for (const c of Object.keys(formulas[r])){
        if (parseInt(c) + 1 > maxCol) maxCol = parseInt(c) + 1;
      }
    }
    if (maxCol === 0) maxCol = (data[0] && data[0].length) || 1;
    const trimmed = data.map(row => {
      const r = row.slice(0, maxCol);
      while (r.length < maxCol) r.push('');
      return r;
    });
    sheets[name] = { data: trimmed, formulas };
  }
  return { sheetNames, sheets };
}
function exportExcelFile(sheetNames, sheets, filename){
  const wb = XLSX.utils.book_new();
  for (const name of sheetNames){
    const info = sheets[name];
    const ws = XLSX.utils.aoa_to_sheet(info.data);
    const formulas = info.formulas || {};
    for (const [rStr, cols] of Object.entries(formulas)){
      for (const [cStr, f] of Object.entries(cols)){
        const addr = XLSX.utils.encode_cell({ r: parseInt(rStr), c: parseInt(cStr) });
        if (!ws[addr]) ws[addr] = { t:'n' };
        ws[addr].f = f;
      }
    }
    XLSX.utils.book_append_sheet(wb, ws, name.substring(0, 31));
  }
  const buf = XLSX.write(wb, { type:'array', bookType:'xlsx' });
  downloadBlob(new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
}

// ========== 列类型识别 ==========
function detectColumnType(header, values){
  const headerLower = (header || '').toString().toLowerCase();
  const headerStr = header || '';
  const KW = {
    name: ['姓名','名字','联系人','负责人','经办人','员工姓名','name'],
    dept: ['部门','团队','组织','区域','分公司','事业部','dept','department'],
    idcard: ['身份证','证件号','身份号','id card','identity'],
    bank: ['银行卡','卡号','账号','bank'],
    phone: ['手机','电话','联系方式','phone','mobile','tel'],
    salary: ['薪资','工资','薪酬','底薪','基本工资','月薪','年薪','salary','wage','pay'],
    amount: ['金额','奖金','提成','补贴','津贴','社保','公积金','扣款','实发','应发','税','amount','bonus'],
    perf: ['业绩','营收','收入','产值','销售额','利润','产量','指标','performance','revenue','sales'],
    date: ['日期','时间','入职','生日','出生','离职','date','time'],
    seq: ['序号','编号','行号','#','no','seq','index'],
    job: ['职级','职位','岗位','职称','level','rank','title','position'],
    empid: ['工号','员工编号','人员编号','emp id','employee id','staff id'],
  };
  const mk = ks => ks.some(k => headerStr.includes(k) || headerLower.includes(k.toLowerCase()));
  if (mk(KW.name)) return { type:'name', label:'人名', strategy:'mapping', defaultPrefix:'P' };
  if (mk(KW.idcard)) return { type:'idcard', label:'证件号', strategy:'format-replace', defaultPrefix:'' };
  if (mk(KW.bank)) return { type:'bankcard', label:'银行卡', strategy:'format-replace', defaultPrefix:'' };
  if (mk(KW.phone)) return { type:'phone', label:'手机号', strategy:'format-replace', defaultPrefix:'' };
  if (mk(KW.empid)) return { type:'empid', label:'员工编号', strategy:'mapping', defaultPrefix:'E' };
  if (mk(KW.dept)) return { type:'department', label:'组织/部门', strategy:'mapping', defaultPrefix:'D' };
  if (mk(KW.salary)) return { type:'salary', label:'薪资金额', strategy:'scale', defaultPrefix:'' };
  if (mk(KW.amount)) return { type:'amount', label:'金额', strategy:'scale', defaultPrefix:'' };
  if (mk(KW.perf)) return { type:'performance', label:'业绩数据', strategy:'scale', defaultPrefix:'' };
  if (mk(KW.date)) return { type:'date', label:'日期', strategy:'date-offset', defaultPrefix:'' };
  if (mk(KW.seq)) return { type:'sequence', label:'序号', strategy:'none', defaultPrefix:'' };
  if (mk(KW.job)) return { type:'category', label:'分类/枚举', strategy:'mapping', defaultPrefix:'L' };

  const sv = values.filter(v => v !== '' && v !== null && v !== undefined).slice(0, 50);
  if (sv.length === 0) return { type:'unknown', label:'未知', strategy:'none', defaultPrefix:'' };
  if (sv.every(v => /^\d{17}[\dXx]$/.test(String(v).trim()))) return { type:'idcard', label:'证件号', strategy:'format-replace', defaultPrefix:'' };
  if (sv.every(v => /^1[3-9]\d{9}$/.test(String(v).trim()))) return { type:'phone', label:'手机号', strategy:'format-replace', defaultPrefix:'' };
  const ymP = /^\d{4}[\-\/年]\s?\d{1,2}[月]?$/;
  if (sv.filter(v => ymP.test(String(v).trim())).length > sv.length * 0.8) return { type:'date', label:'年月', strategy:'date-offset', defaultPrefix:'' };
  const dc = sv.filter(v => { const d = new Date(v); return !isNaN(d.getTime()) && String(v).match(/[\-\/年月日]/); }).length;
  if (dc > sv.length * 0.8) return { type:'date', label:'日期', strategy:'date-offset', defaultPrefix:'' };
  const numVals = sv.filter(v => !isNaN(Number(v)) && String(v).trim() !== '');
  if (numVals.length > sv.length * 0.8){
    const nums = numVals.map(Number);
    const min = Math.min(...nums), max = Math.max(...nums);
    if (min >= 0 && max <= sv.length * 2 && nums.every(n => Number.isInteger(n))) return { type:'sequence', label:'序号', strategy:'none', defaultPrefix:'' };
    if (max > 1000) return { type:'amount', label:'数值(可能是金额)', strategy:'scale', defaultPrefix:'' };
    return { type:'number', label:'数值', strategy:'scale', defaultPrefix:'' };
  }
  const avgLen = sv.reduce((s, v) => s + String(v).length, 0) / sv.length;
  const hasCn = sv.some(v => /[\u4e00-\u9fff]/.test(String(v)));
  if (hasCn && avgLen <= 4) return { type:'name', label:'文本(可能是人名)', strategy:'mapping', defaultPrefix:'P' };
  if (new Set(sv.map(String)).size < sv.length * 0.3) return { type:'category', label:'分类/枚举', strategy:'mapping', defaultPrefix:'C' };
  return { type:'text', label:'文本', strategy:'mapping', defaultPrefix:'T' };
}
function detectAllColumns(sheetData){
  const data = sheetData.data;
  if (!data || data.length < 2) return [];
  const headers = data[0];
  const results = [];
  for (let col = 0; col < headers.length; col++){
    const values = [];
    for (let row = 1; row < data.length; row++){
      if (data[row][col] !== undefined && data[row][col] !== '') values.push(data[row][col]);
    }
    const header = headers[col];
    const headerEmpty = (header === undefined || header === null || String(header).trim() === '');
    if (values.length === 0 && headerEmpty) continue;
    if (values.length === 0){ results.push({ colIndex:col, header, sample:[], type:'empty', label:'空列', strategy:'none', defaultPrefix:'' }); continue; }
    const d = detectColumnType(header, values);
    results.push({ colIndex:col, header: header || `列${col+1}`, sample: values.slice(0, 3), ...d });
  }
  const pc = {}, pi = {};
  for (const r of results){ if (r.defaultPrefix && r.strategy === 'mapping') pc[r.defaultPrefix] = (pc[r.defaultPrefix] || 0) + 1; }
  for (const r of results){
    if (r.defaultPrefix && r.strategy === 'mapping' && pc[r.defaultPrefix] > 1){
      pi[r.defaultPrefix] = (pi[r.defaultPrefix] || 0) + 1;
      if (pi[r.defaultPrefix] > 1) r.defaultPrefix = r.defaultPrefix + pi[r.defaultPrefix];
    }
  }
  return results;
}

// ========== 随机假数据 ==========
function randomChineseName(){
  const s = ['王','李','张','刘','陈','杨','赵','黄','周','吴','徐','孙','胡','朱','高','林','何','郭','马','罗'];
  const c = ['伟','芳','娜','秀英','敏','静','丽','强','磊','军','洋','勇','艳','杰','娟','涛','明','超','秀兰','霞'];
  return s[Math.floor(Math.random()*s.length)] + c[Math.floor(Math.random()*c.length)] + (Math.random() > 0.5 ? c[Math.floor(Math.random()*c.length)] : '');
}
function randomIdCard(){
  const areas = ['110101','310101','440103','330102','320102','510104','420102','610103'];
  const a = areas[Math.floor(Math.random()*areas.length)];
  const y = 1960 + Math.floor(Math.random()*40);
  const m = String(Math.floor(Math.random()*12)+1).padStart(2,'0');
  const d = String(Math.floor(Math.random()*28)+1).padStart(2,'0');
  const s = String(Math.floor(Math.random()*999)+1).padStart(3,'0');
  const base = `${a}${y}${m}${d}${s}`;
  const w = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2];
  const cc = '10X98765432';
  let sum = 0;
  for (let i = 0; i < 17; i++) sum += parseInt(base[i]) * w[i];
  return base + cc[sum % 11];
}
function randomBankCard(){
  const p = ['6222','6228','6214','6217','6225'];
  let card = p[Math.floor(Math.random()*p.length)];
  while (card.length < 19) card += Math.floor(Math.random()*10);
  return card;
}
function randomPhone(){
  const p = ['138','139','136','137','158','159','188','189','135','186'];
  let ph = p[Math.floor(Math.random()*p.length)];
  while (ph.length < 11) ph += Math.floor(Math.random()*10);
  return ph;
}

// ========== 脱敏核心引擎 ==========
const STRATEGY_NAMES = { mapping:'编号替换', fakename:'随机假名', 'format-replace':'格式保留替换', scale:'等比缩放', 'scale-noise':'等比缩放+扰动', offset:'固定偏移', 'range-map':'区间映射', 'date-offset':'日期偏移' };

function doDesensitizeCore(sheetNames, sheets, configs, baseKeyData){
  const keyData = { version:'1.0', timestamp:new Date().toISOString(), sheets:{} };
  if (baseKeyData) keyData.basedOn = baseKeyData.timestamp || 'unknown';
  const resultSheets = {};

  for (const sheetName of sheetNames){
    const si = sheets[sheetName];
    const data = si.data;
    const formulas = si.formulas || {};
    const config = configs[sheetName] || [];
    if (!data || data.length < 2){ resultSheets[sheetName] = { data:[...data], formulas }; continue; }

    const resultData = data.map(r => [...r]);
    const sheetKey = { columns:{} };

    for (const cc of config){
      const { colIndex, strategy, params } = cc;
      if (strategy === 'none') continue;
      const hasNonEmpty = data.slice(1).some(r => { const v = r[colIndex]; return v !== undefined && v !== null && v !== ''; });
      if (!hasNonEmpty) continue;

      const headerName = data[0] ? (data[0][colIndex] || `列${colIndex+1}`) : `列${colIndex+1}`;
      const colKey = { strategy, header:String(headerName), params:{...params}, mappings:{} };
      const baseColKey = baseKeyData && baseKeyData.sheets && baseKeyData.sheets[sheetName] && baseKeyData.sheets[sheetName].columns && baseKeyData.sheets[sheetName].columns[colIndex];

      if (strategy === 'mapping'){
        const prefix = params && params.prefix || 'A';
        const uv = new Map();
        let counter = 1;
        if (baseColKey && baseColKey.mappings && ['mapping','fakename','format-replace'].includes(baseColKey.strategy)){
          for (const [o, m] of Object.entries(baseColKey.mappings)){
            uv.set(o, m);
            const match = String(m).match(/(\d+)$/);
            if (match){ const n = parseInt(match[1]); if (n >= counter) counter = n + 1; }
          }
        }
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (val && !uv.has(val)){ uv.set(val, `${prefix}${String(counter).padStart(3,'0')}`); counter++; }
        }
        colKey.mappings = Object.fromEntries(uv);
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (uv.has(val)) resultData[row][colIndex] = uv.get(val);
        }
      } else if (strategy === 'fakename'){
        const uv = new Map();
        if (baseColKey && baseColKey.mappings) for (const [o,m] of Object.entries(baseColKey.mappings)) uv.set(o, m);
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (val && !uv.has(val)) uv.set(val, randomChineseName());
        }
        colKey.mappings = Object.fromEntries(uv);
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (uv.has(val)) resultData[row][colIndex] = uv.get(val);
        }
      } else if (strategy === 'scale'){
        const factor = (params && params.factor) || (baseColKey && baseColKey.params && baseColKey.params.factor) || (0.5 + Math.random() * 0.8);
        colKey.params.factor = factor;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = Number(data[row][colIndex]);
          if (!isNaN(val) && data[row][colIndex] !== '') resultData[row][colIndex] = Math.round(val * factor * 100) / 100;
        }
      } else if (strategy === 'scale-noise'){
        const factor = (params && params.factor) || (baseColKey && baseColKey.params && baseColKey.params.factor) || (0.5 + Math.random() * 0.8);
        const np = (params && params.noisePercent) || (baseColKey && baseColKey.params && baseColKey.params.noisePercent) || 2;
        colKey.params.factor = factor; colKey.params.noisePercent = np;
        const noises = {};
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = Number(data[row][colIndex]);
          if (!isNaN(val) && data[row][colIndex] !== ''){
            const noise = val * (np / 100) * (Math.random() * 2 - 1);
            noises[row] = noise;
            resultData[row][colIndex] = Math.round((val * factor + noise) * 100) / 100;
          }
        }
        colKey.noises = noises;
      } else if (strategy === 'offset'){
        const off = (params && params.offset) || (baseColKey && baseColKey.params && baseColKey.params.offset) || Math.floor(Math.random() * 10000 - 5000);
        colKey.params.offset = off;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = Number(data[row][colIndex]);
          if (!isNaN(val) && data[row][colIndex] !== '') resultData[row][colIndex] = Math.round((val + off) * 100) / 100;
        }
      } else if (strategy === 'range-map'){
        const nums = [];
        for (let row = 1; row < data.length; row++){ const v = Number(data[row][colIndex]); if (!isNaN(v) && data[row][colIndex] !== '') nums.push(v); }
        const origMin = Math.min(...nums), origMax = Math.max(...nums);
        const newMin = (params && params.newMin) || 1000, newMax = (params && params.newMax) || 10000;
        colKey.params = { origMin, origMax, newMin, newMax };
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = Number(data[row][colIndex]);
          if (!isNaN(val) && data[row][colIndex] !== ''){
            const mapped = origMax === origMin ? (newMin + newMax) / 2 : newMin + (val - origMin) / (origMax - origMin) * (newMax - newMin);
            resultData[row][colIndex] = Math.round(mapped * 100) / 100;
          }
        }
      } else if (strategy === 'format-replace'){
        const colType = cc.type;
        const uv = new Map();
        if (baseColKey && baseColKey.mappings) for (const [o,m] of Object.entries(baseColKey.mappings)) uv.set(o, m);
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (val && !uv.has(val)){
            let fake;
            if (colType === 'idcard') fake = randomIdCard();
            else if (colType === 'bankcard') fake = randomBankCard();
            else if (colType === 'phone') fake = randomPhone();
            else fake = `FAKE-${String(uv.size + 1).padStart(4,'0')}`;
            uv.set(val, fake);
          }
        }
        colKey.mappings = Object.fromEntries(uv);
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = String(data[row][colIndex] || '');
          if (uv.has(val)) resultData[row][colIndex] = uv.get(val);
        }
      } else if (strategy === 'date-offset'){
        const od = (params && params.offsetDays) || (baseColKey && baseColKey.params && baseColKey.params.offsetDays) || Math.floor(Math.random() * 365 - 180);
        colKey.params.offsetDays = od;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][colIndex]) continue;
          const val = data[row][colIndex];
          if (!val) continue;
          try { const d = new Date(val); if (!isNaN(d.getTime())){ d.setDate(d.getDate() + od); resultData[row][colIndex] = d.toISOString().split('T')[0]; } } catch(e){}
        }
      }
      sheetKey.columns[colIndex] = colKey;
    }
    keyData.sheets[sheetName] = sheetKey;
    resultSheets[sheetName] = { data:resultData, formulas };
  }

  // 报告
  const report = { sheets:[] };
  let tp = 0, ts = 0;
  for (const sn of sheetNames){
    const si = sheets[sn];
    const data = si.data;
    const formulas = si.formulas || {};
    const sk = keyData.sheets[sn];
    if (!sk) continue;
    const sr = { name:sn, columns:[], totalRows:data.length - 1 };
    for (const [ciStr, ck] of Object.entries(sk.columns)){
      if (ck.strategy === 'none') continue;
      const ci = parseInt(ciStr);
      let p = 0, s = 0;
      for (let row = 1; row < data.length; row++){
        if (formulas[row] && formulas[row][ci]) s++;
        else { const v = data[row][ci]; if (v !== undefined && v !== null && v !== '') p++; }
      }
      sr.columns.push({ header:ck.header, strategy:STRATEGY_NAMES[ck.strategy] || ck.strategy, processed:p, skippedFormula:s });
      tp += p; ts += s;
    }
    if (sr.columns.length > 0) report.sheets.push(sr);
  }
  report.totalProcessed = tp; report.totalSkippedFormula = ts;
  return { keyData, resultSheets, report };
}

// ========== 还原核心 ==========
function doRestoreCore(sheetNames, sheets, keyData, skipHeaderCheck){
  if (!skipHeaderCheck){
    const mm = [];
    for (const sn of sheetNames){
      const si = sheets[sn];
      const data = si.data;
      const sk = keyData.sheets[sn];
      if (!sk || !data || data.length < 2) continue;
      const headers = data[0] || [];
      for (const [ciStr, ck] of Object.entries(sk.columns)){
        const ci = parseInt(ciStr);
        if (ck.strategy === 'none') continue;
        const cur = String(headers[ci] || '').trim();
        const keyH = String(ck.header || '').trim();
        if (keyH && cur && keyH !== cur) mm.push({ sheet:sn, colIndex:ci, keyHeader:keyH, currentHeader:cur });
      }
    }
    if (mm.length > 0) return { headerMismatches:mm };
  }

  const resultSheets = {};
  for (const sn of sheetNames){
    const si = sheets[sn];
    const data = si.data;
    const formulas = si.formulas || {};
    const sk = keyData.sheets[sn];
    if (!sk || !data || data.length < 2){ resultSheets[sn] = { data:[...data], formulas }; continue; }
    const resultData = data.map(r => [...r]);

    for (const [ciStr, ck] of Object.entries(sk.columns)){
      const ci = parseInt(ciStr);
      const { strategy, params, mappings, noises } = ck;
      if (strategy === 'none') continue;

      if (['mapping','fakename','format-replace'].includes(strategy)){
        const rev = {};
        for (const [o,m] of Object.entries(mappings)) rev[m] = o;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = String(data[row][ci] || '');
          if (rev[val] !== undefined) resultData[row][ci] = rev[val];
        }
      } else if (strategy === 'scale'){
        const f = params.factor;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = Number(data[row][ci]);
          if (!isNaN(val) && data[row][ci] !== '') resultData[row][ci] = Math.round(val / f * 100) / 100;
        }
      } else if (strategy === 'scale-noise'){
        const f = params.factor;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = Number(data[row][ci]);
          if (!isNaN(val) && data[row][ci] !== ''){ const n = (noises && noises[row]) ? noises[row] : 0; resultData[row][ci] = Math.round((val - n) / f * 100) / 100; }
        }
      } else if (strategy === 'offset'){
        const off = params.offset;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = Number(data[row][ci]);
          if (!isNaN(val) && data[row][ci] !== '') resultData[row][ci] = Math.round((val - off) * 100) / 100;
        }
      } else if (strategy === 'range-map'){
        const { origMin, origMax, newMin, newMax } = params;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = Number(data[row][ci]);
          if (!isNaN(val) && data[row][ci] !== ''){
            const r = newMax === newMin ? origMin : origMin + (val - newMin) / (newMax - newMin) * (origMax - origMin);
            resultData[row][ci] = Math.round(r * 100) / 100;
          }
        }
      } else if (strategy === 'date-offset'){
        const od = params.offsetDays;
        for (let row = 1; row < data.length; row++){
          if (formulas[row] && formulas[row][ci]) continue;
          const val = data[row][ci]; if (!val) continue;
          try { const d = new Date(val); if (!isNaN(d.getTime())){ d.setDate(d.getDate() - od); resultData[row][ci] = d.toISOString().split('T')[0]; } } catch(e){}
        }
      }
    }
    resultSheets[sn] = { data:resultData, formulas };
  }

  const report = { sheets:[] };
  let tr = 0, ts = 0;
  for (const sn of sheetNames){
    const si = sheets[sn];
    const data = si.data;
    const formulas = si.formulas || {};
    const sk = keyData.sheets[sn];
    if (!sk) continue;
    const sr = { name:sn, columns:[], totalRows:data.length - 1 };
    for (const [ciStr, ck] of Object.entries(sk.columns)){
      if (ck.strategy === 'none') continue;
      const ci = parseInt(ciStr);
      let r = 0, s = 0;
      for (let row = 1; row < data.length; row++){
        if (formulas[row] && formulas[row][ci]) s++;
        else { const a = String(data[row][ci] || ''); const b = String(resultSheets[sn].data[row][ci] || ''); if (a !== b) r++; }
      }
      sr.columns.push({ header:ck.header, strategy:STRATEGY_NAMES[ck.strategy] || ck.strategy, restored:r, skippedFormula:s });
      tr += r; ts += s;
    }
    if (sr.columns.length > 0) report.sheets.push(sr);
  }
  report.totalRestored = tr; report.totalSkippedFormula = ts;
  return { resultSheets, report };
}

// ========== 参数默认值 & localStorage 偏好 ==========
function getDefaultParams(strategy, defaultPrefix){
  switch (strategy){
    case 'scale': return { factor:'' };
    case 'scale-noise': return { factor:'', noisePercent:2 };
    case 'offset': return { offset:'' };
    case 'range-map': return { newMin:1000, newMax:10000 };
    case 'mapping': return { prefix: defaultPrefix || 'A' };
    case 'fakename': return {};
    case 'format-replace': return {};
    case 'date-offset': return { offsetDays:'' };
    default: return {};
  }
}
function loadPrefs(){
  try { return JSON.parse(localStorage.getItem(LS_KEY_PREF) || '{}'); } catch(e){ return {}; }
}
function savePrefs(prefs){
  try { localStorage.setItem(LS_KEY_PREF, JSON.stringify(prefs)); } catch(e){}
}
function applyLocalPreferences(){
  const prefs = loadPrefs();
  for (const sheetName of Object.keys(dState.columnConfigs)){
    const configs = dState.columnConfigs[sheetName];
    for (const c of configs){
      const key = String(c.header || '').trim();
      if (!key) continue;
      const saved = prefs[key];
      if (saved && saved.strategy){
        c.strategy = saved.strategy;
        c.params = saved.params ? { ...saved.params } : getDefaultParams(saved.strategy, c.defaultPrefix);
      }
    }
  }
}
function recordPreferences(configs){
  const prefs = loadPrefs();
  for (const sheetName of Object.keys(configs)){
    for (const c of configs[sheetName]){
      const key = String(c.header || '').trim();
      if (!key) continue;
      // 只记录策略 + 必要参数（不保存具体随机值，避免跨任务污染）
      prefs[key] = { strategy: c.strategy, params: { prefix: c.params?.prefix } };
    }
  }
  savePrefs(prefs);
}

// ========== 脱敏 UI ==========
function initDesensitizeUI(){
  const dz = document.getElementById('d-dropzone');
  dz.addEventListener('click', () => document.getElementById('d-file-input').click());
  dz.addEventListener('dragover', e => { e.preventDefault(); e.stopPropagation(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', async (e) => {
    e.preventDefault(); e.stopPropagation();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) await handleDExcel(e.dataTransfer.files[0]);
  });
  document.getElementById('d-file-input').addEventListener('change', async (e) => {
    if (e.target.files.length === 0) return;
    await handleDExcel(e.target.files[0]);
    e.target.value = '';
  });
  document.getElementById('d-base-key-btn').addEventListener('click', () => document.getElementById('d-base-key-input').click());
  document.getElementById('d-base-key-clear').addEventListener('click', () => { dState.baseKeyData = null; updateBaseKeyUI(); });
  document.getElementById('d-base-key-input').addEventListener('change', async (e) => {
    if (e.target.files.length === 0) return;
    try {
      const text = await e.target.files[0].text();
      dState.baseKeyData = JSON.parse(text);
      updateBaseKeyUI();
    } catch (err) { alert('密钥文件解析失败：' + err.message); }
    e.target.value = '';
  });
}

async function handleDExcel(file){
  showLoading('正在读取 Excel 文件...');
  try {
    const result = await readExcelFile(file);
    dState.fileName = file.name;
    dState.workbook = result;
    dState.currentSheet = result.sheetNames[0];
    dState.columnConfigs = {};
    for (const name of result.sheetNames){
      const cols = detectAllColumns(result.sheets[name]);
      dState.columnConfigs[name] = cols.map(c => ({ ...c, strategy:c.strategy, params:getDefaultParams(c.strategy, c.defaultPrefix) }));
    }
    applyLocalPreferences();
    hideLoading();
    updateDStep(2);
  } catch(e){ hideLoading(); alert('读取文件失败：' + e.message); }
}

function updateBaseKeyUI(){
  const bar = document.getElementById('d-base-key-bar');
  const hint = document.getElementById('d-base-key-hint');
  const btn = document.getElementById('d-base-key-btn');
  const clr = document.getElementById('d-base-key-clear');
  if (dState.baseKeyData){
    bar.classList.add('loaded');
    const ts = dState.baseKeyData.timestamp ? new Date(dState.baseKeyData.timestamp).toLocaleString() : '未知';
    const sc = Object.keys(dState.baseKeyData.sheets || {}).length;
    let cc = 0, mc = 0;
    for (const s of Object.values(dState.baseKeyData.sheets || {})){
      for (const col of Object.values(s.columns || {})){ cc++; mc += Object.keys(col.mappings || {}).length; }
    }
    hint.textContent = `✅ 已载入 · ${ts} · ${sc}个工作表 · ${cc}列 · ${mc}条映射`;
    btn.classList.add('hidden'); clr.classList.remove('hidden');
  } else {
    bar.classList.remove('loaded');
    hint.textContent = '导入已有密钥，复用映射关系和参数，适合长期维护的表格';
    btn.classList.remove('hidden'); clr.classList.add('hidden');
  }
}

function updateDStep(step){
  dState.step = step;
  for (let i = 1; i <= 3; i++){
    const el = document.getElementById(`d-step-${i}`);
    el.classList.remove('active', 'done');
    if (i < step) el.classList.add('done'); else if (i === step) el.classList.add('active');
  }
  document.getElementById('d-import').classList.toggle('hidden', step !== 1);
  document.getElementById('d-config').classList.toggle('hidden', step !== 2);
  document.getElementById('d-result').classList.toggle('hidden', step !== 3);

  const btnBack = document.getElementById('d-btn-back');
  const btnNext = document.getElementById('d-btn-next');
  btnBack.onclick = () => { if (dState.step === 2) updateDStep(1); else if (dState.step === 3) updateDStep(2); };

  if (step === 1){
    btnBack.classList.add('hidden');
    btnNext.textContent = '开始配置 →';
    btnNext.disabled = !dState.workbook;
    btnNext.className = 'btn btn-primary';
    btnNext.onclick = () => updateDStep(2);
  } else if (step === 2){
    btnBack.classList.remove('hidden');
    btnNext.textContent = '🔐 执行脱敏';
    btnNext.disabled = false;
    btnNext.className = 'btn btn-primary';
    btnNext.onclick = doDesensitizeAction;
    renderDConfig();
  } else if (step === 3){
    btnBack.classList.remove('hidden');
    btnNext.textContent = '📥 下载脱敏文件和密钥';
    btnNext.disabled = false;
    btnNext.className = 'btn btn-success';
    btnNext.onclick = doExportDesensitized;
  }
}

function renderDConfig(){
  const wb = dState.workbook;
  const fi = document.getElementById('d-file-info');
  fi.innerHTML = `<span class="file-icon">📊</span><div><div class="file-name">${escapeHtml(dState.fileName)}</div><div class="file-detail">${wb.sheetNames.length} 个工作表</div></div>`;
  const cb = document.createElement('button');
  cb.className = 'change-btn'; cb.textContent = '更换文件';
  cb.onclick = () => document.getElementById('d-file-input').click();
  fi.appendChild(cb);

  renderSheetTabs('d-sheet-tabs', wb.sheetNames, dState.currentSheet, (n) => { dState.currentSheet = n; renderDConfig(); });
  renderDataPreview('d-preview-table', wb.sheets[dState.currentSheet].data, 5);
  renderColumnConfig();
}

function renderSheetTabs(cid, names, cur, onClick){
  const c = document.getElementById(cid);
  c.innerHTML = '';
  names.forEach(name => {
    const btn = document.createElement('button');
    btn.className = 'sheet-tab' + (name === cur ? ' active' : '');
    btn.textContent = name;
    btn.onclick = () => onClick(name);
    c.appendChild(btn);
  });
}

function renderDataPreview(tableId, data, maxRows){
  const table = document.getElementById(tableId);
  if (!data || data.length === 0){ table.innerHTML = '<tr><td>无数据</td></tr>'; return; }
  const headers = data[0];
  let html = '<thead><tr>' + headers.map(h => `<th>${escapeHtml(String(h || '-'))}</th>`).join('') + '</tr></thead><tbody>';
  const rows = Math.min(data.length - 1, maxRows);
  for (let i = 1; i <= rows; i++){
    html += '<tr>' + headers.map((_, j) => `<td>${escapeHtml(String(data[i][j] !== undefined ? data[i][j] : ''))}</td>`).join('') + '</tr>';
  }
  html += '</tbody>';
  table.innerHTML = html;
}

function escapeHtml(s){ return String(s).replace(/[&<>"']/g, m => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;' }[m])); }

function getTypeBadgeClass(strategy){
  if (strategy === 'none') return 'none';
  if (['mapping','fakename'].includes(strategy)) return 'name';
  if (['scale','scale-noise','offset','range-map'].includes(strategy)) return 'number';
  if (strategy === 'format-replace') return 'id';
  if (strategy === 'date-offset') return 'date';
  return 'none';
}

function renderColumnConfig(){
  const table = document.getElementById('d-config-table');
  const configs = dState.columnConfigs[dState.currentSheet] || [];
  let html = `<thead><tr><th style="width:30px;">#</th><th>列名</th><th>样本数据</th><th>识别类型</th><th>脱敏策略</th><th>参数</th><th>预览效果</th><th style="width:50px;text-align:center;">操作</th></tr></thead><tbody>`;
  for (let i = 0; i < configs.length; i++){
    const c = configs[i];
    html += `<tr>
      <td style="color:var(--text-dim)">${i + 1}</td>
      <td style="font-weight:500;">${escapeHtml(String(c.header))}</td>
      <td><span class="sample-data">${escapeHtml(c.sample.slice(0, 2).join(', ') || '-')}</span></td>
      <td><span class="type-badge ${getTypeBadgeClass(c.strategy)}">${escapeHtml(c.label)}</span></td>
      <td>${renderStrategySelect(i, c)}</td>
      <td>${renderParams(i, c)}</td>
      <td>${renderPreview(c)}</td>
      <td style="text-align:center;"><button class="remove-col-btn" data-idx="${i}" title="移除此列">✕</button></td>
    </tr>`;
  }
  html += '</tbody>';
  table.innerHTML = html;
  // 绑定事件
  table.querySelectorAll('select[data-idx]').forEach(sel => {
    sel.onchange = () => onStrategyChange(parseInt(sel.dataset.idx), sel.value);
  });
  table.querySelectorAll('input[data-idx]').forEach(inp => {
    inp.onchange = () => onParamChange(parseInt(inp.dataset.idx), inp.dataset.key, inp.value);
  });
  table.querySelectorAll('.remove-col-btn').forEach(btn => {
    btn.onclick = () => removeColumn(parseInt(btn.dataset.idx));
  });
}

function renderStrategySelect(idx, config){
  const options = [
    { value:'none', label:'不处理' },
    { value:'mapping', label:'编号替换' },
    { value:'fakename', label:'随机假名' },
    { value:'format-replace', label:'格式保留替换' },
    { value:'scale', label:'等比缩放' },
    { value:'scale-noise', label:'等比缩放+扰动' },
    { value:'offset', label:'固定偏移' },
    { value:'range-map', label:'区间映射' },
    { value:'date-offset', label:'日期偏移' },
  ];
  return `<select data-idx="${idx}">${options.map(o => `<option value="${o.value}"${o.value === config.strategy ? ' selected' : ''}>${o.label}</option>`).join('')}</select>`;
}

function renderParams(idx, config){
  const s = config.strategy;
  const p = config.params || {};
  if (s === 'mapping') return `<span class="param-input"><label>前缀</label><input class="prefix-input" value="${escapeHtml(p.prefix || 'A')}" data-idx="${idx}" data-key="prefix" maxlength="5"></span>`;
  if (s === 'scale') return `<span class="param-input"><label>系数</label><input value="${escapeHtml(p.factor || '')}" placeholder="随机" data-idx="${idx}" data-key="factor" title="留空则自动随机生成"></span>`;
  if (s === 'scale-noise') return `<span class="param-input"><label>系数</label><input value="${escapeHtml(p.factor || '')}" placeholder="随机" data-idx="${idx}" data-key="factor" style="width:55px;"><label>扰动</label><input value="${escapeHtml(p.noisePercent || 2)}" data-idx="${idx}" data-key="noisePercent" style="width:40px;">%</span>`;
  if (s === 'offset') return `<span class="param-input"><label>偏移</label><input value="${escapeHtml(p.offset || '')}" placeholder="随机" data-idx="${idx}" data-key="offset"></span>`;
  if (s === 'range-map') return `<span class="param-input"><input value="${escapeHtml(p.newMin || 1000)}" data-idx="${idx}" data-key="newMin" style="width:55px;"> ~ <input value="${escapeHtml(p.newMax || 10000)}" data-idx="${idx}" data-key="newMax" style="width:55px;"></span>`;
  if (s === 'date-offset') return `<span class="param-input"><label>偏移</label><input value="${escapeHtml(p.offsetDays || '')}" placeholder="随机" data-idx="${idx}" data-key="offsetDays"><label>天</label></span>`;
  return '<span style="color:var(--text-dim);font-size:12px;">-</span>';
}

function renderPreview(config){
  if (config.strategy === 'none' || !config.sample || config.sample.length === 0) return '<span style="color:var(--text-dim);font-size:12px;">不处理</span>';
  const orig = String(config.sample[0]);
  let masked = '...';
  if (config.strategy === 'mapping') masked = `${(config.params && config.params.prefix) || 'A'}001`;
  else if (config.strategy === 'fakename') masked = '王伟芳';
  else if (config.strategy === 'scale'){ const val = Number(orig); const factor = Number(config.params && config.params.factor) || 0.73; masked = isNaN(val) ? '?' : String(Math.round(val * factor * 100) / 100); }
  else if (config.strategy === 'scale-noise'){ const val = Number(orig); const factor = Number(config.params && config.params.factor) || 0.73; masked = isNaN(val) ? '?' : '≈' + String(Math.round(val * factor * 100) / 100); }
  else if (config.strategy === 'offset'){ const val = Number(orig); const off = Number(config.params && config.params.offset) || 5000; masked = isNaN(val) ? '?' : String(Math.round((val + off) * 100) / 100); }
  else if (config.strategy === 'range-map') masked = `${(config.params && config.params.newMin) || 1000}~${(config.params && config.params.newMax) || 10000}`;
  else if (config.strategy === 'format-replace') masked = '随机替换';
  else if (config.strategy === 'date-offset') masked = '±偏移';
  return `<span class="preview-cell"><span class="orig">${escapeHtml(orig.length > 10 ? orig.slice(0, 10) + '…' : orig)}</span><span class="arrow">→</span><span class="masked">${escapeHtml(masked)}</span></span>`;
}

function onStrategyChange(idx, value){
  const configs = dState.columnConfigs[dState.currentSheet];
  configs[idx].strategy = value;
  configs[idx].params = getDefaultParams(value, configs[idx].defaultPrefix);
  renderColumnConfig();
}
function onParamChange(idx, key, value){
  const configs = dState.columnConfigs[dState.currentSheet];
  if (!configs[idx].params) configs[idx].params = {};
  configs[idx].params[key] = value;
  renderColumnConfig();
}
function removeColumn(idx){
  const configs = dState.columnConfigs[dState.currentSheet];
  if (!configs || idx < 0 || idx >= configs.length) return;
  configs.splice(idx, 1);
  renderColumnConfig();
}

// ========== 执行脱敏 ==========
async function doDesensitizeAction(){
  showLoading('正在执行脱敏...');
  await new Promise(r => setTimeout(r, 50));  // 让 loading 显示
  try {
    const configs = {};
    for (const [sn, cs] of Object.entries(dState.columnConfigs)){
      configs[sn] = cs.map(c => ({
        colIndex: c.colIndex,
        type: c.type,
        strategy: c.strategy,
        params: {
          ...c.params,
          factor: (c.params && c.params.factor) ? Number(c.params.factor) : undefined,
          offset: (c.params && c.params.offset) ? Number(c.params.offset) : undefined,
          offsetDays: (c.params && c.params.offsetDays) ? Number(c.params.offsetDays) : undefined,
          noisePercent: (c.params && c.params.noisePercent) ? Number(c.params.noisePercent) : undefined,
          newMin: (c.params && c.params.newMin) ? Number(c.params.newMin) : undefined,
          newMax: (c.params && c.params.newMax) ? Number(c.params.newMax) : undefined,
        },
      }));
    }
    const result = doDesensitizeCore(dState.workbook.sheetNames, dState.workbook.sheets, configs, dState.baseKeyData);
    dState.desensitizedResult = result.resultSheets;
    dState.keyData = result.keyData;
    recordPreferences(dState.columnConfigs);
    renderComparePreview();
    if (result.report) renderDesensitizeReport(result.report);
    hideLoading();
    updateDStep(3);
  } catch(e){ hideLoading(); alert('脱敏失败：' + e.message); }
}

function renderComparePreview(){
  const table = document.getElementById('d-compare-table');
  const sn = dState.currentSheet;
  const origData = dState.workbook.sheets[sn].data;
  const maskedData = dState.desensitizedResult[sn].data;
  const headers = origData[0];
  let html = '<thead><tr><th>#</th>';
  for (const h of headers) html += `<th class="col-orig">${escapeHtml(String(h))}(原)</th><th class="col-arrow"></th><th class="col-masked">${escapeHtml(String(h))}(脱敏)</th>`;
  html += '</tr></thead><tbody>';
  const rows = Math.min(origData.length - 1, 8);
  for (let i = 1; i <= rows; i++){
    html += `<tr><td style="color:var(--text-dim)">${i}</td>`;
    for (let j = 0; j < headers.length; j++){
      const orig = origData[i][j] !== undefined ? origData[i][j] : '';
      const masked = maskedData[i][j] !== undefined ? maskedData[i][j] : '';
      const changed = String(orig) !== String(masked);
      html += `<td class="col-orig">${escapeHtml(String(orig))}</td><td class="col-arrow">${changed ? '→' : ''}</td><td class="col-masked" style="${changed ? 'color:var(--success);font-weight:500;' : ''}">${escapeHtml(String(masked))}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';
  table.innerHTML = html;
}

function doExportDesensitized(){
  try {
    const baseName = dState.fileName.replace(/\.[^.]+$/, '');
    const excelName = `${baseName}_脱敏.xlsx`;
    exportExcelFile(dState.workbook.sheetNames, dState.desensitizedResult, excelName);
    document.getElementById('d-result-file').textContent = excelName + '（已下载到浏览器默认下载目录）';

    const keyName = `脱敏密钥_${new Date().toISOString().slice(0, 10)}.json`;
    const keyBlob = new Blob([JSON.stringify(dState.keyData, null, 2)], { type:'application/json' });
    downloadBlob(keyBlob, keyName);
    document.getElementById('d-result-key').textContent = keyName + '（已下载）';

    // 映射表 Excel
    const mapName = `脱敏映射表_${new Date().toISOString().slice(0, 10)}.xlsx`;
    exportKeyAsExcel(dState.keyData, dState.workbook.sheets, mapName);
    document.getElementById('d-result-map').textContent = mapName + '（已下载）';
  } catch(e){ alert('导出失败：' + e.message); }
}

function exportKeyAsExcel(keyData, origSheets, filename){
  const wb = XLSX.utils.book_new();
  for (const [sn, sheetKey] of Object.entries(keyData.sheets)){
    const headers = (origSheets[sn] && origSheets[sn].data && origSheets[sn].data[0]) || [];
    const rows = [['列名','脱敏策略','原始值','脱敏值','参数']];
    for (const [ciStr, ck] of Object.entries(sheetKey.columns)){
      const ci = parseInt(ciStr);
      const colName = headers[ci] || `第${ci+1}列`;
      if (ck.strategy === 'none') continue;
      if (ck.mappings && Object.keys(ck.mappings).length > 0){
        let first = true;
        for (const [orig, masked] of Object.entries(ck.mappings)){
          rows.push([first ? colName : '', first ? ck.strategy : '', orig, masked, first ? JSON.stringify(ck.params || {}) : '']);
          first = false;
        }
      } else {
        const paramStr = Object.entries(ck.params || {}).map(([k,v]) => `${k}: ${v}`).join(', ');
        rows.push([colName, ck.strategy, '(数值类)', '(按参数计算)', paramStr]);
      }
      rows.push(['','','','','']);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:16},{wch:16},{wch:24},{wch:24},{wch:32}];
    XLSX.utils.book_append_sheet(wb, ws, sn.substring(0, 31));
  }
  const buf = XLSX.write(wb, { type:'array', bookType:'xlsx' });
  downloadBlob(new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
}

function renderDesensitizeReport(report){
  const card = document.getElementById('d-report-card');
  const summary = document.getElementById('d-report-summary');
  const detail = document.getElementById('d-report-detail');
  card.style.display = '';
  const colCount = report.sheets.reduce((s, sh) => s + sh.columns.length, 0);
  let summaryText = `共处理 ${report.sheets.length} 个工作表、${colCount} 列、${report.totalProcessed} 个单元格`;
  if (report.totalSkippedFormula > 0) summaryText += `，跳过 ${report.totalSkippedFormula} 个公式单元格`;
  summary.textContent = summaryText;
  let html = '';
  for (const sh of report.sheets){
    html += `<div class="report-sheet-name">📄 ${escapeHtml(sh.name)}（${sh.totalRows} 行数据）</div>`;
    html += `<table class="report-table"><thead><tr><th>列名</th><th>脱敏策略</th><th>处理单元格</th><th>跳过公式</th></tr></thead><tbody>`;
    for (const col of sh.columns){
      html += `<tr><td style="font-weight:500;">${escapeHtml(col.header)}</td><td><span class="report-stat strategy">${escapeHtml(col.strategy)}</span></td><td><span class="report-stat processed">${col.processed} 个</span></td><td>${col.skippedFormula > 0 ? `<span class="report-stat skipped">${col.skippedFormula} 个</span>` : '<span style="color:var(--text-dim);font-size:12px;">无</span>'}</td></tr>`;
    }
    html += '</tbody></table>';
  }
  detail.innerHTML = html;
}

// ========== 还原 UI ==========
function initRestoreUI(){
  document.getElementById('r-excel-card').addEventListener('click', () => document.getElementById('r-excel-input').click());
  document.getElementById('r-key-card').addEventListener('click', () => document.getElementById('r-key-input').click());
  document.getElementById('r-excel-input').addEventListener('change', async (e) => {
    if (e.target.files.length === 0) return;
    await handleRExcel(e.target.files[0]);
    e.target.value = '';
  });
  document.getElementById('r-key-input').addEventListener('change', async (e) => {
    if (e.target.files.length === 0) return;
    try {
      const text = await e.target.files[0].text();
      rState.keyData = JSON.parse(text);
      document.getElementById('r-key-card').classList.add('loaded');
      const st = document.getElementById('r-key-status');
      st.classList.remove('hidden');
      st.textContent = '✅ 密钥已载入';
      checkRestoreReady();
    } catch(err){ alert('密钥读取失败：' + err.message); }
    e.target.value = '';
  });
}

async function handleRExcel(file){
  showLoading('正在读取 Excel 文件...');
  try {
    const result = await readExcelFile(file);
    rState.excelName = file.name;
    rState.workbook = result;
    rState.currentSheet = result.sheetNames[0];
    document.getElementById('r-excel-card').classList.add('loaded');
    const st = document.getElementById('r-excel-status');
    st.classList.remove('hidden');
    st.textContent = '✅ ' + file.name;
    checkRestoreReady();
  } catch(e){ alert('读取文件失败：' + e.message); }
  hideLoading();
}

function checkRestoreReady(){
  const btn = document.getElementById('r-btn-next');
  btn.disabled = !(rState.workbook && rState.keyData);
  btn.onclick = rGoNext;
}

function updateRStep(step){
  rState.step = step;
  for (let i = 1; i <= 3; i++){
    const el = document.getElementById(`r-step-${i}`);
    el.classList.remove('active', 'done');
    if (i < step) el.classList.add('done'); else if (i === step) el.classList.add('active');
  }
  document.getElementById('r-import').classList.toggle('hidden', step !== 1);
  document.getElementById('r-preview').classList.toggle('hidden', step !== 2);
  document.getElementById('r-result').classList.toggle('hidden', step !== 3);

  const btnBack = document.getElementById('r-btn-back');
  const btnNext = document.getElementById('r-btn-next');
  btnBack.onclick = () => { if (rState.step === 2) updateRStep(1); else if (rState.step === 3) updateRStep(2); };

  if (step === 1){
    btnBack.classList.add('hidden');
    btnNext.textContent = '开始还原 →';
    btnNext.className = 'btn btn-primary';
    btnNext.disabled = !(rState.workbook && rState.keyData);
    btnNext.onclick = rGoNext;
  } else if (step === 2){
    btnBack.classList.remove('hidden');
    btnNext.textContent = '📥 下载还原文件';
    btnNext.className = 'btn btn-success';
    btnNext.disabled = false;
    btnNext.onclick = doExportRestored;
  } else if (step === 3){
    btnBack.classList.remove('hidden');
    btnNext.textContent = '🔄 重新开始';
    btnNext.className = 'btn btn-secondary';
    btnNext.disabled = false;
    btnNext.onclick = () => {
      rState = { step:1, excelName:'', workbook:null, keyData:null, currentSheet:'', restoredResult:null };
      document.getElementById('r-excel-card').classList.remove('loaded');
      document.getElementById('r-key-card').classList.remove('loaded');
      document.getElementById('r-excel-status').classList.add('hidden');
      document.getElementById('r-key-status').classList.add('hidden');
      updateRStep(1);
    };
  }
}

async function rGoNext(){
  if (rState.step !== 1) return;
  showLoading('正在还原数据...');
  await new Promise(r => setTimeout(r, 50));
  try {
    let result = doRestoreCore(rState.workbook.sheetNames, rState.workbook.sheets, rState.keyData, false);
    if (result.headerMismatches && result.headerMismatches.length > 0){
      hideLoading();
      const details = result.headerMismatches.map(m => `  ⚠️ 工作表「${m.sheet}」→ 第${m.colIndex + 1}列：密钥记录「${m.keyHeader}」，当前为「${m.currentHeader}」`).join('\n');
      const proceed = confirm(`⚠️ 列名不匹配！\n\n以下列的名称与脱敏时不一致，可能导致还原到错误的列：\n\n${details}\n\n是否仍然继续还原？`);
      if (!proceed) return;
      showLoading('正在还原数据...');
      result = doRestoreCore(rState.workbook.sheetNames, rState.workbook.sheets, rState.keyData, true);
    }
    rState.restoredResult = result.resultSheets;
    renderRestorePreview();
    if (result.report) renderRestoreReport(result.report);
    hideLoading();
    updateRStep(2);
  } catch(e){ hideLoading(); alert('还原失败：' + e.message); }
}

function renderRestorePreview(){
  const fi = document.getElementById('r-file-info');
  fi.innerHTML = `<span class="file-icon">📊</span><div><div class="file-name">${escapeHtml(rState.excelName)}</div><div class="file-detail">${rState.workbook.sheetNames.length} 个工作表 · 使用密钥还原</div></div>`;
  renderSheetTabs('r-sheet-tabs', rState.workbook.sheetNames, rState.currentSheet, (n) => { rState.currentSheet = n; renderRestoreCompare(); });
  renderRestoreCompare();
}

function renderRestoreCompare(){
  const table = document.getElementById('r-compare-table');
  const sn = rState.currentSheet;
  const maskedData = rState.workbook.sheets[sn].data;
  const restoredData = rState.restoredResult[sn].data;
  const headers = maskedData[0];
  let html = '<thead><tr><th>#</th>';
  for (const h of headers) html += `<th class="col-orig">${escapeHtml(String(h))}(脱敏)</th><th class="col-arrow"></th><th class="col-masked">${escapeHtml(String(h))}(还原)</th>`;
  html += '</tr></thead><tbody>';
  const rows = Math.min(maskedData.length - 1, 10);
  for (let i = 1; i <= rows; i++){
    html += `<tr><td style="color:var(--text-dim)">${i}</td>`;
    for (let j = 0; j < headers.length; j++){
      const masked = maskedData[i][j] !== undefined ? maskedData[i][j] : '';
      const restored = restoredData[i][j] !== undefined ? restoredData[i][j] : '';
      const changed = String(masked) !== String(restored);
      html += `<td class="col-orig">${escapeHtml(String(masked))}</td><td class="col-arrow">${changed ? '→' : ''}</td><td class="col-masked" style="${changed ? 'color:var(--success);font-weight:500;' : ''}">${escapeHtml(String(restored))}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';
  table.innerHTML = html;
}

function doExportRestored(){
  try {
    const baseName = rState.excelName.replace(/\.[^.]+$/, '');
    const filename = `${baseName}_还原.xlsx`;
    exportExcelFile(rState.workbook.sheetNames, rState.restoredResult, filename);
    document.getElementById('r-result-file').textContent = filename + '（已下载）';
    updateRStep(3);
  } catch(e){ alert('导出失败：' + e.message); }
}

function renderRestoreReport(report){
  const card = document.getElementById('r-report-card');
  const summary = document.getElementById('r-report-summary');
  const detail = document.getElementById('r-report-detail');
  card.style.display = '';
  const colCount = report.sheets.reduce((s, sh) => s + sh.columns.length, 0);
  let summaryText = `共还原 ${report.sheets.length} 个工作表、${colCount} 列、${report.totalRestored} 个单元格`;
  if (report.totalSkippedFormula > 0) summaryText += `，跳过 ${report.totalSkippedFormula} 个公式单元格`;
  summary.textContent = summaryText;
  let html = '';
  for (const sh of report.sheets){
    html += `<div class="report-sheet-name">📄 ${escapeHtml(sh.name)}（${sh.totalRows} 行数据）</div>`;
    html += `<table class="report-table"><thead><tr><th>列名</th><th>还原策略</th><th>还原单元格</th><th>跳过公式</th></tr></thead><tbody>`;
    for (const col of sh.columns){
      html += `<tr><td style="font-weight:500;">${escapeHtml(col.header)}</td><td><span class="report-stat strategy">${escapeHtml(col.strategy)}</span></td><td><span class="report-stat processed">${col.restored} 个</span></td><td>${col.skippedFormula > 0 ? `<span class="report-stat skipped">${col.skippedFormula} 个</span>` : '<span style="color:var(--text-dim);font-size:12px;">无</span>'}</td></tr>`;
    }
    html += '</tbody></table>';
  }
  detail.innerHTML = html;
}
