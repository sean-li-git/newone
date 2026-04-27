/* ======================================================================
 * Offer审批助手 · Excel 解析器
 * 将上传的薪酬 Excel → 14 字段标准 Offer 对象
 * ====================================================================== */

/**
 * 14 个标准薪酬字段定义
 * group: 分类（元信息/固定/浮动/长期激励/一次性/计算项）
 * key: 字段标识符
 * label: 中文显示名
 * type: 数据类型（text/number/computed）
 * computed: 是否为系统计算项
 */
const SALARY_FIELDS = [
  // 元信息（5 项）
  { group: '元信息', key: 'currency',    label: '币种',       type: 'text',    computed: false },
  { group: '元信息', key: 'country',     label: '国家/城市',   type: 'text',    computed: false },
  { group: '元信息', key: 'level',       label: '职级',       type: 'text',    computed: false },
  { group: '元信息', key: 'channel',     label: '职位通道',    type: 'text',    computed: false },
  { group: '元信息', key: 'jobFamily',   label: '职位类型',    type: 'text',    computed: false },
  // 固定（2 项）
  { group: '固定',   key: 'baseSalary',  label: '基本月薪',    type: 'number',  computed: false },
  { group: '固定',   key: 'allowance',   label: '津贴及其他现金补贴', type: 'number', computed: false },
  // 浮动（2 项）
  { group: '浮动',   key: 'perfMonthly', label: '绩效（月均）', type: 'number',  computed: false },
  { group: '浮动',   key: 'annualBonus', label: '年终奖',      type: 'number',  computed: false },
  // 长期激励（1 项）
  { group: '长期激励', key: 'equityAnnual', label: '入职授予（年化）', type: 'number', computed: false },
  // 一次性（3 项）
  { group: '一次性', key: 'signStock',   label: '签字股票总额', type: 'number',  computed: false },
  { group: '一次性', key: 'signBonus',   label: '签字费总额',   type: 'number',  computed: false },
  { group: '一次性', key: 'relocation',  label: '安家费',      type: 'number',  computed: false },
  // 计算项（3 项 — 系统自动核算）
  { group: '计算项', key: 'cashIncome',        label: '现金收入',                  type: 'number', computed: true },
  { group: '计算项', key: 'annualIncome',      label: '年收入',                    type: 'number', computed: true },
  { group: '计算项', key: 'annualIncomeTotal',  label: '年收入（含一次性）',          type: 'number', computed: true },
];

/**
 * 字段别名映射表 — 兼容不同模板的列名写法
 */
const FIELD_ALIASES = {
  currency:    ['币种', 'currency', 'ccy'],
  country:     ['国家/城市', '国家', '城市', 'country', 'city', 'location'],
  level:       ['职级', 'level', 'grade', 'band'],
  channel:     ['职位通道', '通道', 'channel', 'track', 'job track'],
  jobFamily:   ['职位类型', '职位族', 'job family', 'job type', 'family'],
  baseSalary:  ['基本月薪', '月薪', 'base', 'base salary', 'monthly salary', '基本工资'],
  allowance:   ['津贴', '补贴', '津贴及其他现金补贴', 'allowance', '其他现金'],
  perfMonthly: ['绩效', '绩效月均', '月均绩效', 'performance', 'perf'],
  annualBonus: ['年终奖', '年奖', 'annual bonus', 'bonus', '年终'],
  equityAnnual:['入职授予', '入职授予年化', 'equity', 'rsu', 'stock grant', '年化授予'],
  signStock:   ['签字股票', '签字股票总额', 'sign-on stock', 'sign stock'],
  signBonus:   ['签字费', '签字费总额', 'sign-on bonus', 'signing bonus', 'sign bonus'],
  relocation:  ['安家费', 'relocation', 'relo', '搬迁费'],
  cashIncome:  ['现金收入', 'cash income', 'total cash'],
  annualIncome:['年收入', 'annual income', 'total comp'],
  annualIncomeTotal: ['年收入含一次性', '总年收入', 'total comp with one-time'],
};

/**
 * 自动匹配 Excel 列名 → 标准字段 key
 */
function matchFieldKey(header) {
  if (!header) return null;
  const h = String(header).trim().toLowerCase();
  for (const [key, aliases] of Object.entries(FIELD_ALIASES)) {
    for (const alias of aliases) {
      if (h === alias.toLowerCase() || h.includes(alias.toLowerCase())) {
        return key;
      }
    }
  }
  return null;
}

/**
 * 解析薪酬 Excel 文件 → Offer 对象数组
 * @param {File} file
 * @returns {Promise<{offers: Array, sheetName: string, rawData: Array, fieldMapping: Object, errors: Array}>}
 */
async function parseSalaryExcel(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false, cellFormula: true });
  
  // 取第一个 Sheet
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  
  if (!rawData || rawData.length < 2) {
    return { offers: [], sheetName, rawData, fieldMapping: {}, errors: ['Excel 中无数据行'] };
  }
  
  // 自动匹配列名
  const headers = rawData[0];
  const fieldMapping = {}; // colIndex → fieldKey
  const unmatchedCols = [];
  
  for (let i = 0; i < headers.length; i++) {
    const key = matchFieldKey(headers[i]);
    if (key) {
      fieldMapping[i] = key;
    } else if (headers[i] && String(headers[i]).trim()) {
      unmatchedCols.push({ colIndex: i, header: headers[i] });
    }
  }
  
  // 解析数据行
  const offers = [];
  const errors = [];
  
  for (let row = 1; row < rawData.length; row++) {
    const rowData = rawData[row];
    // 跳过空行
    if (!rowData || rowData.every(v => !v && v !== 0)) continue;
    
    const offer = { id: `offer_${Date.now()}_${row}`, _row: row };
    const excelValues = {}; // 保存 Excel 原始值（含计算项）
    
    for (const [colIdx, fieldKey] of Object.entries(fieldMapping)) {
      const rawVal = rowData[parseInt(colIdx)];
      const fieldDef = SALARY_FIELDS.find(f => f.key === fieldKey);
      
      if (fieldDef && fieldDef.type === 'number') {
        const num = parseFloat(String(rawVal).replace(/,/g, ''));
        offer[fieldKey] = isNaN(num) ? 0 : num;
      } else {
        offer[fieldKey] = rawVal !== undefined && rawVal !== null ? String(rawVal).trim() : '';
      }
      excelValues[fieldKey] = offer[fieldKey];
    }
    
    // 补齐缺失字段
    for (const f of SALARY_FIELDS) {
      if (offer[f.key] === undefined) {
        offer[f.key] = f.type === 'number' ? 0 : '';
      }
    }
    
    offer._excelValues = excelValues;
    offers.push(offer);
  }
  
  // 检查必填字段覆盖率
  const requiredKeys = ['baseSalary', 'level', 'country'];
  for (const key of requiredKeys) {
    const mapped = Object.values(fieldMapping).includes(key);
    if (!mapped) {
      errors.push(`未找到「${SALARY_FIELDS.find(f => f.key === key)?.label || key}」字段的匹配列`);
    }
  }
  
  return { offers, sheetName, rawData, fieldMapping, headers, unmatchedCols, errors };
}

/**
 * 自动核算 3 个计算项
 * @param {Object} offer
 * @param {number} amortizationYears - 签字费/签字股票分摊年限（规则库可配置）
 * @returns {Object} { cashIncome, annualIncome, annualIncomeTotal }
 */
function computeSalaryFields(offer, amortizationYears) {
  const N = amortizationYears || 2;
  
  const base = Number(offer.baseSalary) || 0;
  const allowance = Number(offer.allowance) || 0;
  const perf = Number(offer.perfMonthly) || 0;
  const bonus = Number(offer.annualBonus) || 0;
  const equity = Number(offer.equityAnnual) || 0;
  const signStock = Number(offer.signStock) || 0;
  const signBonus = Number(offer.signBonus) || 0;
  const relo = Number(offer.relocation) || 0;
  
  // 现金收入 = 基本月薪×12 + 绩效×12 + 年终奖 + 津贴
  const cashIncome = Math.round((base * 12 + perf * 12 + bonus + allowance) * 100) / 100;
  
  // 年收入 = 现金收入 + 入职授予年化
  const annualIncome = Math.round((cashIncome + equity) * 100) / 100;
  
  // 年收入（含一次性）= 年收入 + 签字股票/N + 签字费/N + 安家费分摊
  const annualIncomeTotal = Math.round((annualIncome + signStock / N + signBonus / N + relo / N) * 100) / 100;
  
  return { cashIncome, annualIncome, annualIncomeTotal };
}

/**
 * 对每个 Offer 执行自动核算，并与 Excel 原始值对比
 * @returns {Array<{key, label, excelValue, computedValue, match, diff}>}
 */
function validateComputedFields(offer, amortizationYears) {
  const computed = computeSalaryFields(offer, amortizationYears);
  const results = [];
  
  const computedFields = [
    { key: 'cashIncome', label: '现金收入' },
    { key: 'annualIncome', label: '年收入' },
    { key: 'annualIncomeTotal', label: '年收入（含一次性）' },
  ];
  
  for (const f of computedFields) {
    const excelVal = offer._excelValues?.[f.key];
    const computedVal = computed[f.key];
    const excelNum = parseFloat(String(excelVal).replace(/,/g, '')) || 0;
    
    // 判断是否一致（允许 ±0.01 浮点误差）
    const match = !excelVal || Math.abs(excelNum - computedVal) < 0.02;
    const diff = excelVal ? Math.round((computedVal - excelNum) * 100) / 100 : null;
    
    results.push({
      key: f.key,
      label: f.label,
      excelValue: excelVal || '(未填)',
      computedValue: computedVal,
      match,
      diff,
    });
  }
  
  // 同时写回计算值
  offer.cashIncome = computed.cashIncome;
  offer.annualIncome = computed.annualIncome;
  offer.annualIncomeTotal = computed.annualIncomeTotal;
  
  return results;
}
