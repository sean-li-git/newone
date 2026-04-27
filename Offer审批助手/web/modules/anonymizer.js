/* ======================================================================
 * Offer审批助手 · 脱敏管道
 * 在调用 AI 前对薪酬数据进行脱敏处理
 * 复用「数据脱敏助手」的策略思路
 * ====================================================================== */

/**
 * 9 种脱敏策略
 */
const ANON_STRATEGIES = {
  // 1. 数值偏移（加减随机比例）
  numericShift: (val, ratio) => {
    const n = Number(val);
    if (isNaN(n)) return val;
    const shift = n * (ratio || 0.1) * (Math.random() > 0.5 ? 1 : -1);
    return Math.round(n + shift);
  },
  // 2. 数值取整（到千/万）
  numericRound: (val, precision) => {
    const n = Number(val);
    if (isNaN(n)) return val;
    const p = precision || 1000;
    return Math.round(n / p) * p;
  },
  // 3. 文本替换（用占位符）
  textReplace: (val, placeholder) => placeholder || '[已脱敏]',
  // 4. 文本截断
  textTruncate: (val, keepChars) => {
    const s = String(val);
    return s.slice(0, keepChars || 1) + '***';
  },
  // 5. 哈希（取前 6 位）
  hash: (val) => {
    let h = 0;
    const s = String(val);
    for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
    return 'H' + Math.abs(h).toString(36).slice(0, 6).toUpperCase();
  },
  // 6. 日期模糊（只保留年月）
  dateFuzzy: (val) => String(val).slice(0, 7),
  // 7. 枚举泛化（如职级 L7 → 高级）
  generalize: (val, mapping) => (mapping && mapping[val]) || val,
  // 8. 删除
  suppress: () => '',
  // 9. 保留原值
  keep: (val) => val,
};

/**
 * 默认脱敏配置
 */
const DEFAULT_ANON_CONFIG = {
  baseSalary: { strategy: 'numericRound', params: [10000] },
  allowance: { strategy: 'numericRound', params: [1000] },
  perfMonthly: { strategy: 'numericRound', params: [1000] },
  annualBonus: { strategy: 'numericRound', params: [10000] },
  equityAnnual: { strategy: 'numericRound', params: [10000] },
  signStock: { strategy: 'numericRound', params: [10000] },
  signBonus: { strategy: 'numericRound', params: [10000] },
  relocation: { strategy: 'numericRound', params: [10000] },
  cashIncome: { strategy: 'numericRound', params: [10000] },
  annualIncome: { strategy: 'numericRound', params: [10000] },
  country: { strategy: 'keep' },
  level: { strategy: 'keep' },
  channel: { strategy: 'keep' },
  jobFamily: { strategy: 'keep' },
  sourceCompany: { strategy: 'textTruncate', params: [1] },
};

/**
 * 对文案中的数字和敏感信息进行脱敏
 */
function desensitizeForAI(text, offer) {
  let result = text;
  // 替换具体公司名
  if (offer && offer.sourceCompany) {
    result = result.replace(new RegExp(offer.sourceCompany, 'g'), ANON_STRATEGIES.textTruncate(offer.sourceCompany, 1));
  }
  // 数字取整（所有 5 位以上数字）
  result = result.replace(/\d{5,}/g, (match) => ANON_STRATEGIES.numericRound(match, 10000).toString());
  return result;
}

/**
 * 对 Offer 对象进行完整脱敏
 */
function desensitizeOffer(offer, config) {
  config = config || DEFAULT_ANON_CONFIG;
  const result = { ...offer };
  for (const [field, rule] of Object.entries(config)) {
    if (result[field] !== undefined) {
      const fn = ANON_STRATEGIES[rule.strategy];
      if (fn) result[field] = fn(result[field], ...(rule.params || []));
    }
  }
  return result;
}
