/* ======================================================================
 * Offer审批助手 · 画像学习器
 * 核心能力：
 * 1. 按 4 维（国家/城市 × 职级 × 职位通道 × 职位类型）强制切分
 * 2. 每个子集内独立计算 P25/P50/P75
 * 3. 条件型规律识别（来源公司溢价、能力标签效应、紧迫度加成、结构偏好）
 * 4. 从历史数据拟合因子权重默认值
 * ====================================================================== */

/**
 * 解析历史 Offer Excel 并导入 IndexedDB
 * @param {File} file
 * @returns {Promise<{count: number, errors: string[]}>}
 */
async function importHistoryOffers(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false, cellFormula: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  if (!rawData || rawData.length < 2) return { count: 0, errors: ['Excel 中无数据行'] };

  const headers = rawData[0].map(h => String(h).trim().toLowerCase());
  const errors = [];
  let count = 0;

  const colMap = {};
  const HISTORY_ALIASES = {
    id:            ['offer id', 'offerid', 'id'],
    approvalDate:  ['审批日期', 'approval date', 'date'],
    approvalResult:['审批结果', 'approval result', 'result'],
    currency:      ['币种', 'currency'],
    country:       ['国家/城市', '国家', '城市', 'country', 'city', 'location'],
    level:         ['职级', 'level', 'grade', 'band'],
    channel:       ['职位通道', '通道', 'channel', 'track'],
    jobFamily:     ['职位类型', '职位族', 'job family', 'job type'],
    baseSalary:    ['基本月薪', '月薪', 'base', 'base salary'],
    allowance:     ['津贴', '补贴', '津贴及其他现金补贴', 'allowance'],
    perfMonthly:   ['绩效', '绩效月均', 'performance', 'perf'],
    annualBonus:   ['年终奖', 'annual bonus', 'bonus'],
    equityAnnual:  ['入职授予', '入职授予年化', 'equity', 'rsu'],
    signStock:     ['签字股票', '签字股票总额', 'sign stock'],
    signBonus:     ['签字费', '签字费总额', 'sign bonus', 'signing bonus'],
    relocation:    ['安家费', 'relocation'],
    cashIncome:    ['现金收入', 'cash income'],
    annualIncome:  ['年收入', 'annual income'],
    annualIncomeTotal: ['年收入含一次性', '总年收入'],
    sourceCompany: ['来源公司', 'source company', 'from company'],
    currentSalary: ['候选人现薪', '现薪', 'current salary'],
    raisePercent:  ['涨幅', '涨幅%', 'raise', 'raise%'],
    bizUrgency:    ['业务紧迫度', 'business urgency'],
    onboardUrgency:['入职时间紧迫度', 'onboard urgency'],
    abilityTags:   ['能力标签', 'ability tags', 'tags'],
    notes:         ['备注', 'notes', 'remark'],
  };

  for (let i = 0; i < headers.length; i++) {
    for (const [key, aliases] of Object.entries(HISTORY_ALIASES)) {
      for (const alias of aliases) {
        if (headers[i] === alias.toLowerCase() || headers[i].includes(alias.toLowerCase())) {
          colMap[key] = i;
          break;
        }
      }
    }
  }

  for (let row = 1; row < rawData.length; row++) {
    const d = rawData[row];
    if (!d || d.every(v => !v && v !== 0)) continue;

    const g = (key) => colMap[key] !== undefined ? d[colMap[key]] : '';
    const gn = (key) => { const v = g(key); const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? 0 : n; };

    const offer = {
      id: String(g('id') || ('ho_' + Date.now().toString(36) + '_' + row)),
      approvalDate: String(g('approvalDate') || ''),
      approvalResult: String(g('approvalResult') || ''),
      currency: String(g('currency') || 'CNY'),
      country: String(g('country') || ''),
      level: String(g('level') || ''),
      channel: String(g('channel') || ''),
      jobFamily: String(g('jobFamily') || ''),
      baseSalary: gn('baseSalary'),
      allowance: gn('allowance'),
      perfMonthly: gn('perfMonthly'),
      annualBonus: gn('annualBonus'),
      equityAnnual: gn('equityAnnual'),
      signStock: gn('signStock'),
      signBonus: gn('signBonus'),
      relocation: gn('relocation'),
      cashIncome: gn('cashIncome'),
      annualIncome: gn('annualIncome'),
      annualIncomeTotal: gn('annualIncomeTotal'),
      sourceCompany: String(g('sourceCompany') || ''),
      currentSalary: gn('currentSalary'),
      raisePercent: gn('raisePercent'),
      bizUrgency: String(g('bizUrgency') || '正常'),
      onboardUrgency: String(g('onboardUrgency') || '1-3个月'),
      abilityTags: String(g('abilityTags') || '').split(/[;；,，]/).filter(Boolean).map(s => s.trim()),
      notes: String(g('notes') || ''),
      importedAt: new Date().toISOString(),
    };

    // 自动核算缺失的计算项
    if (!offer.cashIncome) {
      offer.cashIncome = offer.baseSalary * 12 + offer.perfMonthly * 12 + offer.annualBonus + offer.allowance;
    }
    if (!offer.annualIncome) {
      offer.annualIncome = offer.cashIncome + offer.equityAnnual;
    }

    if (!offer.country || !offer.level) {
      errors.push('第 ' + (row + 1) + ' 行缺少国家/城市或职级');
      continue;
    }

    await dbPut(STORES.historyOffers, offer);
    count++;
  }

  return { count, errors };
}

/**
 * 学习画像 — 核心入口
 * 按 4 维切分，对每个子集独立计算统计量
 * @returns {Promise<ProfileInsight[]>}
 */
async function learnProfile() {
  const offers = await dbGetAll(STORES.historyOffers);
  if (offers.length === 0) return [];

  const insights = [];

  // Step 1: 按 4 维 groupBy
  const groups = groupBySlice(offers);

  // Step 2: 对每个子集计算数值型画像
  const numericDims = [
    { key: 'baseSalary', label: '基本月薪' },
    { key: 'cashIncome', label: '现金收入' },
    { key: 'annualIncome', label: '年收入' },
    { key: 'annualIncomeTotal', label: '年总收入(含一次性)' },
    { key: 'equityAnnual', label: '入职授予年化' },
    { key: 'perfMonthly', label: '绩效月均' },
    { key: 'annualBonus', label: '年终奖' },
  ];

  for (const [sliceKeyStr, groupOffers] of Object.entries(groups)) {
    const sliceKey = JSON.parse(sliceKeyStr);
    for (const dim of numericDims) {
      const values = groupOffers.map(o => Number(o[dim.key]) || 0).filter(v => v > 0).sort((a, b) => a - b);
      if (values.length < 2) continue;

      insights.push({
        id: 'pi_' + generateId(),
        sliceKey,
        type: 'numeric',
        dimension: dim.key,
        dimensionLabel: dim.label,
        pattern: {
          p25: percentile(values, 25),
          median: percentile(values, 50),
          p75: percentile(values, 75),
          min: values[0],
          max: values[values.length - 1],
          mean: Math.round(values.reduce((s, v) => s + v, 0) / values.length),
        },
        sampleCount: values.length,
        confidence: calcConfidence(values.length),
        userStatus: 'suggested',
        sourceOfferIds: groupOffers.map(o => o.id),
      });
    }

    // Step 3: 条件型规律 — 来源公司溢价
    const compInsights = calcCompetitorPremium(groupOffers, sliceKey);
    insights.push(...compInsights);

    // Step 4: 条件型规律 — 紧迫度效应
    const urgInsights = calcUrgencyEffect(groupOffers, sliceKey);
    insights.push(...urgInsights);

    // Step 5: 条件型规律 — 结构偏好
    const structInsight = calcStructurePreference(groupOffers, sliceKey);
    if (structInsight) insights.push(structInsight);
  }

  return insights;
}

/**
 * 按 4 维 groupBy
 */
function groupBySlice(offers) {
  const groups = {};
  for (const o of offers) {
    const key = JSON.stringify({
      country: o.country || '未知',
      level: o.level || '未知',
      channel: o.channel || '未知',
      jobFamily: o.jobFamily || '未知',
    });
    if (!groups[key]) groups[key] = [];
    groups[key].push(o);
  }
  return groups;
}

/**
 * 百分位数计算
 */
function percentile(sortedArr, p) {
  if (sortedArr.length === 0) return 0;
  const idx = (p / 100) * (sortedArr.length - 1);
  const low = Math.floor(idx);
  const high = Math.ceil(idx);
  if (low === high) return Math.round(sortedArr[low]);
  return Math.round(sortedArr[low] + (sortedArr[high] - sortedArr[low]) * (idx - low));
}

/**
 * 置信度评估
 */
function calcConfidence(sampleCount) {
  if (sampleCount >= 30) return 1.0;
  if (sampleCount >= 15) return 0.8;
  if (sampleCount >= 5) return 0.5;
  return 0.3;
}

/**
 * 来源公司溢价分析
 */
function calcCompetitorPremium(groupOffers, sliceKey) {
  const insights = [];
  const bySource = {};
  for (const o of groupOffers) {
    const src = o.sourceCompany || '未知';
    if (!bySource[src]) bySource[src] = [];
    bySource[src].push(o);
  }

  const allBase = groupOffers.map(o => o.baseSalary).filter(v => v > 0);
  if (allBase.length < 3) return insights;
  const overallMedian = percentile(allBase.sort((a, b) => a - b), 50);

  for (const [src, srcOffers] of Object.entries(bySource)) {
    if (src === '未知' || srcOffers.length < 2) continue;
    const srcBase = srcOffers.map(o => o.baseSalary).filter(v => v > 0).sort((a, b) => a - b);
    if (srcBase.length < 2) continue;
    const srcMedian = percentile(srcBase, 50);
    const premium = overallMedian > 0 ? Math.round(((srcMedian / overallMedian) - 1) * 10000) / 10000 : 0;

    insights.push({
      id: 'pi_' + generateId(),
      sliceKey,
      type: 'conditional',
      dimension: '来源公司.' + src + '.base_premium',
      dimensionLabel: '来自「' + src + '」的 base 溢价',
      pattern: { condition: '来源公司 = ' + src, effect: premium },
      sampleCount: srcBase.length,
      confidence: calcConfidence(srcBase.length),
      userStatus: 'suggested',
      sourceOfferIds: srcOffers.map(o => o.id),
    });
  }
  return insights;
}

/**
 * 紧迫度效应分析
 */
function calcUrgencyEffect(groupOffers, sliceKey) {
  const insights = [];
  const allCash = groupOffers.map(o => o.cashIncome).filter(v => v > 0);
  if (allCash.length < 5) return insights;
  const overallMedian = percentile(allCash.sort((a, b) => a - b), 50);

  const byUrgency = {};
  for (const o of groupOffers) {
    const u = o.bizUrgency || '正常';
    if (!byUrgency[u]) byUrgency[u] = [];
    byUrgency[u].push(o);
  }

  for (const [urg, urgOffers] of Object.entries(byUrgency)) {
    if (urgOffers.length < 2) continue;
    const urgCash = urgOffers.map(o => o.cashIncome).filter(v => v > 0).sort((a, b) => a - b);
    if (urgCash.length < 2) continue;
    const urgMedian = percentile(urgCash, 50);
    const effect = overallMedian > 0 ? Math.round(((urgMedian / overallMedian) - 1) * 10000) / 10000 : 0;

    insights.push({
      id: 'pi_' + generateId(),
      sliceKey,
      type: 'conditional',
      dimension: '业务紧迫度.' + urg + '.cash_effect',
      dimensionLabel: '业务紧迫度「' + urg + '」对现金收入的影响',
      pattern: { condition: '业务紧迫度 = ' + urg, effect },
      sampleCount: urgCash.length,
      confidence: calcConfidence(urgCash.length),
      userStatus: 'suggested',
      sourceOfferIds: urgOffers.map(o => o.id),
    });
  }
  return insights;
}

/**
 * 结构偏好分析
 */
function calcStructurePreference(groupOffers, sliceKey) {
  const valid = groupOffers.filter(o => o.cashIncome > 0 && o.baseSalary > 0);
  if (valid.length < 3) return null;

  const baseRatios = valid.map(o => (o.baseSalary * 12) / o.cashIncome).sort((a, b) => a - b);
  const equityRatios = valid.filter(o => o.annualIncome > 0).map(o => o.equityAnnual / o.annualIncome).sort((a, b) => a - b);

  return {
    id: 'pi_' + generateId(),
    sliceKey,
    type: 'conditional',
    dimension: '结构偏好',
    dimensionLabel: '薪酬结构偏好',
    pattern: {
      condition: '结构分析',
      basePctMedian: Math.round(percentile(baseRatios, 50) * 1000) / 1000,
      equityPctMedian: equityRatios.length >= 3 ? Math.round(percentile(equityRatios, 50) * 1000) / 1000 : null,
    },
    sampleCount: valid.length,
    confidence: calcConfidence(valid.length),
    userStatus: 'suggested',
    sourceOfferIds: valid.map(o => o.id),
  };
}

/**
 * 从历史数据拟合因子权重默认值
 * @returns {Object} 部分因子权重覆盖
 */
async function fitFactorWeights() {
  const offers = await dbGetAll(STORES.historyOffers);
  if (offers.length < 10) return {};

  const fitted = {};

  // 涨幅中位数
  const raises = offers.map(o => o.raisePercent).filter(v => v > 0).sort((a, b) => a - b);
  if (raises.length >= 5) {
    fitted.raise = { learnedMedian: percentile(raises, 50) };
  }

  // 竞对溢价
  const competitors = await loadCompetitorList();
  if (competitors.length > 0) {
    const compOffers = offers.filter(o => competitors.includes(o.sourceCompany));
    const nonCompOffers = offers.filter(o => o.sourceCompany && !competitors.includes(o.sourceCompany));
    if (compOffers.length >= 3 && nonCompOffers.length >= 3) {
      const compBase = compOffers.map(o => o.baseSalary).filter(v => v > 0).sort((a, b) => a - b);
      const nonBase = nonCompOffers.map(o => o.baseSalary).filter(v => v > 0).sort((a, b) => a - b);
      const compMedian = percentile(compBase, 50);
      const nonMedian = percentile(nonBase, 50);
      if (nonMedian > 0) {
        const premium = Math.round(((compMedian / nonMedian) - 1) * 100) / 100;
        fitted.competition = { tierPremium: { '一线竞对': premium } };
      }
    }
  }

  return fitted;
}
