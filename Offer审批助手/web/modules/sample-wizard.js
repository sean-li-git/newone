/* ======================================================================
 * Offer审批助手 · 样本规划向导
 * 帮助用户评估画像所需的最低样本量
 * ====================================================================== */

/**
 * 样本分档标准
 */
const SAMPLE_TIERS = [
  { level: 'minimum', label: '最低可用', perGroup: 5,  totalMin: 50,  color: '#EF4444', desc: '每组 ≥5 条，共 ≥50 条' },
  { level: 'recommended', label: '推荐', perGroup: 15, totalMin: 150, color: '#F59E0B', desc: '每组 ≥15 条，共 ≥150 条' },
  { level: 'ideal', label: '理想', perGroup: 30, totalMin: 300, color: '#10B981', desc: '每组 ≥30 条，共 ≥300 条' },
];

/**
 * 计算有效组合数
 * @param {Object} dimensions - { countries: [], levels: [], channels: [], jobFamilies: [] }
 * @returns {number}
 */
function calcEffectiveGroups(dimensions) {
  const c = Math.max(1, (dimensions.countries || []).length);
  const l = Math.max(1, (dimensions.levels || []).length);
  const ch = Math.max(1, (dimensions.channels || []).length);
  const j = Math.max(1, (dimensions.jobFamilies || []).length);
  return c * l * ch * j;
}

/**
 * 生成样本规划报告
 * @param {Object} dimensions - 用户勾选的维度值
 * @param {number} currentSamples - 当前已有样本数
 * @returns {Object}
 */
function generateSamplePlan(dimensions, currentSamples) {
  const groups = calcEffectiveGroups(dimensions);
  
  const tiers = SAMPLE_TIERS.map(tier => {
    const needed = Math.max(tier.totalMin, groups * tier.perGroup);
    const gap = Math.max(0, needed - currentSamples);
    const progress = Math.min(100, Math.round((currentSamples / needed) * 100));
    return { ...tier, needed, gap, progress };
  });
  
  // 判断当前所在档位
  let currentTier = 'insufficient';
  if (currentSamples >= tiers[2].needed) currentTier = 'ideal';
  else if (currentSamples >= tiers[1].needed) currentTier = 'recommended';
  else if (currentSamples >= tiers[0].needed) currentTier = 'minimum';
  
  return {
    dimensions,
    groups,
    currentSamples,
    currentTier,
    tiers,
  };
}

/**
 * 从历史 Offer 中提取去重的维度值
 * @param {Array} historyOffers
 * @returns {Object}
 */
function extractDimensions(historyOffers) {
  const countries = new Set();
  const levels = new Set();
  const channels = new Set();
  const jobFamilies = new Set();
  
  for (const offer of historyOffers) {
    if (offer.country) countries.add(offer.country);
    if (offer.level) levels.add(offer.level);
    if (offer.channel) channels.add(offer.channel);
    if (offer.jobFamily) jobFamilies.add(offer.jobFamily);
  }
  
  return {
    countries: Array.from(countries).sort(),
    levels: Array.from(levels).sort(),
    channels: Array.from(channels).sort(),
    jobFamilies: Array.from(jobFamilies).sort(),
  };
}

/**
 * 计算每个组合的样本覆盖度
 * @returns {Array<{country, level, channel, jobFamily, count}>}
 */
function calcCoverageMatrix(historyOffers, dimensions) {
  const matrix = [];
  
  for (const country of (dimensions.countries.length ? dimensions.countries : [''])) {
    for (const level of (dimensions.levels.length ? dimensions.levels : [''])) {
      for (const channel of (dimensions.channels.length ? dimensions.channels : [''])) {
        for (const jobFamily of (dimensions.jobFamilies.length ? dimensions.jobFamilies : [''])) {
          const count = historyOffers.filter(o =>
            (!country || o.country === country) &&
            (!level || o.level === level) &&
            (!channel || o.channel === channel) &&
            (!jobFamily || o.jobFamily === jobFamily)
          ).length;
          
          matrix.push({ country, level, channel, jobFamily, count });
        }
      }
    }
  }
  
  return matrix;
}

/**
 * 获取覆盖度热力图数据（用于 Chart.js）
 * X 轴：职级，Y 轴：职位通道，按国家分组
 * @returns {Object}
 */
function getCoverageHeatmapData(matrix) {
  // 按国家分组
  const byCountry = {};
  for (const cell of matrix) {
    const key = cell.country || '全部';
    if (!byCountry[key]) byCountry[key] = [];
    byCountry[key].push(cell);
  }
  
  return byCountry;
}
