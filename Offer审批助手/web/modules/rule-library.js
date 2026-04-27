/* ======================================================================
 * Offer审批助手 · 双层规则库 + 配置表管理
 * 统一规则层 + 个人规则层 + 配置表（因子权重/竞企清单/稀缺岗位/能力标签/分摊年限）
 * ====================================================================== */

/**
 * 加载所有规则
 * @param {string} scope - 'unified' | 'personal' | 'all'
 */
async function loadRules(scope) {
  const all = await dbGetAll(STORES.rules);
  if (scope === 'all') return all;
  return all.filter(r => r.scope === scope);
}

/**
 * 保存规则（新增或更新）
 */
async function saveRule(rule) {
  rule.updatedAt = new Date().toISOString();
  if (!rule.createdAt) rule.createdAt = rule.updatedAt;
  await dbPut(STORES.rules, rule);
}

/**
 * 删除规则
 */
async function deleteRule(id) {
  await dbDelete(STORES.rules, id);
}

/**
 * 批量导入规则（从 JSON 或 Excel）
 */
async function importRules(rules, scope) {
  for (const rule of rules) {
    rule.scope = scope || rule.scope || 'unified';
    rule.id = rule.id || generateId();
    rule.enabled = rule.enabled !== false;
    rule.source = rule.source || 'user';
    await saveRule(rule);
  }
}

/**
 * 导出规则为 JSON
 */
async function exportRulesJSON(scope) {
  const rules = await loadRules(scope || 'all');
  return JSON.stringify(rules, null, 2);
}

// ========== 配置表管理 ==========

/**
 * 配置表 key 定义
 */
const CONFIG_KEYS = {
  FACTOR_WEIGHTS: 'factor_weights',       // 因子权重表
  COMPETITOR_LIST: 'competitor_list',      // 竞企清单
  RARE_POSITIONS: 'rare_positions',        // 稀缺岗位清单
  ABILITY_TAGS: 'ability_tags',            // 能力标签库
  AMORTIZATION_YEARS: 'amortization_years', // 分摊年限
};

/**
 * 获取配置
 */
async function getConfig(key, defaultValue) {
  const record = await dbGet(STORES.configs, key);
  return record ? record.value : defaultValue;
}

/**
 * 设置配置
 */
async function setConfig(key, value) {
  await dbPut(STORES.configs, { key, value, updatedAt: new Date().toISOString() });
}

/**
 * 默认因子权重表
 */
function getDefaultFactorWeights() {
  return {
    baseline: { enabled: true },
    raise: { defaultCap: 0.35, learnedMedian: null },
    competition: {
      tierPremium: {
        '一线竞对': 0.15,
        '二线竞对': 0.08,
        '非竞对': 0,
      },
    },
    internal: { ceilingRatio: 1.2, floorRatio: 0.85 },
    urgency: {
      business: { '紧急': 0.10, '正常': 0, '不紧急': -0.05 },
      onboarding: { '1个月内': 0.08, '1-3个月': 0, '3个月以上': -0.03 },
    },
    ability: { perTag: {}, capPercent: 0.20 },
    rarity: { positions: [], uplift: 0.10 },
    structure: { basePct: 0.55, stockPct: 0.25, signBonusPct: 0.10 },
    userOverrides: {},
  };
}

/**
 * 加载因子权重（合并系统默认 + 用户覆盖）
 */
async function loadFactorWeights() {
  const saved = await getConfig(CONFIG_KEYS.FACTOR_WEIGHTS);
  const defaults = getDefaultFactorWeights();
  if (!saved) return defaults;
  // 深度合并
  return deepMerge(defaults, saved);
}

/**
 * 保存因子权重
 */
async function saveFactorWeights(weights) {
  await setConfig(CONFIG_KEYS.FACTOR_WEIGHTS, weights);
}

/**
 * 加载竞企清单
 */
async function loadCompetitorList() {
  return await getConfig(CONFIG_KEYS.COMPETITOR_LIST, []);
}

async function saveCompetitorList(list) {
  await setConfig(CONFIG_KEYS.COMPETITOR_LIST, list);
}

/**
 * 加载稀缺岗位清单
 */
async function loadRarePositions() {
  return await getConfig(CONFIG_KEYS.RARE_POSITIONS, []);
}

async function saveRarePositions(list) {
  await setConfig(CONFIG_KEYS.RARE_POSITIONS, list);
}

/**
 * 加载能力标签库
 */
async function loadAbilityTags() {
  return await getConfig(CONFIG_KEYS.ABILITY_TAGS, []);
}

async function saveAbilityTags(tags) {
  await setConfig(CONFIG_KEYS.ABILITY_TAGS, tags);
}

/**
 * 深度合并工具
 */
function deepMerge(target, source) {
  const result = { ...target };
  for (const key of Object.keys(source)) {
    if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key]) &&
        target[key] && typeof target[key] === 'object' && !Array.isArray(target[key])) {
      result[key] = deepMerge(target[key], source[key]);
    } else {
      result[key] = source[key];
    }
  }
  return result;
}

/**
 * 导入行业通用规则（初始化用）
 */
async function loadIndustryDefaultRules() {
  const existing = await loadRules('unified');
  if (existing.length > 0) return false; // 已有规则，不重复导入
  
  const defaults = getIndustryDefaultRules();
  await importRules(defaults, 'unified');
  return true;
}

/**
 * 行业通用规则集
 */
function getIndustryDefaultRules() {
  return [
    {
      id: 'rule_bandwidth_base',
      name: '基本月薪带宽检查',
      scope: 'unified',
      category: 'bandwidth',
      enabled: true,
      priority: 10,
      conditions: [
        { field: 'offer.baseSalary', operator: 'gt', value: 0 },
      ],
      action: { type: 'check', message: '基本月薪需在职级带宽范围内（需配置带宽数据后生效）' },
      source: 'system',
    },
    {
      id: 'rule_raise_cap',
      name: '涨幅上限检查（35%）',
      scope: 'unified',
      category: 'raise',
      enabled: true,
      priority: 20,
      conditions: [
        { field: 'context.raisePercent', operator: 'gt', value: 0.35 },
      ],
      action: { type: 'warn', message: '候选人薪资增幅超过 35%，需额外审批' },
      source: 'system',
    },
    {
      id: 'rule_raise_extreme',
      name: '涨幅红线（50%）',
      scope: 'unified',
      category: 'raise',
      enabled: true,
      priority: 15,
      conditions: [
        { field: 'context.raisePercent', operator: 'gt', value: 0.50 },
      ],
      action: { type: 'block', message: '候选人薪资增幅超过 50%，禁止通过' },
      source: 'system',
    },
    {
      id: 'rule_structure_base_ratio',
      name: 'Base 占比检查',
      scope: 'unified',
      category: 'structure',
      enabled: true,
      priority: 30,
      conditions: [
        { field: 'context.baseRatio', operator: 'lt', value: 0.40 },
      ],
      action: { type: 'warn', message: '基本月薪占年现金收入不足 40%，结构偏激进' },
      source: 'system',
    },
    {
      id: 'rule_sign_bonus_check',
      name: '签字费合理性检查',
      scope: 'unified',
      category: 'fairness',
      enabled: true,
      priority: 40,
      conditions: [
        { field: 'offer.signBonus', operator: 'gt', value: 0 },
      ],
      action: { type: 'check', message: '含签字费，请确认业务必要性和审批权限' },
      source: 'system',
    },
    {
      id: 'rule_equity_check',
      name: '股票授予检查',
      scope: 'unified',
      category: 'authority',
      enabled: true,
      priority: 35,
      conditions: [
        { field: 'offer.equityAnnual', operator: 'gt', value: 0 },
      ],
      action: { type: 'check', message: '含股票授予，请确认授予额度在权限范围内' },
      source: 'system',
    },
  ];
}
