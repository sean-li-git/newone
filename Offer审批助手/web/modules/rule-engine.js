/* ======================================================================
 * Offer审批助手 · 声明式规则引擎
 * 轻量自研执行器，约 250 行
 * 规则格式：JSON 声明式，支持条件组合、数值比较、枚举匹配
 * ====================================================================== */

/**
 * 规则结构定义
 * {
 *   id: string,
 *   name: string,
 *   scope: 'unified' | 'personal',
 *   category: 'bandwidth' | 'authority' | 'raise' | 'fairness' | 'structure' | 'custom',
 *   enabled: boolean,
 *   priority: number,  // 优先级，数字越小越先执行
 *   conditions: [{ field, operator, value }],  // AND 组合
 *   action: { type: 'check' | 'warn' | 'block', message: string },
 *   source: 'system' | 'user' | 'profile',  // 来源
 * }
 *
 * 条件操作符：
 *   eq, ne, gt, gte, lt, lte, in, notIn, contains, between, regex
 */

const RULE_OPERATORS = {
  eq:       (a, b) => a == b,
  ne:       (a, b) => a != b,
  gt:       (a, b) => Number(a) > Number(b),
  gte:      (a, b) => Number(a) >= Number(b),
  lt:       (a, b) => Number(a) < Number(b),
  lte:      (a, b) => Number(a) <= Number(b),
  in:       (a, b) => Array.isArray(b) && b.includes(a),
  notIn:    (a, b) => Array.isArray(b) && !b.includes(a),
  contains: (a, b) => String(a).includes(String(b)),
  between:  (a, b) => Array.isArray(b) && b.length === 2 && Number(a) >= Number(b[0]) && Number(a) <= Number(b[1]),
  regex:    (a, b) => new RegExp(b).test(String(a)),
};

const RULE_CATEGORIES = {
  bandwidth:  { label: '带宽', icon: '📏' },
  authority:  { label: '审批权限', icon: '🔑' },
  raise:      { label: '涨幅', icon: '📈' },
  fairness:   { label: '公平性', icon: '⚖️' },
  structure:  { label: '结构', icon: '🏗️' },
  custom:     { label: '自定义', icon: '✏️' },
};

const SEVERITY_ORDER = { block: 0, warn: 1, check: 2 };

/**
 * 评估单个条件
 */
function evaluateCondition(condition, context) {
  const { field, operator, value } = condition;
  
  // 从 context 中获取字段值（支持点号路径，如 offer.baseSalary）
  const fieldValue = getNestedValue(context, field);
  
  const op = RULE_OPERATORS[operator];
  if (!op) return false;
  
  return op(fieldValue, value);
}

/**
 * 获取嵌套对象的值
 */
function getNestedValue(obj, path) {
  if (!path || !obj) return undefined;
  const parts = String(path).split('.');
  let current = obj;
  for (const part of parts) {
    if (current === null || current === undefined) return undefined;
    current = current[part];
  }
  return current;
}

/**
 * 执行单条规则
 * @returns {{ passed: boolean, rule: Object, message: string, severity: string }}
 */
function executeRule(rule, context) {
  if (!rule.enabled) return null;
  
  // 所有条件必须满足（AND 逻辑）
  const allMatch = (rule.conditions || []).every(cond => evaluateCondition(cond, context));
  
  if (allMatch) {
    // 条件命中 → 触发规则动作
    return {
      passed: false,
      ruleId: rule.id,
      ruleName: rule.name,
      category: rule.category,
      severity: rule.action?.type || 'check',
      message: rule.action?.message || rule.name,
      source: rule.source || 'system',
    };
  }
  
  return { passed: true, ruleId: rule.id, ruleName: rule.name, category: rule.category, severity: 'pass', message: '', source: rule.source };
}

/**
 * 批量执行规则集
 * @param {Array} rules - 规则列表
 * @param {Object} context - 上下文（包含 offer 数据、画像数据等）
 * @returns {{ results: Array, summary: { pass: number, check: number, warn: number, block: number } }}
 */
function executeRules(rules, context) {
  // 按优先级排序
  const sorted = [...rules].sort((a, b) => (a.priority || 99) - (b.priority || 99));
  
  const results = [];
  const summary = { pass: 0, check: 0, warn: 0, block: 0 };
  
  for (const rule of sorted) {
    if (!rule.enabled) continue;
    const result = executeRule(rule, context);
    if (!result) continue;
    
    results.push(result);
    if (result.passed) {
      summary.pass++;
    } else {
      summary[result.severity] = (summary[result.severity] || 0) + 1;
    }
  }
  
  // 按严重程度排序（block > warn > check > pass）
  results.sort((a, b) => {
    if (a.passed && !b.passed) return 1;
    if (!a.passed && b.passed) return -1;
    return (SEVERITY_ORDER[a.severity] || 99) - (SEVERITY_ORDER[b.severity] || 99);
  });
  
  return { results, summary };
}

/**
 * 分类汇总规则执行结果
 * @returns {{ compliance: Array, risk: Array, violation: Array }}
 */
function categorizeResults(results) {
  const compliance = results.filter(r => r.passed);           // 合规（绿）
  const risk = results.filter(r => !r.passed && (r.severity === 'check' || r.severity === 'warn'));  // 风险（黄）
  const violation = results.filter(r => !r.passed && r.severity === 'block');  // 违规（红）
  
  return { compliance, risk, violation };
}

/**
 * 创建新规则的默认模板
 */
function createRuleTemplate(scope) {
  return {
    id: generateId(),
    name: '',
    scope: scope || 'personal',
    category: 'custom',
    enabled: true,
    priority: 50,
    conditions: [{ field: '', operator: 'eq', value: '' }],
    action: { type: 'check', message: '' },
    source: 'user',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}
