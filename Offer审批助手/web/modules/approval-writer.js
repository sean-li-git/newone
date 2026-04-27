/* ======================================================================
 * Offer审批助手 · 审批说明生成器
 * 三段式结构：总包 / 能力匹配 / 预算
 * ====================================================================== */

/**
 * 生成审批说明（基础版 — 模板拼装）
 * @param {Object} offer - Offer 对象
 * @param {Object} pkg - 采用的推荐方案 (recommendation.standard 等)
 * @param {Object} context - 候选人上下文
 * @param {Array} factors - 因子分析结果
 * @param {Array} risks - 风险提示
 * @returns {Object} { section1, section2, section3, fullText, markdown }
 */
function generateApprovalStatement(offer, pkg, context, factors, risks) {
  // 第一段：总包方案
  const section1 = buildPackageSection(offer, pkg, context);

  // 第二段：能力匹配
  const section2 = buildAbilitySection(offer, context, factors);

  // 第三段：预算与风险
  const section3 = buildBudgetSection(offer, pkg, context, risks);

  const fullText = section1 + '\n\n' + section2 + '\n\n' + section3;
  const markdown = '## 一、薪酬方案\n\n' + section1 + '\n\n## 二、能力匹配\n\n' + section2 + '\n\n## 三、预算与风险\n\n' + section3;

  return { section1, section2, section3, fullText, markdown };
}

function buildPackageSection(offer, pkg, context) {
  const lines = [];
  lines.push('**候选人信息**：' + (offer.country || '-') + ' · ' + (offer.level || '-') + ' · ' + (offer.channel || '-') + ' · ' + (offer.jobFamily || '-'));
  if (context.sourceCompany) lines.push('**来源公司**：' + context.sourceCompany);
  lines.push('');
  lines.push('**推荐薪酬方案（' + ({ conservative: '保守', standard: '标准', aggressive: '进取' }[pkg.tier] || pkg.tier) + '档）**：');
  lines.push('- 基本月薪：' + formatNumber(pkg.baseSalary));
  lines.push('- 绩效（月均）：' + formatNumber(pkg.perfMonthly));
  lines.push('- 年终奖：' + formatNumber(pkg.annualBonus));
  lines.push('- **年现金收入：' + formatNumber(pkg.cashIncome) + '**');
  if (pkg.equityAnnual > 0) lines.push('- 入职授予（年化）：' + formatNumber(pkg.equityAnnual));
  lines.push('- **年总收入：' + formatNumber(pkg.totalAnnual) + '**');
  if (pkg.signStock > 0) lines.push('- 签字股票：' + formatNumber(pkg.signStock));
  if (pkg.signBonus > 0) lines.push('- 签字费：' + formatNumber(pkg.signBonus));
  if (pkg.relocation > 0) lines.push('- 安家费：' + formatNumber(pkg.relocation));
  if (context.currentSalary > 0) {
    const raise = ((pkg.totalAnnual - context.currentSalary) / context.currentSalary * 100).toFixed(1);
    lines.push('');
    lines.push('**增幅**：' + raise + '%（现薪 ' + formatNumber(context.currentSalary) + '）');
  }
  return lines.join('\n');
}

function buildAbilitySection(offer, context, factors) {
  const lines = [];
  const tags = context.abilityTags || [];
  if (tags.length > 0) {
    lines.push('**能力标签**：' + tags.join('、'));
  }
  // 从因子中提取关键信息
  const baseline = factors.find(f => f.id === 'baseline');
  if (baseline && baseline.evidence && baseline.evidence.baseline && baseline.evidence.baseline.base) {
    lines.push('**画像参照**：同切片 base 中位数 ' + formatNumber(baseline.evidence.baseline.base) + '，当前方案 base ' + formatNumber(offer.baseSalary));
  }
  const competition = factors.find(f => f.id === 'competition');
  if (competition && competition.reason && competition.reason !== '无竞争溢价因子') {
    lines.push('**竞争因素**：' + competition.reason);
  }
  const ability = factors.find(f => f.id === 'ability');
  if (ability && ability.reason && ability.reason !== '未标注能力标签' && ability.reason !== '标签无匹配加成') {
    lines.push('**能力加成**：' + ability.reason);
  }
  if (lines.length === 0) lines.push('候选人能力评估信息待补充。');
  return lines.join('\n');
}

function buildBudgetSection(offer, pkg, context, risks) {
  const lines = [];
  if (risks && risks.length > 0) {
    lines.push('**风险提示**：');
    for (const r of risks) {
      lines.push('- ' + (r.level === 'danger' ? '🚫' : '⚠️') + ' ' + r.message);
    }
    lines.push('');
  }
  lines.push('**预算影响**：');
  lines.push('- 年度新增现金成本：' + formatNumber(pkg.cashIncome));
  if (pkg.signBonus > 0 || pkg.signStock > 0 || pkg.relocation > 0) {
    const oneTime = (pkg.signBonus || 0) + (pkg.signStock || 0) + (pkg.relocation || 0);
    lines.push('- 一次性费用总计：' + formatNumber(oneTime));
  }
  lines.push('');
  lines.push('**审批建议**：' + (risks.some(r => r.level === 'danger') ? '建议慎重审批，存在红线风险。' : risks.length > 0 ? '建议关注风险项后审批通过。' : '无显著风险，建议审批通过。'));
  return lines.join('\n');
}

/**
 * 使用 AI 润色（需先脱敏）
 * @returns {Promise<string>} 润色后的文案
 */
async function aiPolishStatement(statement, offer) {
  // 先脱敏
  const desensitized = desensitizeForAI(statement, offer);
  // 调用 AI
  const result = await callAIProvider('润色以下审批说明，保持三段式结构，语言更专业流畅：\n\n' + desensitized);
  return result;
}
