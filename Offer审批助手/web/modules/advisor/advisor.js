/* ======================================================================
 * Offer审批助手 · 推荐引擎主流程
 * Pipeline 架构：8 因子依次执行，合成三档方案
 * ====================================================================== */

const ADVISOR_FACTORS = [
  FactorBaseline,
  FactorRaise,
  FactorCompetition,
  FactorInternal,
  FactorUrgency,
  FactorAbility,
  FactorRarity,
  FactorStructure,
];

/**
 * 运行推荐引擎
 * @param {Object} offer - 当前 Offer 对象（14 字段）
 * @param {Object} context - 附加上下文 { currentSalary, sourceCompany, competitorTier, counterOffer, bizUrgency, onboardUrgency, abilityTags }
 * @returns {Promise<{factors: Array, recommendation: {conservative, standard, aggressive}, risks: Array}>}
 */
async function runAdvisor(offer, context) {
  const profiles = await loadAllProfiles();
  const weights = await loadFactorWeights();
  const competitors = await loadCompetitorList();
  const rarePositions = await loadRarePositions();

  // 补充 weights 中的清单
  weights.rarity.positions = rarePositions;
  // 标记来源公司的竞企层级
  if (context.sourceCompany && competitors.includes(context.sourceCompany)) {
    context.competitorTier = '一线竞对';
  }

  const factorResults = [];
  const hardConstraints = [];
  let totalDelta = { conservative: 0, standard: 0, aggressive: 0 };

  for (const factor of ADVISOR_FACTORS) {
    try {
      const result = factor.run(offer, profiles, context, weights);
      if (!result) continue;
      factorResults.push({ id: factor.id, name: factor.name, ...result });
      totalDelta.conservative += result.delta.conservative;
      totalDelta.standard += result.delta.standard;
      totalDelta.aggressive += result.delta.aggressive;
      if (result.hardConstraint) hardConstraints.push(result.hardConstraint);
    } catch (err) {
      factorResults.push({ id: factor.id, name: factor.name, delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '执行异常: ' + err.message, evidence: { sampleIds: [], note: '' } });
    }
  }

  // 基础年收入
  const baseAnnual = offer.annualIncome || offer.cashIncome || (offer.baseSalary * 12);

  // 合成三档方案
  const recommendation = {
    conservative: buildPackage(offer, baseAnnual + totalDelta.conservative, hardConstraints, 'conservative'),
    standard: buildPackage(offer, baseAnnual + totalDelta.standard, hardConstraints, 'standard'),
    aggressive: buildPackage(offer, baseAnnual + totalDelta.aggressive, hardConstraints, 'aggressive'),
  };

  // 风险检测
  const risks = detectRisks(offer, recommendation, context, factorResults);

  return { factors: factorResults, recommendation, risks, hardConstraints };
}

/**
 * 构建薪酬包明细
 */
function buildPackage(offer, targetAnnual, hardConstraints, tier) {
  // 应用硬约束
  let adjustedBase = offer.baseSalary;
  for (const hc of hardConstraints) {
    if (hc.type === 'ceiling') adjustedBase = Math.min(adjustedBase, Math.round(hc.value / 12));
    if (hc.type === 'floor') adjustedBase = Math.max(adjustedBase, Math.round(hc.value / 12));
  }

  const base = adjustedBase;
  const perf = offer.perfMonthly || 0;
  const bonus = offer.annualBonus || 0;
  const allowance = offer.allowance || 0;
  const cashIncome = base * 12 + perf * 12 + bonus + allowance;
  const equity = Math.max(0, targetAnnual - cashIncome);

  return {
    tier,
    totalAnnual: Math.round(targetAnnual),
    baseSalary: base,
    perfMonthly: perf,
    annualBonus: bonus,
    allowance: allowance,
    cashIncome: Math.round(cashIncome),
    equityAnnual: Math.round(equity),
    signStock: offer.signStock || 0,
    signBonus: offer.signBonus || 0,
    relocation: offer.relocation || 0,
  };
}

/**
 * 风险检测
 */
function detectRisks(offer, recommendation, context, factorResults) {
  const risks = [];
  const std = recommendation.standard;

  // 倒挂检测
  if (context.currentSalary > 0 && std.cashIncome < context.currentSalary * 0.95) {
    risks.push({ level: 'danger', message: '现金收入低于候选人现薪，可能倒挂' });
  }

  // 增幅超限
  if (context.currentSalary > 0) {
    const raise = (std.totalAnnual - context.currentSalary) / context.currentSalary;
    if (raise > 0.50) risks.push({ level: 'danger', message: '增幅超过 50%，触发审批红线' });
    else if (raise > 0.35) risks.push({ level: 'warning', message: '增幅超过 35%，需额外审批' });
  }

  // 硬约束
  for (const fr of factorResults) {
    if (fr.hardConstraint) {
      risks.push({ level: fr.hardConstraint.type === 'ceiling' ? 'danger' : 'warning', message: fr.hardConstraint.message });
    }
  }

  // 偏低检测
  const baseline = factorResults.find(f => f.id === 'baseline');
  if (baseline && baseline.evidence && baseline.evidence.baseline && baseline.evidence.baseline.base) {
    const medianBase = baseline.evidence.baseline.base;
    if (offer.baseSalary < medianBase * 0.85) {
      risks.push({ level: 'warning', message: 'base 低于画像 P50 的 85%（' + formatNumber(medianBase) + '），偏低风险' });
    }
  }

  return risks;
}
