/* 因子2：候选人增幅 — 现薪→目标 jump% 上下限约束 */
const FactorRaise = {
  id: 'raise',
  name: '候选人增幅',
  run(offer, profiles, context, weights) {
    const currentSalary = context.currentSalary || 0;
    if (currentSalary <= 0) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '未提供候选人现薪，跳过增幅分析', evidence: { sampleIds: [], note: '' } };
    const targetCash = offer.cashIncome || (offer.baseSalary * 12);
    const raisePct = (targetCash - currentSalary) / currentSalary;
    const cap = weights.raise.defaultCap || 0.35;
    const learned = weights.raise.learnedMedian;
    let hardConstraint = null;
    if (raisePct > 0.50) {
      hardConstraint = { type: 'ceiling', value: currentSalary * 1.50, message: '增幅超过 50% 红线' };
    } else if (raisePct > cap) {
      hardConstraint = { type: 'ceiling', value: currentSalary * (1 + cap), message: '增幅超过 ' + (cap * 100).toFixed(0) + '% 上限' };
    }
    return {
      delta: { conservative: 0, standard: 0, aggressive: 0 },
      reason: '现薪 ' + formatNumber(currentSalary) + ' → 目标 ' + formatNumber(targetCash) + '（增幅 ' + (raisePct * 100).toFixed(1) + '%' + (learned ? '，历史中位 ' + (learned * 100).toFixed(1) + '%' : '') + '）',
      evidence: { sampleIds: [], note: '增幅上限 ' + (cap * 100).toFixed(0) + '%' },
      hardConstraint,
    };
  }
};
