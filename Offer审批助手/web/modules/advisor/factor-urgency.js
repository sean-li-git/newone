/* 因子5：紧迫度 — 业务紧迫度动总包上浮，入职时间紧迫度动一次性费用 */
const FactorUrgency = {
  id: 'urgency',
  name: '紧迫度',
  run(offer, profiles, context, weights) {
    const bizUrg = context.bizUrgency || '正常';
    const onbUrg = context.onboardUrgency || '1-3个月';
    const bizRate = (weights.urgency.business || {})[bizUrg] || 0;
    const onbRate = (weights.urgency.onboarding || {})[onbUrg] || 0;
    const base = offer.baseSalary || 0;
    const annualBase = base * 12;
    // 业务紧迫度 → 总包上浮
    const bizDelta = Math.round(annualBase * bizRate);
    // 入职时间紧迫 → 一次性费用杠杆（不动年薪）
    const onbDelta = Math.round((offer.signBonus || 0) * onbRate + (offer.relocation || 0) * onbRate);
    const totalDelta = bizDelta + onbDelta;
    const reasons = [];
    if (bizRate !== 0) reasons.push('业务「' + bizUrg + '」→ 总包 ' + (bizRate >= 0 ? '+' : '') + (bizRate * 100).toFixed(0) + '%');
    if (onbRate !== 0) reasons.push('入职「' + onbUrg + '」→ 一次性 ' + (onbRate >= 0 ? '+' : '') + (onbRate * 100).toFixed(0) + '%');
    return {
      delta: { conservative: Math.round(totalDelta * 0.6), standard: totalDelta, aggressive: Math.round(totalDelta * 1.4) },
      reason: reasons.length > 0 ? reasons.join('；') : '紧迫度无调整',
      evidence: { sampleIds: [], note: '业务: ' + bizUrg + '，入职: ' + onbUrg },
    };
  }
};
