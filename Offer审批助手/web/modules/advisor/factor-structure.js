/* 因子8：薪酬结构偏好 — 从历史学习 base/股票/签字费比例偏好，不动总包只调结构 */
const FactorStructure = {
  id: 'structure',
  name: '薪酬结构偏好',
  run(offer, profiles, context, weights) {
    const prefBasePct = weights.structure.basePct || 0.55;
    const prefStockPct = weights.structure.stockPct || 0.25;
    const prefSignPct = weights.structure.signBonusPct || 0.10;
    // 从画像获取实际偏好
    const structProfile = profiles.find(p =>
      p.type === 'conditional' && p.dimension === '结构偏好' &&
      p.sliceKey.country === offer.country && p.sliceKey.level === offer.level &&
      p.sliceKey.channel === offer.channel && p.sliceKey.jobFamily === offer.jobFamily
    );
    const actualBasePct = structProfile && structProfile.pattern.basePctMedian ? structProfile.pattern.basePctMedian : prefBasePct;
    const totalComp = offer.annualIncome || (offer.cashIncome || 0) + (offer.equityAnnual || 0);
    if (totalComp <= 0) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '无年收入数据，跳过结构分析', evidence: { sampleIds: [], note: '' } };
    const currentBasePct = (offer.baseSalary * 12) / totalComp;
    const deviation = Math.abs(currentBasePct - actualBasePct);
    return {
      delta: { conservative: 0, standard: 0, aggressive: 0 }, // 结构因子不改总包
      reason: 'Base 占比 ' + (currentBasePct * 100).toFixed(1) + '%（偏好 ' + (actualBasePct * 100).toFixed(1) + '%' + (deviation > 0.05 ? '，偏差 ' + (deviation * 100).toFixed(1) + '%' : '，匹配') + '）',
      evidence: { sampleIds: structProfile ? (structProfile.sourceOfferIds || []) : [], note: '建议结构: Base ' + (actualBasePct * 100).toFixed(0) + '% / 股票 ' + (prefStockPct * 100).toFixed(0) + '% / 签字费 ' + (prefSignPct * 100).toFixed(0) + '%' },
      structureSuggestion: {
        basePct: actualBasePct,
        stockPct: prefStockPct,
        signBonusPct: prefSignPct,
        suggestedBase: Math.round(totalComp * actualBasePct / 12),
      },
    };
  }
};
