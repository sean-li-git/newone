/* 因子3：外部竞争 — 来源公司/在谈offer/Counter Offer 三子维度 */
const FactorCompetition = {
  id: 'competition',
  name: '外部竞争',
  run(offer, profiles, context, weights) {
    let totalPremium = 0;
    const reasons = [];
    const src = context.sourceCompany || '';
    const tiers = weights.competition.tierPremium || {};
    // 检查是否在竞企清单
    let matchedTier = null;
    for (const [tier, premium] of Object.entries(tiers)) {
      if (tier === src || (context.competitorTier && context.competitorTier === tier)) {
        matchedTier = tier;
        totalPremium += premium;
        reasons.push('来源「' + tier + '」溢价 +' + (premium * 100).toFixed(0) + '%');
        break;
      }
    }
    // 画像中的来源公司溢价
    if (!matchedTier && src) {
      const srcProfile = profiles.find(p => p.type === 'conditional' && p.dimension.startsWith('来源公司.' + src));
      if (srcProfile && srcProfile.pattern.effect) {
        totalPremium += srcProfile.pattern.effect;
        reasons.push('画像来源溢价 ' + (srcProfile.pattern.effect >= 0 ? '+' : '') + (srcProfile.pattern.effect * 100).toFixed(1) + '%');
      }
    }
    // Counter Offer
    if (context.counterOffer && context.counterOffer > 0) {
      const counterPremium = 0.05;
      totalPremium += counterPremium;
      reasons.push('有 Counter Offer（' + formatNumber(context.counterOffer) + '），+5%');
    }
    const base = offer.baseSalary || 0;
    const delta = Math.round(base * 12 * totalPremium);
    return {
      delta: { conservative: Math.round(delta * 0.5), standard: delta, aggressive: Math.round(delta * 1.3) },
      reason: reasons.length > 0 ? reasons.join('；') : '无竞争溢价因子',
      evidence: { sampleIds: [], note: '来源: ' + (src || '未知') },
    };
  }
};
