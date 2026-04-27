/* 因子7：岗位稀缺度 — 命中稀缺岗位清单 → 整档上浮 */
const FactorRarity = {
  id: 'rarity',
  name: '岗位稀缺度',
  run(offer, profiles, context, weights) {
    const positions = weights.rarity.positions || [];
    const uplift = weights.rarity.uplift || 0.10;
    const jobFamily = offer.jobFamily || '';
    const channel = offer.channel || '';
    const isRare = positions.some(p => p === jobFamily || p === channel || (jobFamily + '-' + channel).includes(p));
    if (!isRare) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '非稀缺岗位', evidence: { sampleIds: [], note: '' } };
    const base = offer.baseSalary || 0;
    const delta = Math.round(base * 12 * uplift);
    return {
      delta: { conservative: Math.round(delta * 0.7), standard: delta, aggressive: Math.round(delta * 1.5) },
      reason: '命中稀缺岗位清单，整档上浮 +' + (uplift * 100).toFixed(0) + '%',
      evidence: { sampleIds: [], note: '匹配: ' + jobFamily + ' / ' + channel },
    };
  }
};
