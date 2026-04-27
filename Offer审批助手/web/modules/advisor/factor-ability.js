/* 因子6：能力标签 — 用户完全自定义，单项加成 + 累计封顶 */
const FactorAbility = {
  id: 'ability',
  name: '能力标签',
  run(offer, profiles, context, weights) {
    const tags = context.abilityTags || [];
    if (tags.length === 0) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '未标注能力标签', evidence: { sampleIds: [], note: '' } };
    const perTag = weights.ability.perTag || {};
    const cap = weights.ability.capPercent || 0.20;
    let totalPct = 0;
    const applied = [];
    for (const tag of tags) {
      const pct = perTag[tag];
      if (pct && pct > 0) {
        totalPct += pct;
        applied.push(tag + ' +' + (pct * 100).toFixed(0) + '%');
      }
    }
    if (totalPct > cap) totalPct = cap;
    const base = offer.baseSalary || 0;
    const delta = Math.round(base * 12 * totalPct);
    return {
      delta: { conservative: Math.round(delta * 0.5), standard: delta, aggressive: Math.round(delta * 1.3) },
      reason: applied.length > 0 ? applied.join('、') + (totalPct >= cap ? '（已封顶 ' + (cap * 100).toFixed(0) + '%）' : '') : '标签无匹配加成',
      evidence: { sampleIds: [], note: '封顶 ' + (cap * 100).toFixed(0) + '%' },
    };
  }
};
