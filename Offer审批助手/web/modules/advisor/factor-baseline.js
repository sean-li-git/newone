/* 因子1：基线水位 — 从画像 4 维切片查询 P50 作为基准 */
const FactorBaseline = {
  id: 'baseline',
  name: '基线水位',
  run(offer, profiles, context, weights) {
    if (!weights.baseline.enabled) return null;
    const slice = profiles.filter(p =>
      p.type === 'numeric' && p.sliceKey.country === offer.country &&
      p.sliceKey.level === offer.level && p.sliceKey.channel === offer.channel &&
      p.sliceKey.jobFamily === offer.jobFamily
    );
    const baseP = slice.find(p => p.dimension === 'baseSalary');
    const cashP = slice.find(p => p.dimension === 'cashIncome');
    const annualP = slice.find(p => p.dimension === 'annualIncome');
    if (!baseP) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '无匹配画像基线', evidence: { sampleIds: [], note: '未找到匹配的4维切片画像' } };
    const baseMedian = baseP.pattern.median;
    const diff = offer.baseSalary - baseMedian;
    const pct = baseMedian > 0 ? diff / baseMedian : 0;
    return {
      delta: { conservative: 0, standard: 0, aggressive: 0 },
      reason: 'base 中位数 ' + formatNumber(baseMedian) + '，当前 ' + formatNumber(offer.baseSalary) + '（' + (pct >= 0 ? '+' : '') + (pct * 100).toFixed(1) + '%）',
      evidence: { sampleIds: baseP.sourceOfferIds || [], note: '来自 ' + baseP.sampleCount + ' 个样本的 P50', baseline: { base: baseMedian, cash: cashP ? cashP.pattern.median : null, annual: annualP ? annualP.pattern.median : null } },
    };
  }
};
