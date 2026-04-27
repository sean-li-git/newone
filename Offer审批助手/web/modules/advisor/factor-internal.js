/* 因子4：内部平衡 — 与在司员工薪酬比对，硬约束天花板/地板 */
const FactorInternal = {
  id: 'internal',
  name: '内部平衡',
  run(offer, profiles, context, weights) {
    const ceiling = weights.internal.ceilingRatio || 1.2;
    const floor = weights.internal.floorRatio || 0.85;
    // 从画像取同切片 base 中位数作为内部参照
    const baseProfile = profiles.find(p =>
      p.type === 'numeric' && p.dimension === 'baseSalary' &&
      p.sliceKey.country === offer.country && p.sliceKey.level === offer.level &&
      p.sliceKey.channel === offer.channel && p.sliceKey.jobFamily === offer.jobFamily
    );
    if (!baseProfile) return { delta: { conservative: 0, standard: 0, aggressive: 0 }, reason: '无内部参照数据', evidence: { sampleIds: [], note: '' } };
    const internalMedian = baseProfile.pattern.median;
    const ceilingVal = Math.round(internalMedian * ceiling);
    const floorVal = Math.round(internalMedian * floor);
    let hardConstraint = null;
    if (offer.baseSalary > ceilingVal) {
      hardConstraint = { type: 'ceiling', value: ceilingVal, message: 'base ' + formatNumber(offer.baseSalary) + ' 超过内部天花板 ' + formatNumber(ceilingVal) + '（中位数×' + ceiling + '）' };
    } else if (offer.baseSalary < floorVal) {
      hardConstraint = { type: 'floor', value: floorVal, message: 'base ' + formatNumber(offer.baseSalary) + ' 低于内部地板 ' + formatNumber(floorVal) + '（中位数×' + floor + '）' };
    }
    return {
      delta: { conservative: 0, standard: 0, aggressive: 0 },
      reason: '内部中位数 ' + formatNumber(internalMedian) + '，安全区间 ' + formatNumber(floorVal) + '~' + formatNumber(ceilingVal),
      evidence: { sampleIds: baseProfile.sourceOfferIds || [], note: '天花板系数 ' + ceiling + '×，地板系数 ' + floor + '×' },
      hardConstraint,
    };
  }
};
