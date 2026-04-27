/* ======================================================================
 * Offer审批助手 · 画像 CRUD + 用户标注
 * 管理画像条目的存储、查询、编辑、状态标注
 * ====================================================================== */

/**
 * 保存画像学习结果到 IndexedDB
 * @param {ProfileInsight[]} insights
 */
async function saveProfileInsights(insights) {
  // 先清空旧画像
  await dbClear(STORES.profiles);
  for (const insight of insights) {
    await dbPut(STORES.profiles, insight);
  }
}

/**
 * 加载所有画像条目
 */
async function loadAllProfiles() {
  return await dbGetAll(STORES.profiles);
}

/**
 * 按切片查询画像
 * @param {Object} sliceKey - { country, level, channel, jobFamily }
 */
async function loadProfilesBySlice(sliceKey) {
  const all = await loadAllProfiles();
  return all.filter(p =>
    p.sliceKey.country === sliceKey.country &&
    p.sliceKey.level === sliceKey.level &&
    p.sliceKey.channel === sliceKey.channel &&
    p.sliceKey.jobFamily === sliceKey.jobFamily
  );
}

/**
 * 更新画像条目的用户标注状态
 * @param {string} id
 * @param {string} status - 'suggested' | 'confirmed' | 'edited' | 'rejected'
 */
async function updateProfileStatus(id, status) {
  const profile = await dbGet(STORES.profiles, id);
  if (profile) {
    profile.userStatus = status;
    await dbPut(STORES.profiles, profile);
  }
}

/**
 * 编辑画像条目的 pattern
 * @param {string} id
 * @param {Object} newPattern
 */
async function editProfilePattern(id, newPattern) {
  const profile = await dbGet(STORES.profiles, id);
  if (profile) {
    profile.pattern = { ...profile.pattern, ...newPattern };
    profile.userStatus = 'edited';
    await dbPut(STORES.profiles, profile);
  }
}

/**
 * 删除画像条目
 */
async function deleteProfile(id) {
  await dbDelete(STORES.profiles, id);
}

/**
 * 将画像规律转为个人规则
 */
async function convertToPersonalRule(insight) {
  if (insight.type !== 'conditional') return;

  const rule = createRuleTemplate('personal');
  rule.name = '从画像转化：' + insight.dimensionLabel;
  rule.category = 'custom';
  rule.source = 'profile';
  rule.conditions = [{
    field: 'context.' + insight.dimension.split('.')[0],
    operator: 'eq',
    value: insight.pattern.condition ? insight.pattern.condition.split('= ')[1] : '',
  }];
  rule.action = {
    type: 'check',
    message: insight.dimensionLabel + '（效应: ' + ((insight.pattern.effect || 0) * 100).toFixed(1) + '%）',
  };

  await saveRule(rule);
  await updateProfileStatus(insight.id, 'confirmed');
  return rule;
}

/**
 * 按切片分组统计画像概览
 * @returns {Array<{sliceKey, insightCount, numericCount, conditionalCount, avgConfidence}>}
 */
async function getProfileOverview() {
  const all = await loadAllProfiles();
  const groups = {};

  for (const p of all) {
    const key = JSON.stringify(p.sliceKey);
    if (!groups[key]) {
      groups[key] = { sliceKey: p.sliceKey, items: [], numericCount: 0, conditionalCount: 0, totalConfidence: 0 };
    }
    groups[key].items.push(p);
    if (p.type === 'numeric') groups[key].numericCount++;
    else groups[key].conditionalCount++;
    groups[key].totalConfidence += p.confidence;
  }

  return Object.values(groups).map(g => ({
    sliceKey: g.sliceKey,
    insightCount: g.items.length,
    numericCount: g.numericCount,
    conditionalCount: g.conditionalCount,
    avgConfidence: g.items.length > 0 ? Math.round((g.totalConfidence / g.items.length) * 100) / 100 : 0,
    sampleIds: [...new Set(g.items.flatMap(i => i.sourceOfferIds || []))],
  }));
}
