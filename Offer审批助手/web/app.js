/* ======================================================================
 * Offer审批助手 · 核心逻辑（路由 + 全局状态 + Dashboard + 薪酬导入页）
 * 纯前端，全部本地处理，无任何网络请求
 * ====================================================================== */

const APP_VERSION = 'v1.0';

// ========== 全局状态 ==========
const AppState = {
  currentPage: 'dashboard',
  // 薪酬导入
  importFile: null,
  importResult: null, // parseSalaryExcel 的返回值
  currentOffers: [],  // 当前正在处理的 Offer 列表
  currentValidations: [], // 对应的校验结果
  // 配置
  amortizationYears: 2, // 签字费/签字股票分摊年限（规则库可配置）
};

// ========== 初始化 ==========
document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('app-version').textContent = APP_VERSION;
  
  // 初始化存储
  await initDB();
  
  // 加载配置
  AppState.amortizationYears = lsGet('amortizationYears', 2);
  
  // 初始化路由
  initRouter();
  
  // 初始化各页面
  initDashboard();
  initImportPage();
  initRulesPage();
  initSampleWizard();
  initProfilePage();
  initCheckupPage();
  initApprovalPage();
  
  // 刷新 Dashboard 数据
  await refreshDashboard();
});

// ========== 工具函数 ==========
function showLoading(text) {
  document.getElementById('loading-text').textContent = text || '处理中...';
  document.getElementById('loading').classList.add('show');
}

function hideLoading() {
  document.getElementById('loading').classList.remove('show');
}

function showToast(message, type) {
  type = type || 'success';
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = 'toast toast-' + type;
  const icons = { success: '✅', error: '❌', warning: '⚠️' };
  toast.innerHTML = '<span>' + (icons[type] || '') + '</span><span>' + escapeHtml(message) + '</span>';
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'toastOut 0.3s forwards';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, m => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m]));
}

function formatNumber(n) {
  if (n === null || n === undefined || n === '') return '-';
  const num = Number(n);
  if (isNaN(num)) return String(n);
  return num.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// ========== 路由 ==========
function initRouter() {
  // 侧边栏导航
  document.querySelectorAll('.nav-item[data-page]').forEach(item => {
    item.addEventListener('click', () => navigateTo(item.dataset.page));
  });
  
  // 快速入口 / data-goto 链接
  document.querySelectorAll('[data-goto]').forEach(el => {
    el.addEventListener('click', (e) => {
      e.preventDefault();
      navigateTo(el.dataset.goto);
    });
  });
}

function navigateTo(page) {
  AppState.currentPage = page;
  
  // 更新侧边栏高亮
  document.querySelectorAll('.nav-item').forEach(item => {
    item.classList.toggle('active', item.dataset.page === page);
  });
  
  // 切换页面
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  const pageEl = document.getElementById('page-' + page);
  if (pageEl) pageEl.classList.add('active');
  
  // 更新顶栏标题
  const titles = {
    dashboard: ['工作台', 'Offer审批提效工具'],
    import: ['薪酬导入', '上传薪酬 Excel，14 字段自动解析与校验'],
    checkup: ['Offer 体检', '合规检查 + 8 因子推荐 + 三档方案'],
    rules: ['规则库', '统一规则 + 个人规则 + 配置表'],
    profile: ['习惯画像', '从历史 Offer 学习审批偏好'],
    approval: ['审批说明', '一键生成三段式审批说明'],
  };
  const [title, subtitle] = titles[page] || [page, ''];
  document.getElementById('topbar-title').textContent = title;
  document.getElementById('topbar-subtitle').textContent = subtitle;
  
  // 页面切换时的额外初始化
  if (page === 'checkup') updateCheckupEmpty();
  if (page === 'approval') updateApprovalState();
}

// ========== Dashboard ==========
function initDashboard() {
  // 备份/还原按钮
  document.getElementById('btn-export-backup').addEventListener('click', handleExportBackup);
  document.getElementById('btn-import-backup').addEventListener('click', () => document.getElementById('backup-import-input').click());
  document.getElementById('backup-import-input').addEventListener('change', handleImportBackup);
}

async function refreshDashboard() {
  try {
    // 统计数据
    const offerCount = await dbCount(STORES.offers);
    const historyCount = await dbCount(STORES.historyOffers);
    const ruleCount = await dbCount(STORES.rules);
    
    document.getElementById('stat-pending').textContent = offerCount;
    document.getElementById('stat-done').textContent = '0'; // 后续从 activities 统计
    document.getElementById('stat-samples').textContent = historyCount;
    document.getElementById('stat-rules').textContent = ruleCount;
    
    // 活动流
    await renderActivityFeed();
  } catch (e) {
    // 静默处理首次初始化错误
  }
}

async function renderActivityFeed() {
  const container = document.getElementById('dashboard-activity');
  const activities = await dbGetAll(STORES.activities);
  
  if (!activities || activities.length === 0) {
    container.innerHTML = '<div class="empty-state"><div class="es-icon">📭</div><div class="es-title">暂无活动记录</div><div class="es-desc">导入第一份 Offer 开始使用，所有操作记录将在这里展示</div></div>';
    return;
  }
  
  // 倒序取最近 10 条
  const recent = activities.sort((a, b) => b.timestamp.localeCompare(a.timestamp)).slice(0, 10);
  
  let html = '<ul class="activity-feed">';
  for (const act of recent) {
    const dotClass = act.type === 'import' ? 'blue' : act.type === 'checkup' ? 'green' : 'amber';
    const time = new Date(act.timestamp).toLocaleString('zh-CN', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
    html += '<li class="activity-item"><span class="activity-dot ' + dotClass + '"></span><span class="activity-text">' + escapeHtml(act.message) + '</span><span class="activity-time">' + time + '</span></li>';
  }
  html += '</ul>';
  container.innerHTML = html;
}

async function handleExportBackup() {
  showLoading('正在导出备份...');
  try {
    const data = await exportAllData();
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const date = new Date().toISOString().slice(0, 10);
    downloadBlob(blob, 'Offer审批助手_备份_' + date + '.json');
    showToast('备份已导出');
  } catch (e) {
    showToast('备份失败：' + e.message, 'error');
  }
  hideLoading();
}

async function handleImportBackup(e) {
  if (!e.target.files.length) return;
  const file = e.target.files[0];
  e.target.value = '';
  
  if (!confirm('导入备份将覆盖当前所有数据，确定继续吗？')) return;
  
  showLoading('正在导入备份...');
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    await importAllData(data);
    await refreshDashboard();
    showToast('备份已还原');
  } catch (e) {
    showToast('还原失败：' + e.message, 'error');
  }
  hideLoading();
}

// ========== 薪酬导入页 ==========
function initImportPage() {
  const dropzone = document.getElementById('import-dropzone');
  const fileInput = document.getElementById('import-file-input');
  
  // 拖拽
  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) await handleImportExcel(e.dataTransfer.files[0]);
  });
  
  // 文件选择
  fileInput.addEventListener('change', async (e) => {
    if (e.target.files.length > 0) await handleImportExcel(e.target.files[0]);
    e.target.value = '';
  });
  
  // 更换文件
  document.getElementById('import-change-file').addEventListener('click', () => {
    document.getElementById('import-upload').classList.remove('hidden');
    document.getElementById('import-validation').classList.add('hidden');
    AppState.importResult = null;
    AppState.currentOffers = [];
    fileInput.click();
  });
  
  // 进入体检
  document.getElementById('import-to-checkup').addEventListener('click', async () => {
    if (AppState.currentOffers.length === 0) {
      showToast('请先导入有效的薪酬数据', 'warning');
      return;
    }
    // 保存到 IndexedDB
    for (const offer of AppState.currentOffers) {
      await dbPut(STORES.offers, offer);
    }
    await logActivity('import', '导入了 ' + AppState.currentOffers.length + ' 条 Offer 薪酬数据');
    await refreshDashboard();
    showToast('已导入 ' + AppState.currentOffers.length + ' 条 Offer');
    navigateTo('checkup');
  });
}

async function handleImportExcel(file) {
  showLoading('正在解析薪酬 Excel...');
  
  try {
    const result = await parseSalaryExcel(file);
    AppState.importFile = file;
    AppState.importResult = result;
    
    if (result.errors.length > 0 && result.offers.length === 0) {
      hideLoading();
      showToast('解析失败：' + result.errors.join('; '), 'error');
      return;
    }
    
    // 执行自动核算与校验
    const validations = [];
    for (const offer of result.offers) {
      const v = validateOffer(offer, AppState.amortizationYears);
      validations.push(v);
    }
    AppState.currentOffers = result.offers;
    AppState.currentValidations = validations;
    
    // 切换到校验视图
    document.getElementById('import-upload').classList.add('hidden');
    document.getElementById('import-validation').classList.remove('hidden');
    
    // 更新文件信息
    document.getElementById('import-file-name').textContent = file.name;
    document.getElementById('import-file-detail').textContent =
      result.sheetName + ' · ' + result.offers.length + ' 条 Offer · ' +
      Object.keys(result.fieldMapping).length + '/' + SALARY_FIELDS.length + ' 个字段已匹配';
    
    // 渲染校验表
    renderValidationTable(result.offers[0], validations[0], result);
    renderComputedComparison(result.offers[0], validations[0]);
    
    if (result.errors.length > 0) {
      showToast('部分字段未匹配：' + result.errors.join('; '), 'warning');
    }
    
    hideLoading();
  } catch (e) {
    hideLoading();
    showToast('解析失败：' + e.message, 'error');
  }
}

function renderValidationTable(offer, validation, result) {
  const tbody = document.getElementById('import-validation-body');
  let html = '';
  let idx = 0;
  
  for (const field of SALARY_FIELDS) {
    idx++;
    const val = offer[field.key];
    const displayVal = field.type === 'number' ? formatNumber(val) : (val || '-');
    
    // 查找是否有该字段的校验问题
    const issue = validation.allIssues.find(i => i.field === field.key);
    let statusBadge = '<span class="badge badge-success">✓ 正常</span>';
    if (issue) {
      if (issue.severity === 'error') statusBadge = '<span class="badge badge-danger">✗ ' + escapeHtml(issue.message) + '</span>';
      else if (issue.severity === 'warning') statusBadge = '<span class="badge badge-warning">⚠ ' + escapeHtml(issue.message) + '</span>';
      else statusBadge = '<span class="badge badge-info">ℹ ' + escapeHtml(issue.message) + '</span>';
    }
    
    // 计算项显示系统计算值
    let computedCol = '-';
    if (field.computed) {
      const cv = validation.computed.find(c => c.key === field.key);
      if (cv) {
        computedCol = formatNumber(cv.computedValue);
        if (!cv.match) {
          statusBadge = '<span class="badge badge-danger">✗ 差异 ' + (cv.diff > 0 ? '+' : '') + formatNumber(cv.diff) + '</span>';
        }
      }
    }
    
    html += '<tr>' +
      '<td style="color:var(--text-dim)">' + idx + '</td>' +
      '<td><span class="badge badge-neutral">' + escapeHtml(field.group) + '</span></td>' +
      '<td style="font-weight:600">' + escapeHtml(field.label) + '</td>' +
      '<td>' + escapeHtml(displayVal) + '</td>' +
      '<td>' + (field.computed ? computedCol : '<span style="color:var(--text-dim)">-</span>') + '</td>' +
      '<td>' + statusBadge + '</td>' +
      '</tr>';
  }
  
  tbody.innerHTML = html;
}

function renderComputedComparison(offer, validation) {
  const container = document.getElementById('import-computed-comparison');
  const computed = validation.computed;
  
  let html = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;">';
  
  for (const c of computed) {
    const matchClass = c.match ? 'val-match' : 'val-mismatch';
    const matchIcon = c.match ? '✅' : '❌';
    
    html += '<div class="card" style="text-align:center;padding:20px;">' +
      '<div style="font-size:12px;color:var(--text-dim);margin-bottom:8px">' + escapeHtml(c.label) + '</div>' +
      '<div style="font-size:11px;color:var(--text-dim);margin-bottom:4px;">Excel 值</div>' +
      '<div style="font-size:18px;font-weight:700;margin-bottom:8px">' + formatNumber(c.excelValue) + '</div>' +
      '<div style="font-size:16px;margin-bottom:8px">' + matchIcon + '</div>' +
      '<div style="font-size:11px;color:var(--text-dim);margin-bottom:4px;">系统计算</div>' +
      '<div class="' + matchClass + '" style="font-size:18px;font-weight:700">' + formatNumber(c.computedValue) + '</div>' +
      (c.diff !== null && !c.match ? '<div style="font-size:11px;margin-top:6px;color:var(--danger)">差异 ' + (c.diff > 0 ? '+' : '') + formatNumber(c.diff) + '</div>' : '') +
      '</div>';
  }
  
  html += '</div>';
  
  // 分摊年限配置
  html += '<div style="margin-top:16px;padding:12px 16px;background:var(--bg-muted);border-radius:8px;display:flex;align-items:center;gap:12px;">' +
    '<span style="font-size:13px;font-weight:600;">⚙️ 分摊年限配置</span>' +
    '<span style="font-size:12px;color:var(--text-dim);">签字费/签字股票/安家费分摊 N 年（影响「年收入含一次性」计算）</span>' +
    '<input type="number" class="form-input" style="width:80px;text-align:center;" id="amortization-years" value="' + AppState.amortizationYears + '" min="1" max="10">' +
    '<span style="font-size:12px;color:var(--text-dim);">年</span>' +
    '<button class="btn btn-secondary btn-sm" id="btn-recalc">重新计算</button>' +
    '</div>';
  
  container.innerHTML = html;
  
  // 分摊年限变更事件
  document.getElementById('btn-recalc').addEventListener('click', () => {
    const years = parseInt(document.getElementById('amortization-years').value) || 2;
    AppState.amortizationYears = years;
    lsSet('amortizationYears', years);
    
    // 重新核算
    const validations = [];
    for (const offer of AppState.currentOffers) {
      const v = validateOffer(offer, years);
      validations.push(v);
    }
    AppState.currentValidations = validations;
    renderValidationTable(AppState.currentOffers[0], validations[0], AppState.importResult);
    renderComputedComparison(AppState.currentOffers[0], validations[0]);
    showToast('已按 ' + years + ' 年分摊重新计算');
  });
}

// ========== 规则库页面 ==========
let _editingRule = null; // 当前编辑中的规则对象

function initRulesPage() {
  // Tab 切换：统一规则 / 我的规则 / 配置表
  document.querySelectorAll('#rules-tabs .tab-s-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#rules-tabs .tab-s-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const tab = btn.dataset.rulesTab;
      document.querySelectorAll('#page-rules .config-panel').forEach(p => p.classList.remove('active'));
      const panel = document.getElementById('rules-panel-' + tab);
      if (panel) panel.classList.add('active');
    });
  });

  // 配置表子 Tab
  document.querySelectorAll('#config-sub-tabs .config-sub-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#config-sub-tabs .config-sub-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const tab = btn.dataset.configTab;
      document.querySelectorAll('#rules-panel-config .config-panel').forEach(p => p.classList.remove('active'));
      const panel = document.getElementById('config-' + tab);
      if (panel) panel.classList.add('active');
    });
  });

  // 加载行业通用规则
  document.getElementById('btn-init-default-rules').addEventListener('click', async () => {
    const loaded = await loadIndustryDefaultRules();
    if (loaded) {
      showToast('已加载 6 条行业通用规则');
      await logActivity('rules', '加载了行业通用规则模板');
    } else {
      showToast('已有统一规则，跳过默认加载', 'warning');
    }
    await renderRuleList('unified');
    await refreshDashboard();
  });

  // 导入规则
  document.getElementById('btn-import-unified-rules').addEventListener('click', () => {
    const input = document.getElementById('rules-import-input');
    input.dataset.scope = 'unified';
    input.click();
  });

  document.getElementById('rules-import-input').addEventListener('change', async (e) => {
    if (!e.target.files.length) return;
    const file = e.target.files[0];
    const scope = e.target.dataset.scope || 'unified';
    e.target.value = '';
    try {
      const text = await file.text();
      const rules = JSON.parse(text);
      if (!Array.isArray(rules)) { showToast('JSON 格式无效，需要数组', 'error'); return; }
      await importRules(rules, scope);
      showToast('导入了 ' + rules.length + ' 条规则');
      await logActivity('rules', '导入了 ' + rules.length + ' 条' + (scope === 'unified' ? '统一' : '个人') + '规则');
      await renderRuleList(scope);
      await refreshDashboard();
    } catch (err) {
      showToast('导入失败：' + err.message, 'error');
    }
  });

  // 导出规则
  document.getElementById('btn-export-unified-rules').addEventListener('click', async () => {
    const json = await exportRulesJSON('unified');
    const blob = new Blob([json], { type: 'application/json' });
    downloadBlob(blob, '统一规则_' + new Date().toISOString().slice(0, 10) + '.json');
  });
  document.getElementById('btn-export-personal-rules').addEventListener('click', async () => {
    const json = await exportRulesJSON('personal');
    const blob = new Blob([json], { type: 'application/json' });
    downloadBlob(blob, '个人规则_' + new Date().toISOString().slice(0, 10) + '.json');
  });

  // 添加个人规则
  document.getElementById('btn-add-personal-rule').addEventListener('click', () => {
    openRuleEditor(null, 'personal');
  });

  // 规则编辑模态框
  document.getElementById('rule-edit-close').addEventListener('click', closeRuleEditor);
  document.getElementById('rule-edit-cancel').addEventListener('click', closeRuleEditor);
  document.getElementById('modal-rule-edit').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-rule-edit')) closeRuleEditor();
  });
  document.getElementById('re-add-condition').addEventListener('click', addConditionRow);
  document.getElementById('rule-edit-save').addEventListener('click', saveCurrentRule);

  // 配置表 — 竞企清单
  document.getElementById('btn-add-competitor').addEventListener('click', addCompetitor);
  document.getElementById('competitor-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addCompetitor(); });

  // 配置表 — 稀缺岗位
  document.getElementById('btn-add-rare').addEventListener('click', addRarePosition);
  document.getElementById('rare-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addRarePosition(); });

  // 配置表 — 能力标签
  document.getElementById('btn-add-ability').addEventListener('click', addAbilityTag);
  document.getElementById('ability-name-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addAbilityTag(); });
  document.getElementById('btn-save-ability-cap').addEventListener('click', saveAbilityCap);

  // 配置表 — 因子权重
  document.getElementById('btn-reset-weights').addEventListener('click', resetWeights);
  document.getElementById('btn-save-weights').addEventListener('click', saveWeightsFromUI);

  // 配置表 — 分摊年限
  document.getElementById('btn-save-amort').addEventListener('click', async () => {
    const years = parseInt(document.getElementById('config-amort-years').value) || 2;
    await setConfig(CONFIG_KEYS.AMORTIZATION_YEARS, years);
    AppState.amortizationYears = years;
    lsSet('amortizationYears', years);
    showToast('分摊年限已保存为 ' + years + ' 年');
  });

  // 初始加载
  renderRuleList('unified');
  renderRuleList('personal');
  loadConfigPanels();
}

// ---- 规则列表渲染 ----
async function renderRuleList(scope) {
  const rules = await loadRules(scope);
  const container = document.getElementById(scope + '-rule-list');
  const counter = document.getElementById(scope + '-count');
  
  if (counter) counter.textContent = '共 ' + rules.length + ' 条';
  
  if (rules.length === 0) {
    container.innerHTML = '<div class="empty-state" style="padding:32px"><div class="es-icon">' +
      (scope === 'unified' ? '📐' : '✏️') + '</div><div class="es-title">暂无' +
      (scope === 'unified' ? '统一' : '个人') + '规则</div><div class="es-desc">' +
      (scope === 'unified' ? '点击「加载行业通用规则」快速开始' : '点击「添加规则」创建您的个人规则') + '</div></div>';
    return;
  }
  
  let html = '';
  for (const rule of rules.sort((a, b) => (a.priority || 99) - (b.priority || 99))) {
    const catInfo = RULE_CATEGORIES[rule.category] || { label: '未知', icon: '❓' };
    const enabledClass = rule.enabled ? '' : ' disabled';
    const toggleClass = rule.enabled ? ' on' : '';
    const sevClass = rule.action?.type || 'check';
    const sevLabel = { block: '阻断', warn: '警告', check: '提示' }[sevClass] || sevClass;
    
    html += '<div class="rule-item' + enabledClass + '" data-rule-id="' + rule.id + '">' +
      '<div class="rule-top">' +
      '<span class="rule-cat">' + catInfo.icon + '</span>' +
      '<span class="rule-name">' + escapeHtml(rule.name) + '</span>' +
      '<span class="rule-scope-tag ' + rule.scope + '">' + (rule.scope === 'unified' ? '统一' : '个人') + '</span>' +
      '<span class="rule-severity ' + sevClass + '">' + sevLabel + '</span>' +
      '<div class="rule-toggle' + toggleClass + '" data-toggle-id="' + rule.id + '" title="' + (rule.enabled ? '已启用' : '已禁用') + '"></div>' +
      '</div>' +
      '<div class="rule-bottom">';
    
    // 展示条件摘要
    if (rule.conditions && rule.conditions.length > 0) {
      for (const c of rule.conditions) {
        html += '<span class="condition-tag">' + escapeHtml(c.field) + ' ' + c.operator + ' ' + escapeHtml(JSON.stringify(c.value)) + '</span>';
      }
    }
    
    html += '<span style="color:var(--text-dim);font-size:11px">P' + (rule.priority || 99) + '</span>';
    html += '<div class="rule-actions">' +
      '<button class="edit" data-edit-id="' + rule.id + '">编辑</button>' +
      '<button class="del" data-del-id="' + rule.id + '">删除</button>' +
      '</div></div></div>';
  }
  
  container.innerHTML = html;
  
  // 绑定事件
  container.querySelectorAll('.rule-toggle').forEach(toggle => {
    toggle.addEventListener('click', async () => {
      const ruleId = toggle.dataset.toggleId;
      const rule = await dbGet(STORES.rules, ruleId);
      if (rule) {
        rule.enabled = !rule.enabled;
        await saveRule(rule);
        await renderRuleList(scope);
      }
    });
  });
  
  container.querySelectorAll('.edit').forEach(btn => {
    btn.addEventListener('click', async () => {
      const rule = await dbGet(STORES.rules, btn.dataset.editId);
      if (rule) openRuleEditor(rule, rule.scope);
    });
  });
  
  container.querySelectorAll('.del').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm('确定删除此规则？')) return;
      await deleteRule(btn.dataset.delId);
      showToast('规则已删除');
      await renderRuleList(scope);
      await refreshDashboard();
    });
  });
}

// ---- 规则编辑模态框 ----
function openRuleEditor(rule, scope) {
  _editingRule = rule ? { ...rule } : createRuleTemplate(scope);
  
  document.getElementById('rule-edit-title').textContent = rule ? '编辑规则' : '新建规则';
  document.getElementById('re-name').value = _editingRule.name || '';
  document.getElementById('re-category').value = _editingRule.category || 'custom';
  document.getElementById('re-action-type').value = (_editingRule.action && _editingRule.action.type) || 'check';
  document.getElementById('re-message').value = (_editingRule.action && _editingRule.action.message) || '';
  document.getElementById('re-priority').value = _editingRule.priority || 50;
  
  // 渲染条件
  const condContainer = document.getElementById('re-conditions');
  condContainer.innerHTML = '';
  const conditions = _editingRule.conditions || [{ field: '', operator: 'eq', value: '' }];
  for (const c of conditions) {
    condContainer.appendChild(createConditionRowEl(c));
  }
  
  document.getElementById('modal-rule-edit').classList.add('show');
}

function closeRuleEditor() {
  document.getElementById('modal-rule-edit').classList.remove('show');
  _editingRule = null;
}

function createConditionRowEl(cond) {
  const div = document.createElement('div');
  div.className = 'rule-edit-row';
  
  const fieldInput = document.createElement('input');
  fieldInput.placeholder = '字段路径（如 offer.baseSalary）';
  fieldInput.value = cond.field || '';
  fieldInput.style.flex = '2';
  
  const opSelect = document.createElement('select');
  const ops = ['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'in', 'notIn', 'contains', 'between', 'regex'];
  for (const op of ops) {
    const opt = document.createElement('option');
    opt.value = op; opt.textContent = op;
    if (op === (cond.operator || 'eq')) opt.selected = true;
    opSelect.appendChild(opt);
  }
  
  const valInput = document.createElement('input');
  valInput.placeholder = '值';
  valInput.value = typeof cond.value === 'object' ? JSON.stringify(cond.value) : (cond.value !== undefined ? String(cond.value) : '');
  
  const delBtn = document.createElement('button');
  delBtn.className = 're-del';
  delBtn.textContent = '✕';
  delBtn.addEventListener('click', () => div.remove());
  
  div.appendChild(fieldInput);
  div.appendChild(opSelect);
  div.appendChild(valInput);
  div.appendChild(delBtn);
  
  div._getData = () => {
    let v = valInput.value;
    try { v = JSON.parse(v); } catch (e) {
      const n = Number(v);
      if (!isNaN(n) && v !== '') v = n;
    }
    return { field: fieldInput.value, operator: opSelect.value, value: v };
  };
  
  return div;
}

function addConditionRow() {
  document.getElementById('re-conditions').appendChild(createConditionRowEl({ field: '', operator: 'eq', value: '' }));
}

async function saveCurrentRule() {
  if (!_editingRule) return;
  
  const name = document.getElementById('re-name').value.trim();
  if (!name) { showToast('请填写规则名称', 'warning'); return; }
  
  _editingRule.name = name;
  _editingRule.category = document.getElementById('re-category').value;
  _editingRule.action = {
    type: document.getElementById('re-action-type').value,
    message: document.getElementById('re-message').value.trim() || name,
  };
  _editingRule.priority = parseInt(document.getElementById('re-priority').value) || 50;
  
  // 收集条件
  const condRows = document.querySelectorAll('#re-conditions .rule-edit-row');
  _editingRule.conditions = [];
  condRows.forEach(row => {
    if (row._getData) _editingRule.conditions.push(row._getData());
  });
  
  await saveRule(_editingRule);
  showToast('规则已保存');
  closeRuleEditor();
  await renderRuleList(_editingRule.scope);
  await refreshDashboard();
}

// ---- 配置表面板加载 ----
async function loadConfigPanels() {
  await renderWeightsPanel();
  await renderCompetitorTags();
  await renderRareTags();
  await renderAbilityTags();
  
  // 分摊年限
  const amort = await getConfig(CONFIG_KEYS.AMORTIZATION_YEARS, 2);
  document.getElementById('config-amort-years').value = amort;
}

// -- 因子权重面板 --
async function renderWeightsPanel() {
  const weights = await loadFactorWeights();
  const grid = document.getElementById('weights-grid');
  
  grid.innerHTML = '' +
    weightCard('基线水位', 'baseline', [
      wRow('启用', 'baseline-enabled', 'checkbox', weights.baseline.enabled),
    ]) +
    weightCard('候选人增幅', 'raise', [
      wRow('默认上限', 'raise-defaultCap', 'number', (weights.raise.defaultCap * 100).toFixed(0), '%'),
      wRow('历史中位数', 'raise-learnedMedian', 'text', weights.raise.learnedMedian !== null ? (weights.raise.learnedMedian * 100).toFixed(1) + '%' : '待学习', '', true),
    ]) +
    weightCard('外部竞争', 'competition', [
      wRow('一线竞对溢价', 'comp-tier1', 'number', (weights.competition.tierPremium['一线竞对'] * 100).toFixed(0), '%'),
      wRow('二线竞对溢价', 'comp-tier2', 'number', (weights.competition.tierPremium['二线竞对'] * 100).toFixed(0), '%'),
    ]) +
    weightCard('内部平衡', 'internal', [
      wRow('天花板系数', 'int-ceiling', 'number', weights.internal.ceilingRatio, '×'),
      wRow('地板系数', 'int-floor', 'number', weights.internal.floorRatio, '×'),
    ]) +
    weightCard('紧迫度', 'urgency', [
      wRow('业务紧急上浮', 'urg-biz-hi', 'number', (weights.urgency.business['紧急'] * 100).toFixed(0), '%'),
      wRow('入职1月内上浮', 'urg-onb-hi', 'number', (weights.urgency.onboarding['1个月内'] * 100).toFixed(0), '%'),
    ]) +
    weightCard('能力标签', 'ability', [
      wRow('累计封顶', 'ability-cap-w', 'number', (weights.ability.capPercent * 100).toFixed(0), '%'),
    ]) +
    weightCard('岗位稀缺度', 'rarity', [
      wRow('整档上浮', 'rarity-uplift', 'number', (weights.rarity.uplift * 100).toFixed(0), '%'),
    ]) +
    weightCard('薪酬结构偏好', 'structure', [
      wRow('Base 占比', 'struct-base', 'number', (weights.structure.basePct * 100).toFixed(0), '%'),
      wRow('股票占比', 'struct-stock', 'number', (weights.structure.stockPct * 100).toFixed(0), '%'),
      wRow('签字费占比', 'struct-sign', 'number', (weights.structure.signBonusPct * 100).toFixed(0), '%'),
    ]);
}

function weightCard(title, id, rows) {
  return '<div class="weight-card"><h4>' + title + '</h4>' + rows.join('') + '</div>';
}

function wRow(label, inputId, type, value, suffix, readonly) {
  suffix = suffix || '';
  if (type === 'checkbox') {
    return '<div class="w-row"><span class="w-label">' + label + '</span><input type="checkbox" id="w-' + inputId + '"' + (value ? ' checked' : '') + '></div>';
  }
  if (readonly) {
    return '<div class="w-row"><span class="w-label">' + label + '</span><span class="w-val">' + value + '</span></div>';
  }
  return '<div class="w-row"><span class="w-label">' + label + '</span><input type="number" id="w-' + inputId + '" value="' + value + '" step="1">' + (suffix ? '<span style="font-size:11px;color:var(--text-dim)">' + suffix + '</span>' : '') + '</div>';
}

async function resetWeights() {
  if (!confirm('恢复为系统默认因子权重？')) return;
  await saveFactorWeights(getDefaultFactorWeights());
  await renderWeightsPanel();
  showToast('因子权重已恢复默认');
}

async function saveWeightsFromUI() {
  const g = (id) => { const el = document.getElementById('w-' + id); return el ? (el.type === 'checkbox' ? el.checked : parseFloat(el.value)) : 0; };
  
  const weights = {
    baseline: { enabled: g('baseline-enabled') },
    raise: { defaultCap: g('raise-defaultCap') / 100, learnedMedian: null },
    competition: {
      tierPremium: { '一线竞对': g('comp-tier1') / 100, '二线竞对': g('comp-tier2') / 100, '非竞对': 0 },
    },
    internal: { ceilingRatio: g('int-ceiling'), floorRatio: g('int-floor') },
    urgency: {
      business: { '紧急': g('urg-biz-hi') / 100, '正常': 0, '不紧急': -0.05 },
      onboarding: { '1个月内': g('urg-onb-hi') / 100, '1-3个月': 0, '3个月以上': -0.03 },
    },
    ability: { perTag: {}, capPercent: g('ability-cap-w') / 100 },
    rarity: { positions: [], uplift: g('rarity-uplift') / 100 },
    structure: { basePct: g('struct-base') / 100, stockPct: g('struct-stock') / 100, signBonusPct: g('struct-sign') / 100 },
    userOverrides: {},
  };
  
  // 保留已有 perTag 和 positions
  const existing = await loadFactorWeights();
  weights.ability.perTag = existing.ability.perTag || {};
  weights.rarity.positions = existing.rarity.positions || [];
  
  await saveFactorWeights(weights);
  showToast('因子权重已保存');
}

// -- 竞企清单 --
async function renderCompetitorTags() {
  const list = await loadCompetitorList();
  const container = document.getElementById('competitor-tags');
  if (list.length === 0) {
    container.innerHTML = '<span style="font-size:12px;color:var(--text-dim)">暂无竞企，请添加</span>';
    return;
  }
  container.innerHTML = list.map((name, i) =>
    '<span class="tag-item">' + escapeHtml(name) + '<span class="tag-del" data-idx="' + i + '">×</span></span>'
  ).join('');
  container.querySelectorAll('.tag-del').forEach(del => {
    del.addEventListener('click', async () => {
      const idx = parseInt(del.dataset.idx);
      const l = await loadCompetitorList();
      l.splice(idx, 1);
      await saveCompetitorList(l);
      await renderCompetitorTags();
    });
  });
}

async function addCompetitor() {
  const input = document.getElementById('competitor-input');
  const val = input.value.trim();
  if (!val) return;
  const list = await loadCompetitorList();
  if (list.includes(val)) { showToast('已存在', 'warning'); return; }
  list.push(val);
  await saveCompetitorList(list);
  input.value = '';
  await renderCompetitorTags();
  showToast('已添加竞企：' + val);
}

// -- 稀缺岗位 --
async function renderRareTags() {
  const list = await loadRarePositions();
  const container = document.getElementById('rare-tags');
  if (list.length === 0) {
    container.innerHTML = '<span style="font-size:12px;color:var(--text-dim)">暂无稀缺岗位，请添加</span>';
    return;
  }
  container.innerHTML = list.map((name, i) =>
    '<span class="tag-item">' + escapeHtml(name) + '<span class="tag-del" data-idx="' + i + '">×</span></span>'
  ).join('');
  container.querySelectorAll('.tag-del').forEach(del => {
    del.addEventListener('click', async () => {
      const idx = parseInt(del.dataset.idx);
      const l = await loadRarePositions();
      l.splice(idx, 1);
      await saveRarePositions(l);
      await renderRareTags();
    });
  });
}

async function addRarePosition() {
  const input = document.getElementById('rare-input');
  const val = input.value.trim();
  if (!val) return;
  const list = await loadRarePositions();
  if (list.includes(val)) { showToast('已存在', 'warning'); return; }
  list.push(val);
  await saveRarePositions(list);
  input.value = '';
  await renderRareTags();
  showToast('已添加稀缺岗位：' + val);
}

// -- 能力标签库 --
async function renderAbilityTags() {
  const weights = await loadFactorWeights();
  const tags = weights.ability.perTag || {};
  const container = document.getElementById('ability-tag-list');
  const entries = Object.entries(tags);
  
  if (entries.length === 0) {
    container.innerHTML = '<span style="font-size:12px;color:var(--text-dim)">暂无能力标签，请添加</span>';
  } else {
    container.innerHTML = '<div class="tag-list">' +
      entries.map(([name, pct]) =>
        '<span class="tag-item">' + escapeHtml(name) + ' <strong style="color:var(--primary)">+' + (pct * 100).toFixed(0) + '%</strong>' +
        '<span class="tag-del" data-tag-name="' + escapeHtml(name) + '">×</span></span>'
      ).join('') + '</div>';
    
    container.querySelectorAll('.tag-del').forEach(del => {
      del.addEventListener('click', async () => {
        const w = await loadFactorWeights();
        delete w.ability.perTag[del.dataset.tagName];
        await saveFactorWeights(w);
        await renderAbilityTags();
      });
    });
  }
  
  // 封顶值
  document.getElementById('ability-cap-input').value = (weights.ability.capPercent * 100).toFixed(0);
}

async function addAbilityTag() {
  const nameInput = document.getElementById('ability-name-input');
  const pctInput = document.getElementById('ability-pct-input');
  const name = nameInput.value.trim();
  const pct = parseFloat(pctInput.value);
  if (!name) { showToast('请输入标签名称', 'warning'); return; }
  if (isNaN(pct) || pct <= 0) { showToast('请输入有效的加成百分比', 'warning'); return; }
  
  const weights = await loadFactorWeights();
  weights.ability.perTag[name] = pct / 100;
  await saveFactorWeights(weights);
  nameInput.value = '';
  pctInput.value = '';
  await renderAbilityTags();
  showToast('已添加能力标签：' + name + ' +' + pct + '%');
}

async function saveAbilityCap() {
  const cap = parseFloat(document.getElementById('ability-cap-input').value);
  if (isNaN(cap)) { showToast('请输入有效的封顶百分比', 'warning'); return; }
  const weights = await loadFactorWeights();
  weights.ability.capPercent = cap / 100;
  await saveFactorWeights(weights);
  showToast('累计封顶已保存为 ' + cap + '%');
}

// ========== 样本规划向导 ==========
let _wizardDims = { countries: [], levels: [], channels: [], jobFamilies: [] };

function initSampleWizard() {
  // 从 Dashboard 打开
  document.getElementById('qa-sample-wizard').addEventListener('click', openSampleWizard);
  
  // 关闭
  document.getElementById('wizard-close').addEventListener('click', closeSampleWizard);
  document.getElementById('wizard-cancel').addEventListener('click', closeSampleWizard);
  document.getElementById('modal-wizard').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-wizard')) closeSampleWizard();
  });
  
  // 添加维度值
  const addDim = (inputId, dimKey) => {
    const input = document.getElementById(inputId);
    const val = input.value.trim();
    if (!val) return;
    if (!_wizardDims[dimKey].includes(val)) _wizardDims[dimKey].push(val);
    input.value = '';
    renderWizardChips();
    calcWizardResult();
  };
  
  document.getElementById('wizard-add-country').addEventListener('click', () => addDim('wizard-country-input', 'countries'));
  document.getElementById('wizard-country-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addDim('wizard-country-input', 'countries'); });
  document.getElementById('wizard-add-level').addEventListener('click', () => addDim('wizard-level-input', 'levels'));
  document.getElementById('wizard-level-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addDim('wizard-level-input', 'levels'); });
  document.getElementById('wizard-add-channel').addEventListener('click', () => addDim('wizard-channel-input', 'channels'));
  document.getElementById('wizard-channel-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addDim('wizard-channel-input', 'channels'); });
  document.getElementById('wizard-add-family').addEventListener('click', () => addDim('wizard-family-input', 'jobFamilies'));
  document.getElementById('wizard-family-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addDim('wizard-family-input', 'jobFamilies'); });
}

async function openSampleWizard() {
  // 从历史 Offer 提取已有维度
  const historyOffers = await dbGetAll(STORES.historyOffers);
  if (historyOffers.length > 0) {
    const dims = extractDimensions(historyOffers);
    _wizardDims = dims;
  } else {
    _wizardDims = { countries: [], levels: [], channels: [], jobFamilies: [] };
  }
  
  renderWizardChips();
  await calcWizardResult();
  document.getElementById('modal-wizard').classList.add('show');
}

function closeSampleWizard() {
  document.getElementById('modal-wizard').classList.remove('show');
}

function renderWizardChips() {
  const renderChipGroup = (containerId, dimKey) => {
    const container = document.getElementById(containerId);
    const items = _wizardDims[dimKey] || [];
    if (items.length === 0) {
      container.innerHTML = '<span style="font-size:12px;color:var(--text-dim)">暂无，请手动添加</span>';
      return;
    }
    container.innerHTML = items.map(v =>
      '<span class="wizard-chip selected" data-dim="' + dimKey + '" data-val="' + escapeHtml(v) + '">' + escapeHtml(v) + ' ×</span>'
    ).join('');
    container.querySelectorAll('.wizard-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        const d = chip.dataset.dim;
        const val = chip.dataset.val;
        _wizardDims[d] = _wizardDims[d].filter(x => x !== val);
        renderWizardChips();
        calcWizardResult();
      });
    });
  };
  
  renderChipGroup('wizard-countries', 'countries');
  renderChipGroup('wizard-levels', 'levels');
  renderChipGroup('wizard-channels', 'channels');
  renderChipGroup('wizard-families', 'jobFamilies');
}

async function calcWizardResult() {
  const historyOffers = await dbGetAll(STORES.historyOffers);
  const currentSamples = historyOffers.length;
  const plan = generateSamplePlan(_wizardDims, currentSamples);
  
  document.getElementById('wizard-groups').textContent = plan.groups;
  document.getElementById('wizard-current').textContent = currentSamples;
  
  const container = document.getElementById('wizard-tiers');
  container.innerHTML = plan.tiers.map(t =>
    '<div class="wizard-tier">' +
    '<span class="tier-label" style="color:' + t.color + '">' + t.label + '</span>' +
    '<div class="tier-bar"><div class="tier-bar-fill" style="width:' + t.progress + '%;background:' + t.color + '"></div></div>' +
    '<span class="tier-text">' + currentSamples + '/' + t.needed + '（' + t.progress + '%）' + (t.gap > 0 ? ' 差 ' + t.gap : ' ✅') + '</span>' +
    '</div>'
  ).join('');
}

// ========== 画像页面 ==========
function initProfilePage() {
  // 导入历史 Offer
  document.getElementById('btn-import-history-offers').addEventListener('click', () => {
    document.getElementById('history-import-input').click();
  });
  document.getElementById('history-import-input').addEventListener('change', async (e) => {
    if (!e.target.files.length) return;
    const file = e.target.files[0];
    e.target.value = '';
    showLoading('正在导入历史 Offer...');
    try {
      const result = await importHistoryOffers(file);
      hideLoading();
      if (result.count > 0) {
        showToast('成功导入 ' + result.count + ' 条历史 Offer');
        await logActivity('import', '导入了 ' + result.count + ' 条历史 Offer');
      }
      if (result.errors.length > 0) {
        showToast(result.errors.slice(0, 3).join('；'), 'warning');
      }
      await refreshProfilePage();
      await refreshDashboard();
    } catch (err) {
      hideLoading();
      showToast('导入失败：' + err.message, 'error');
    }
  });

  // 学习画像
  document.getElementById('btn-learn-profile').addEventListener('click', async () => {
    const count = await dbCount(STORES.historyOffers);
    if (count === 0) {
      showToast('请先导入历史 Offer', 'warning');
      return;
    }
    showLoading('正在学习画像（4 维切分 + 统计分析）...');
    try {
      const insights = await learnProfile();
      await saveProfileInsights(insights);

      // 拟合因子权重
      const fitted = await fitFactorWeights();
      if (Object.keys(fitted).length > 0) {
        const weights = await loadFactorWeights();
        const merged = deepMerge(weights, fitted);
        await saveFactorWeights(merged);
      }

      hideLoading();
      showToast('画像学习完成，共生成 ' + insights.length + ' 条洞察');
      await logActivity('profile', '学习画像完成，生成 ' + insights.length + ' 条洞察');
      await refreshProfilePage();
    } catch (err) {
      hideLoading();
      showToast('画像学习失败：' + err.message, 'error');
    }
  });

  // 从画像页打开样本向导
  document.getElementById('btn-open-wizard-from-profile').addEventListener('click', openSampleWizard);

  // 初始加载
  refreshProfilePage();
}

async function refreshProfilePage() {
  const historyCount = await dbCount(STORES.historyOffers);
  const profileCount = await dbCount(STORES.profiles);
  document.getElementById('profile-summary').textContent =
    '历史样本：' + historyCount + ' 条 · 画像条目：' + profileCount + ' 条';

  await renderCoverageHeatmap();
  await renderProfileGroups();
}

// ---- 覆盖度热力图 ----
async function renderCoverageHeatmap() {
  const container = document.getElementById('coverage-heatmap');
  const offers = await dbGetAll(STORES.historyOffers);

  if (offers.length === 0) {
    container.innerHTML = '<div class="empty-state" style="padding:24px"><div class="es-icon">📊</div><div class="es-title">导入历史 Offer 后查看覆盖度</div></div>';
    return;
  }

  const dims = extractDimensions(offers);
  const matrix = calcCoverageMatrix(offers, dims);
  const byCountry = getCoverageHeatmapData(matrix);

  let html = '<div class="heatmap-wrap">';
  for (const [country, cells] of Object.entries(byCountry)) {
    html += '<div class="heatmap-country"><h4>🌍 ' + escapeHtml(country) + '</h4>';

    // 获取此国家下的唯一职级和通道
    const levels = [...new Set(cells.map(c => c.level))].sort();
    const channels = [...new Set(cells.map(c => c.channel))].sort();

    html += '<table class="heatmap-table"><thead><tr><th></th>';
    for (const lv of levels) html += '<th>' + escapeHtml(lv) + '</th>';
    html += '</tr></thead><tbody>';

    for (const ch of channels) {
      html += '<tr><th style="text-align:left">' + escapeHtml(ch) + '</th>';
      for (const lv of levels) {
        const cell = cells.find(c => c.level === lv && c.channel === ch);
        const n = cell ? cell.count : 0;
        const cls = n === 0 ? 'hm-0' : n < 5 ? 'hm-low' : n < 15 ? 'hm-mid' : n < 30 ? 'hm-good' : 'hm-ideal';
        html += '<td class="' + cls + '">' + n + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

// ---- 画像分组列表 ----
async function renderProfileGroups() {
  const container = document.getElementById('profile-groups-container');
  const overview = await getProfileOverview();

  if (overview.length === 0) {
    container.innerHTML = '<div class="empty-state" id="profile-empty"><div class="es-icon">🧬</div><div class="es-title">暂无画像数据</div><div class="es-desc">导入历史已审批 Offer 后点击「学习画像」，系统将自动生成审批习惯画像</div></div>';
    return;
  }

  const all = await loadAllProfiles();

  let html = '';
  for (const group of overview) {
    const sk = group.sliceKey;
    const confPct = Math.round(group.avgConfidence * 100);
    const confColor = confPct >= 80 ? 'var(--success)' : confPct >= 50 ? 'var(--warning)' : 'var(--danger)';
    const groupId = 'pg_' + btoa(JSON.stringify(sk)).replace(/[^a-zA-Z0-9]/g, '').slice(0, 12);
    const groupInsights = all.filter(p =>
      p.sliceKey.country === sk.country && p.sliceKey.level === sk.level &&
      p.sliceKey.channel === sk.channel && p.sliceKey.jobFamily === sk.jobFamily
    );

    html += '<div class="profile-group" id="' + groupId + '">';
    html += '<div class="profile-group-header" data-group="' + groupId + '">';
    html += '<span class="pg-icon">📁</span>';
    html += '<span class="pg-title">' + escapeHtml(sk.country) + ' · ' + escapeHtml(sk.level) + ' · ' + escapeHtml(sk.channel) + ' · ' + escapeHtml(sk.jobFamily) + '</span>';
    html += '<span class="pg-meta">';
    html += '<span>' + group.insightCount + ' 条洞察</span>';
    html += '<span>' + group.sampleIds.length + ' 样本</span>';
    html += '<span class="pg-confidence"><span class="conf-bar"><span class="conf-fill" style="width:' + confPct + '%;background:' + confColor + '"></span></span> ' + confPct + '%</span>';
    html += '</span>';
    html += '<span class="pg-expand">▼</span>';
    html += '</div>';

    html += '<div class="profile-group-body">';
    for (const insight of groupInsights) {
      html += renderInsightItem(insight);
    }
    html += '</div></div>';
  }

  container.innerHTML = html;

  // 绑定折叠事件
  container.querySelectorAll('.profile-group-header').forEach(header => {
    header.addEventListener('click', () => {
      const group = document.getElementById(header.dataset.group);
      if (group) group.classList.toggle('open');
    });
  });

  // 绑定画像操作按钮
  bindInsightActions(container);
}

function renderInsightItem(insight) {
  const typeClass = insight.type === 'numeric' ? 'numeric' : 'conditional';
  const typeIcon = insight.type === 'numeric' ? '📊' : '🔀';
  const statusClass = insight.userStatus || 'suggested';
  const statusLabel = { suggested: '待确认', confirmed: '已确认', edited: '已编辑', rejected: '已忽略' }[statusClass] || statusClass;

  let patternHtml = '';
  if (insight.type === 'numeric' && insight.pattern) {
    patternHtml = 'P25: ' + formatNumber(insight.pattern.p25) +
      ' · <strong>P50: ' + formatNumber(insight.pattern.median) + '</strong>' +
      ' · P75: ' + formatNumber(insight.pattern.p75) +
      '<br>范围: ' + formatNumber(insight.pattern.min) + ' ~ ' + formatNumber(insight.pattern.max) +
      ' · 均值: ' + formatNumber(insight.pattern.mean);
  } else if (insight.type === 'conditional' && insight.pattern) {
    if (insight.pattern.effect !== undefined) {
      patternHtml = '条件: ' + escapeHtml(insight.pattern.condition || '-') +
        ' → 效应: <strong>' + (insight.pattern.effect >= 0 ? '+' : '') + (insight.pattern.effect * 100).toFixed(1) + '%</strong>';
    } else if (insight.pattern.basePctMedian !== undefined) {
      patternHtml = 'Base 占比中位数: ' + (insight.pattern.basePctMedian * 100).toFixed(1) + '%';
      if (insight.pattern.equityPctMedian !== null) {
        patternHtml += ' · 股票占比中位数: ' + (insight.pattern.equityPctMedian * 100).toFixed(1) + '%';
      }
    }
  }

  return '<div class="profile-insight" data-insight-id="' + insight.id + '">' +
    '<span class="pi-type ' + typeClass + '">' + typeIcon + '</span>' +
    '<div class="pi-info">' +
    '<div class="pi-label">' + escapeHtml(insight.dimensionLabel || insight.dimension) + '</div>' +
    '<div class="pi-pattern">' + patternHtml + '</div>' +
    '<div class="pi-meta">' +
    '<span>样本: ' + insight.sampleCount + '</span>' +
    '<span>置信度: ' + Math.round(insight.confidence * 100) + '%</span>' +
    '<span class="pi-status ' + statusClass + '">' + statusLabel + '</span>' +
    '</div></div>' +
    '<div class="pi-actions">' +
    '<button class="confirm" data-action="confirm" data-id="' + insight.id + '" title="确认">✓</button>' +
    '<button data-action="reject" data-id="' + insight.id + '" title="忽略">✗</button>' +
    (insight.type === 'conditional' ? '<button data-action="toRule" data-id="' + insight.id + '" title="转为规则">→规则</button>' : '') +
    '<button class="reject" data-action="delete" data-id="' + insight.id + '" title="删除">🗑</button>' +
    '</div></div>';
}

function bindInsightActions(container) {
  container.querySelectorAll('.pi-actions button').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const action = btn.dataset.action;
      const id = btn.dataset.id;

      if (action === 'confirm') {
        await updateProfileStatus(id, 'confirmed');
        showToast('已确认');
      } else if (action === 'reject') {
        await updateProfileStatus(id, 'rejected');
        showToast('已忽略');
      } else if (action === 'toRule') {
        const profile = await dbGet(STORES.profiles, id);
        if (profile) {
          await convertToPersonalRule(profile);
          showToast('已转为个人规则');
          await renderRuleList('personal');
        }
      } else if (action === 'delete') {
        if (!confirm('确定删除此画像条目？')) return;
        await deleteProfile(id);
        showToast('已删除');
      }
      await renderProfileGroups();
    });
  });
}

// ========== Offer 体检页面 ==========
let _lastAdvisorResult = null; // 保存最近一次推荐结果供审批说明页使用

function initCheckupPage() {
  document.getElementById('btn-run-checkup').addEventListener('click', runCheckup);
  document.getElementById('btn-to-approval').addEventListener('click', () => navigateTo('approval'));
}

function updateCheckupEmpty() {
  const hasOffers = AppState.currentOffers && AppState.currentOffers.length > 0;
  document.getElementById('checkup-empty').classList.toggle('hidden', hasOffers);
  document.getElementById('checkup-context-card').classList.toggle('hidden', !hasOffers);
  if (hasOffers) {
    const o = AppState.currentOffers[0];
    document.getElementById('checkup-offer-info').textContent =
      '当前 Offer：' + (o.country || '-') + ' · ' + (o.level || '-') + ' · ' + (o.channel || '-') + ' · base ' + formatNumber(o.baseSalary);
  }
}

async function runCheckup() {
  if (!AppState.currentOffers || AppState.currentOffers.length === 0) {
    showToast('请先导入薪酬数据', 'warning');
    return;
  }

  const offer = AppState.currentOffers[0];
  const context = {
    sourceCompany: document.getElementById('ctx-source').value.trim(),
    currentSalary: parseFloat(document.getElementById('ctx-salary').value) || 0,
    counterOffer: parseFloat(document.getElementById('ctx-counter').value) || 0,
    bizUrgency: document.getElementById('ctx-biz-urg').value,
    onboardUrgency: document.getElementById('ctx-onb-urg').value,
    abilityTags: document.getElementById('ctx-tags').value.split(/[,，;；]/).map(s => s.trim()).filter(Boolean),
  };

  showLoading('正在执行 Offer 体检 + 8 因子推荐...');

  try {
    // 1. 规则体检
    const allRules = await loadRules('all');
    const ruleContext = { offer, context, 'context.raisePercent': context.currentSalary > 0 ? (offer.cashIncome - context.currentSalary) / context.currentSalary : 0, 'context.baseRatio': offer.cashIncome > 0 ? (offer.baseSalary * 12) / offer.cashIncome : 0 };
    // 展平 context 到顶层供规则引擎使用
    for (const [k, v] of Object.entries(offer)) ruleContext['offer.' + k] = v;
    ruleContext['context.raisePercent'] = ruleContext['context.raisePercent'];
    ruleContext['context.baseRatio'] = ruleContext['context.baseRatio'];
    const ruleResults = executeRules(allRules, ruleContext);
    const categorized = categorizeResults(ruleResults.results);

    // 2. 推荐引擎
    const advisorResult = await runAdvisor(offer, context);
    _lastAdvisorResult = advisorResult;

    // 保存到 AppState
    AppState.lastCheckupResult = { ruleResults, categorized, advisorResult, offer, context };

    hideLoading();

    // 渲染
    document.getElementById('checkup-results').classList.remove('hidden');
    renderCheckupReport(categorized, ruleResults.summary);
    renderRecommendations(advisorResult.recommendation);
    renderFactorDetails(advisorResult.factors);
    renderRisks(advisorResult.risks);

    await logActivity('checkup', '完成 Offer 体检：' + (offer.country || '') + ' ' + (offer.level || '') + ' base ' + formatNumber(offer.baseSalary));
    await refreshDashboard();

  } catch (err) {
    hideLoading();
    showToast('体检失败：' + err.message, 'error');
  }
}

function renderCheckupReport(categorized, summary) {
  const container = document.getElementById('checkup-report');
  const sections = [
    { key: 'violation', label: '🚫 违规', items: categorized.violation, color: 'var(--danger)', bg: 'var(--danger-dim)' },
    { key: 'risk', label: '⚠️ 风险', items: categorized.risk, color: 'var(--warning)', bg: 'var(--warning-dim)' },
    { key: 'compliance', label: '✅ 合规', items: categorized.compliance, color: 'var(--success)', bg: 'var(--success-dim)' },
  ];

  let html = '<div style="display:flex;gap:12px;margin-bottom:16px">';
  html += '<span class="badge badge-success">通过 ' + summary.pass + '</span>';
  html += '<span class="badge badge-warning">风险 ' + (summary.check + summary.warn) + '</span>';
  html += '<span class="badge badge-danger">违规 ' + summary.block + '</span>';
  html += '</div>';

  for (const sec of sections) {
    if (sec.items.length === 0) continue;
    html += '<div class="health-section"><div class="health-header" onclick="this.nextElementSibling.classList.toggle(\'show\')">' +
      '<span class="h-icon">' + sec.label.split(' ')[0] + '</span>' +
      '<span class="h-title">' + sec.label + '</span>' +
      '<span class="h-count">' + sec.items.length + ' 项</span>' +
      '</div><div class="health-detail show">';
    for (const item of sec.items) {
      const catInfo = RULE_CATEGORIES[item.category] || { icon: '❓', label: '' };
      html += '<div style="padding:4px 0">' + catInfo.icon + ' <strong>' + escapeHtml(item.ruleName) + '</strong>：' + escapeHtml(item.message) +
        ' <span style="font-size:10px;color:var(--text-dim)">[' + (item.source === 'system' ? '统一' : item.source === 'profile' ? '画像' : '个人') + ']</span></div>';
    }
    html += '</div></div>';
  }

  container.innerHTML = html;
}

function renderRecommendations(rec) {
  const container = document.getElementById('checkup-recommendations');
  const labels = { conservative: '保守', standard: '标准', aggressive: '进取' };
  const icons = { conservative: '🛡️', standard: '⭐', aggressive: '🚀' };

  let html = '';
  for (const tier of ['conservative', 'standard', 'aggressive']) {
    const pkg = rec[tier];
    const isRec = tier === 'standard';
    html += '<div class="rec-card' + (isRec ? ' recommended' : '') + '">';
    if (isRec) html += '<span class="rec-badge">★ 推荐</span>';
    html += '<div class="rec-label">' + icons[tier] + ' ' + labels[tier] + '</div>';
    html += '<div class="rec-total">' + formatNumber(pkg.totalAnnual) + '</div>';
    html += '<div class="rec-detail">' +
      '<div><span class="label">基本月薪</span>' + formatNumber(pkg.baseSalary) + '</div>' +
      '<div><span class="label">绩效月均</span>' + formatNumber(pkg.perfMonthly) + '</div>' +
      '<div><span class="label">年终奖</span>' + formatNumber(pkg.annualBonus) + '</div>' +
      '<div><span class="label">现金收入</span><strong>' + formatNumber(pkg.cashIncome) + '</strong></div>' +
      '<div><span class="label">股票年化</span>' + formatNumber(pkg.equityAnnual) + '</div>' +
      '<div><span class="label">签字股票</span>' + formatNumber(pkg.signStock) + '</div>' +
      '<div><span class="label">签字费</span>' + formatNumber(pkg.signBonus) + '</div>' +
      '<div><span class="label">安家费</span>' + formatNumber(pkg.relocation) + '</div>' +
      '</div></div>';
  }
  container.innerHTML = html;
}

function renderFactorDetails(factors) {
  const container = document.getElementById('checkup-factors');
  if (!factors || factors.length === 0) {
    container.innerHTML = '<div style="font-size:12px;color:var(--text-dim)">无因子分析数据</div>';
    return;
  }
  const icons = { baseline: '📊', raise: '📈', competition: '🏢', internal: '⚖️', urgency: '⏰', ability: '🏆', rarity: '💎', structure: '🏗️' };
  let html = '<div style="display:flex;flex-direction:column;gap:8px">';
  for (const f of factors) {
    const icon = icons[f.id] || '🔹';
    const hasDelta = f.delta && (f.delta.standard !== 0);
    html += '<div style="display:flex;gap:10px;padding:8px 12px;border-radius:8px;background:var(--bg-muted);align-items:flex-start">' +
      '<span style="font-size:16px">' + icon + '</span>' +
      '<div style="flex:1">' +
      '<div style="font-size:13px;font-weight:600">' + escapeHtml(f.name) + (hasDelta ? ' <span style="color:var(--primary);font-family:monospace">' + (f.delta.standard >= 0 ? '+' : '') + formatNumber(f.delta.standard) + '</span>' : '') + '</div>' +
      '<div style="font-size:12px;color:var(--text-secondary)">' + escapeHtml(f.reason) + '</div>' +
      (f.evidence && f.evidence.note ? '<div style="font-size:11px;color:var(--text-dim);margin-top:2px">' + escapeHtml(f.evidence.note) + (f.evidence.sampleIds && f.evidence.sampleIds.length > 0 ? ' · ' + f.evidence.sampleIds.length + ' 样本' : '') + '</div>' : '') +
      (f.hardConstraint ? '<div style="font-size:11px;color:var(--danger);margin-top:2px">⚠️ ' + escapeHtml(f.hardConstraint.message) + '</div>' : '') +
      '</div></div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

function renderRisks(risks) {
  const container = document.getElementById('checkup-risks');
  if (!risks || risks.length === 0) {
    container.innerHTML = '';
    return;
  }
  let html = '<div class="card"><div class="card-header"><h3>⚠️ 风险提示</h3></div><div class="card-body">';
  for (const r of risks) {
    const badge = r.level === 'danger' ? 'badge-danger' : 'badge-warning';
    html += '<div style="padding:4px 0"><span class="badge ' + badge + '">' + (r.level === 'danger' ? '🚫' : '⚠️') + ' ' + escapeHtml(r.message) + '</span></div>';
  }
  html += '</div></div>';
  container.innerHTML = html;
}

// ========== 审批说明页面 ==========
let _approvalTier = 'standard';
let _approvalStatement = null; // { section1, section2, section3, fullText, markdown }
let _aiEnabled = false;

function initApprovalPage() {
  // 方案切换
  document.querySelectorAll('#approval-tier-tabs .tier-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('#approval-tier-tabs .tier-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      _approvalTier = tab.dataset.tier;
      // 如果已有结果，重新生成
      if (_lastAdvisorResult) generateApproval();
    });
  });

  // 生成按钮
  document.getElementById('btn-gen-approval').addEventListener('click', generateApproval);

  // 单段重新生成
  document.querySelectorAll('[data-regen]').forEach(btn => {
    btn.addEventListener('click', () => {
      const section = btn.dataset.regen;
      regenerateSection(parseInt(section));
    });
  });

  // AI 开关
  document.getElementById('ai-toggle').addEventListener('click', toggleAI);

  // 复制全文
  document.getElementById('btn-copy-approval').addEventListener('click', copyApprovalText);

  // 导出 Markdown
  document.getElementById('btn-export-md').addEventListener('click', exportApprovalMarkdown);
}

function updateApprovalState() {
  const hasResult = !!_lastAdvisorResult;
  document.getElementById('approval-empty').classList.toggle('hidden', hasResult);
  document.getElementById('approval-main').classList.toggle('hidden', !hasResult);

  if (hasResult && AppState.lastCheckupResult) {
    const o = AppState.lastCheckupResult.offer;
    document.getElementById('approval-offer-info').textContent =
      (o.country || '-') + ' · ' + (o.level || '-') + ' · ' + (o.channel || '-') + ' · base ' + formatNumber(o.baseSalary);
  }
}

function generateApproval() {
  if (!_lastAdvisorResult || !AppState.lastCheckupResult) {
    showToast('请先完成 Offer 体检', 'warning');
    return;
  }

  const { offer, context } = AppState.lastCheckupResult;
  const rec = _lastAdvisorResult.recommendation;
  const pkg = rec[_approvalTier];
  if (!pkg) {
    showToast('推荐方案不完整', 'error');
    return;
  }

  // 补充 tier 标识
  pkg.tier = _approvalTier;

  const factors = _lastAdvisorResult.factors || [];
  const risks = _lastAdvisorResult.risks || [];

  _approvalStatement = generateApprovalStatement(offer, pkg, context, factors, risks);

  // 填充三个文本域
  document.getElementById('approval-s1').value = _approvalStatement.section1;
  document.getElementById('approval-s2').value = _approvalStatement.section2;
  document.getElementById('approval-s3').value = _approvalStatement.section3;

  showToast('审批说明已生成（' + { conservative: '保守', standard: '标准', aggressive: '进取' }[_approvalTier] + '档）', 'success');
}

function regenerateSection(sectionNum) {
  if (!_lastAdvisorResult || !AppState.lastCheckupResult) {
    showToast('请先生成完整审批说明', 'warning');
    return;
  }

  const { offer, context } = AppState.lastCheckupResult;
  const pkg = _lastAdvisorResult.recommendation[_approvalTier];
  if (!pkg) return;
  pkg.tier = _approvalTier;

  const factors = _lastAdvisorResult.factors || [];
  const risks = _lastAdvisorResult.risks || [];

  const fresh = generateApprovalStatement(offer, pkg, context, factors, risks);

  if (sectionNum === 1) {
    document.getElementById('approval-s1').value = fresh.section1;
    _approvalStatement.section1 = fresh.section1;
  } else if (sectionNum === 2) {
    document.getElementById('approval-s2').value = fresh.section2;
    _approvalStatement.section2 = fresh.section2;
  } else if (sectionNum === 3) {
    document.getElementById('approval-s3').value = fresh.section3;
    _approvalStatement.section3 = fresh.section3;
  }

  showToast('第 ' + sectionNum + ' 段已重新生成', 'success');
}

function toggleAI() {
  _aiEnabled = !_aiEnabled;
  const toggle = document.getElementById('ai-toggle');
  const label = document.getElementById('ai-status-label');
  const warning = document.getElementById('ai-warning');

  toggle.classList.toggle('on', _aiEnabled);
  label.textContent = _aiEnabled ? '已开启 · 调用前自动脱敏' : '已关闭 · 完全离线模式';
  warning.classList.toggle('hidden', !_aiEnabled);

  if (_aiEnabled) {
    showToast('AI 润色已开启，发送前将自动脱敏', 'warning');
  }
}

async function aiPolishCurrentStatement() {
  if (!_aiEnabled) {
    showToast('请先开启 AI 润色', 'warning');
    return;
  }

  const s1 = document.getElementById('approval-s1').value;
  const s2 = document.getElementById('approval-s2').value;
  const s3 = document.getElementById('approval-s3').value;
  const fullText = s1 + '\n\n' + s2 + '\n\n' + s3;

  const offer = AppState.lastCheckupResult ? AppState.lastCheckupResult.offer : {};

  showLoading('正在脱敏并调用 AI 润色...');
  try {
    const polished = await aiPolishStatement(fullText, offer);
    const parts = polished.split(/\n{2,}/);
    if (parts.length >= 3) {
      document.getElementById('approval-s1').value = parts[0];
      document.getElementById('approval-s2').value = parts[1];
      document.getElementById('approval-s3').value = parts.slice(2).join('\n\n');
    } else {
      document.getElementById('approval-s1').value = polished;
    }
    hideLoading();
    showToast('AI 润色完成', 'success');
  } catch (err) {
    hideLoading();
    showToast('AI 润色失败：' + err.message, 'error');
  }
}

function copyApprovalText() {
  const s1 = document.getElementById('approval-s1').value;
  const s2 = document.getElementById('approval-s2').value;
  const s3 = document.getElementById('approval-s3').value;
  const fullText = s1 + '\n\n' + s2 + '\n\n' + s3;

  if (!fullText.trim()) {
    showToast('请先生成审批说明', 'warning');
    return;
  }

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(fullText).then(() => {
      showToast('已复制到剪贴板', 'success');
    }).catch(() => {
      fallbackCopy(fullText);
    });
  } else {
    fallbackCopy(fullText);
  }
}

function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.cssText = 'position:fixed;left:-9999px';
  document.body.appendChild(ta);
  ta.select();
  document.execCommand('copy');
  document.body.removeChild(ta);
  showToast('已复制到剪贴板', 'success');
}

function exportApprovalMarkdown() {
  const s1 = document.getElementById('approval-s1').value;
  const s2 = document.getElementById('approval-s2').value;
  const s3 = document.getElementById('approval-s3').value;

  if (!s1.trim() && !s2.trim() && !s3.trim()) {
    showToast('请先生成审批说明', 'warning');
    return;
  }

  const md = '## 一、薪酬方案\n\n' + s1 + '\n\n## 二、能力匹配\n\n' + s2 + '\n\n## 三、预算与风险\n\n' + s3;
  const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '审批说明_' + new Date().toISOString().slice(0, 10) + '.md';
  a.click();
  URL.revokeObjectURL(url);
  showToast('Markdown 已导出', 'success');
}