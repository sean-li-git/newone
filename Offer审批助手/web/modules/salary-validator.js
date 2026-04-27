/* ======================================================================
 * Offer审批助手 · 薪酬校验器
 * 自动核算 + 完整性检查 + 字段异常检测
 * ====================================================================== */

/**
 * 完整性检查：哪些必填字段缺失
 */
function checkCompleteness(offer) {
  const issues = [];
  
  // 必填元信息
  if (!offer.country) issues.push({ severity: 'error', field: 'country', message: '「国家/城市」未填写' });
  if (!offer.level)   issues.push({ severity: 'error', field: 'level',   message: '「职级」未填写' });
  if (!offer.channel) issues.push({ severity: 'warning', field: 'channel', message: '「职位通道」未填写（影响画像精度）' });
  if (!offer.jobFamily) issues.push({ severity: 'warning', field: 'jobFamily', message: '「职位类型」未填写（影响画像精度）' });
  
  // 必填薪酬
  if (!offer.baseSalary || offer.baseSalary <= 0) {
    issues.push({ severity: 'error', field: 'baseSalary', message: '「基本月薪」为空或为 0' });
  }
  
  // 可选但建议填写
  if (!offer.currency) issues.push({ severity: 'info', field: 'currency', message: '「币种」未填写，默认 CNY' });
  
  return issues;
}

/**
 * 数值合理性检查
 */
function checkReasonableness(offer) {
  const issues = [];
  const base = Number(offer.baseSalary) || 0;
  
  // 月薪范围检查（宽松范围）
  if (base > 0 && base < 1000) {
    issues.push({ severity: 'warning', field: 'baseSalary', message: `基本月薪 ${base} 偏低，请确认单位是否正确` });
  }
  if (base > 500000) {
    issues.push({ severity: 'warning', field: 'baseSalary', message: `基本月薪 ${base} 异常高，请确认` });
  }
  
  // 绩效合理性
  const perf = Number(offer.perfMonthly) || 0;
  if (perf > 0 && base > 0 && perf > base * 2) {
    issues.push({ severity: 'warning', field: 'perfMonthly', message: `月绩效 ${perf} 超过月薪 2 倍，请确认` });
  }
  
  // 年终奖合理性
  const bonus = Number(offer.annualBonus) || 0;
  if (bonus > 0 && base > 0 && bonus > base * 12) {
    issues.push({ severity: 'warning', field: 'annualBonus', message: `年终奖 ${bonus} 超过年薪，请确认` });
  }
  
  // 一次性费用占比检查
  const signTotal = (Number(offer.signStock) || 0) + (Number(offer.signBonus) || 0) + (Number(offer.relocation) || 0);
  const annual = Number(offer.annualIncome) || (base * 12);
  if (signTotal > 0 && annual > 0 && signTotal > annual * 3) {
    issues.push({ severity: 'warning', field: 'signBonus', message: `一次性费用总计 ${signTotal}，超过年收入 3 倍，请确认` });
  }
  
  return issues;
}

/**
 * 综合校验 — 合并所有检查结果
 * @returns {{ completeness: Array, reasonableness: Array, computed: Array, allIssues: Array }}
 */
function validateOffer(offer, amortizationYears) {
  const completeness = checkCompleteness(offer);
  const reasonableness = checkReasonableness(offer);
  const computed = validateComputedFields(offer, amortizationYears);
  
  const computedIssues = computed
    .filter(c => !c.match)
    .map(c => ({
      severity: 'warning',
      field: c.key,
      message: `${c.label}：Excel 值 ${c.excelValue}，系统计算 ${c.computedValue}（差异 ${c.diff > 0 ? '+' : ''}${c.diff}）`,
    }));
  
  const allIssues = [...completeness, ...reasonableness, ...computedIssues];
  
  return { completeness, reasonableness, computed, allIssues };
}

/**
 * 获取校验摘要统计
 */
function getValidationSummary(validation) {
  const errors = validation.allIssues.filter(i => i.severity === 'error').length;
  const warnings = validation.allIssues.filter(i => i.severity === 'warning').length;
  const infos = validation.allIssues.filter(i => i.severity === 'info').length;
  
  let status = 'pass'; // 通过
  if (errors > 0) status = 'fail';
  else if (warnings > 0) status = 'warning';
  
  return { status, errors, warnings, infos, total: validation.allIssues.length };
}
