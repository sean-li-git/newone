#!/usr/bin/env node
/**
 * 构建脚本：把 xlsx.full.min.js、chart.min.js 和所有模块 JS 内嵌到 app.html 的占位符中，
 * 产出单文件分发版 Offer审批助手.html
 *
 * 用法：node build.js
 */
const fs = require('fs');
const path = require('path');

const WEB_DIR = __dirname;
const HTML_IN  = path.join(WEB_DIR, 'app.html');
const XLSX_JS  = path.join(WEB_DIR, 'lib', 'xlsx.full.min.js');
const CHART_JS = path.join(WEB_DIR, 'lib', 'chart.min.js');
const OUT_FILE = path.join(WEB_DIR, 'Offer审批助手.html');

// 模块文件（按依赖顺序）
const MODULE_FILES = [
  'modules/storage.js',
  'modules/excel-parser.js',
  'modules/salary-validator.js',
  'modules/rule-engine.js',
  'modules/rule-library.js',
  'modules/sample-wizard.js',
  'modules/profile-learner.js',
  'modules/profile-store.js',
  'modules/advisor/factor-baseline.js',
  'modules/advisor/factor-raise.js',
  'modules/advisor/factor-competition.js',
  'modules/advisor/factor-internal.js',
  'modules/advisor/factor-urgency.js',
  'modules/advisor/factor-ability.js',
  'modules/advisor/factor-rarity.js',
  'modules/advisor/factor-structure.js',
  'modules/advisor/advisor.js',
  'modules/approval-writer.js',
  'modules/anonymizer.js',
  'modules/ai-provider.js',
];

function read(p) {
  if (!fs.existsSync(p)) {
    console.error('❌ 找不到文件：' + p);
    process.exit(1);
  }
  return fs.readFileSync(p, 'utf8');
}

function readOptional(p) {
  if (!fs.existsSync(p)) {
    console.warn('⚠️ 跳过不存在的文件：' + p);
    return '';
  }
  return fs.readFileSync(p, 'utf8');
}

console.log('📦 开始构建单文件 HTML 版...');

let html = read(HTML_IN);

// SheetJS
const xlsx = read(XLSX_JS);
const xlsxTag = '<script>\n/* === SheetJS (xlsx.full.min.js) === */\n' + xlsx + '\n</script>';
html = html.replace('<!-- INLINE_SHEETJS_HERE -->', xlsxTag);

// Chart.js
const chart = readOptional(CHART_JS);
if (chart) {
  const chartTag = '<script>\n/* === Chart.js === */\n' + chart + '\n</script>';
  html = html.replace('<!-- INLINE_CHARTJS_HERE -->', chartTag);
} else {
  html = html.replace('<!-- INLINE_CHARTJS_HERE -->', '<!-- Chart.js: not found, skipped -->');
}

// 模块 + app.js
let allModules = '';
for (const mf of MODULE_FILES) {
  const content = readOptional(path.join(WEB_DIR, mf));
  if (content) {
    allModules += '\n/* === ' + mf + ' === */\n' + content + '\n';
  }
}
const appJs = read(path.join(WEB_DIR, 'app.js'));
allModules += '\n/* === app.js (主逻辑) === */\n' + appJs + '\n';

const appTag = '<script>\n' + allModules + '</script>';
html = html.replace('<!-- INLINE_APP_JS_HERE -->', appTag);

fs.writeFileSync(OUT_FILE, html, 'utf8');

const sizeKB = (fs.statSync(OUT_FILE).size / 1024).toFixed(1);
console.log('✅ 构建成功：' + OUT_FILE);
console.log('📏 文件大小：' + sizeKB + ' KB');
console.log('👉 直接双击用浏览器打开即可使用，可分享给团队成员。');
