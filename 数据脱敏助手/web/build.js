#!/usr/bin/env node
/**
 * 构建脚本：把 xlsx.full.min.js 和 app.js 内嵌到 app.html 的占位符中，
 * 产出单文件分发版 数据脱敏助手.html
 *
 * 用法：node build.js
 */
const fs = require('fs');
const path = require('path');

const WEB_DIR = __dirname;
const HTML_IN  = path.join(WEB_DIR, 'app.html');
const XLSX_JS  = path.join(WEB_DIR, 'xlsx.full.min.js');
const APP_JS   = path.join(WEB_DIR, 'app.js');
const OUT_FILE = path.join(WEB_DIR, '数据脱敏助手.html');

function read(p){
  if (!fs.existsSync(p)) {
    console.error('❌ 找不到文件：' + p);
    process.exit(1);
  }
  return fs.readFileSync(p, 'utf8');
}

console.log('📦 开始构建单文件 HTML 版...');

let html   = read(HTML_IN);
const xlsx = read(XLSX_JS);
const app  = read(APP_JS);

// 用 <script> 包裹并替换占位符。注意：SheetJS 和 app.js 都不包含 </script>，安全。
const xlsxTag = '<script>\n/* === SheetJS (xlsx.full.min.js) === */\n' + xlsx + '\n</script>';
const appTag  = '<script>\n/* === 数据脱敏助手业务逻辑 (app.js) === */\n' + app + '\n</script>';

html = html.replace('<!-- INLINE_SHEETJS_HERE -->', xlsxTag);
html = html.replace('<!-- INLINE_APP_JS_HERE -->',  appTag);

fs.writeFileSync(OUT_FILE, html, 'utf8');

const sizeKB = (fs.statSync(OUT_FILE).size / 1024).toFixed(1);
console.log('✅ 构建成功：' + OUT_FILE);
console.log('📏 文件大小：' + sizeKB + ' KB');
console.log('👉 直接双击用浏览器打开即可使用，可分享给任何人。');
