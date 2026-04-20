# 数据脱敏助手 - 开发日志

## 版本索引

| 版本 | 日期 | 核心内容 |
|------|------|---------|
| v1.4.1 | 2026-04-20 | 还原 0 单元格时显示红色失败提示 |
| v1.4 | 2026-04-20 | 新增网页版（单文件 HTML，免安装） |
| v1.3 | 2026-04-20 | 修复年月和小数值列的识别问题 |
| v1.2 | 2026-04-10 | DMG安装包 + ad-hoc签名 |
| v1.1 | 2026-04-08 | 操作报告功能 |
| v1.0 | 2026-03-12 | 首个正式版本 |

---

## v1.4.1 — 2026-04-20

### 核心迭代内容
- 还原报告增强：当 `report.totalRestored === 0` 时，把"✅ 还原完成"改为"⚠️ 还原未成功"，并在 summary 下方增加红色提示框列出常见失败原因

### 技术要点
- 修改点：`web/app.js` 和 `index.html` 的 `renderRestoreReport` 函数
- 仅做"事后诊断提示"，**不改变任何匹配逻辑**——保持"精确匹配失败好过模糊匹配成功"的安全原则
- 放弃了按列名模糊匹配 / 按特征猜 sheet 的改进方案，因为静默的数据错误比明显失败危险得多

### 为什么不做更激进的改进
评估过两个"更智能"的方案：
1. 按列名匹配而非按列位置匹配 → 列名被改就匹配不上，且重名列会乱匹配
2. 按列组合特征模糊匹配 sheet → 一旦猜错，输出的数据看起来正常但其实全错

脱敏还原场景下，**精确失败 + 明显提示** > **模糊成功但用户不知道错了**。

---

## v1.4 — 2026-04-20

### 核心迭代内容
- 新增 `web/` 目录，提供单文件 HTML 版分发产物 `数据脱敏助手.html`
- 技术选型：前端用 SheetJS (xlsx.full.min.js) 替代 Electron 版的 ExcelJS，在浏览器中读写 Excel
- 业务逻辑完整移植自 `main.js`：列识别引擎、9 种脱敏策略、基准密钥模式、公式跳过
- localStorage 记忆列名→策略偏好（key=`desensitizer_prefs_v1`），跨任务复用
- 构建方式：`app.html`（结构+CSS）+ `app.js`（逻辑）+ `xlsx.full.min.js`（库）三文件分离维护，通过 `build.js` 合成单文件产物

### 技术要点
- 文件读取：`FileReader.readAsArrayBuffer` → `XLSX.read(data, { type:'array', cellFormula:true, cellStyles:false })`
- 文件导出：`XLSX.write(wb, { type:'array', bookType:'xlsx' })` → `new Blob([...])` → `URL.createObjectURL` → `<a download>`
- SheetJS 的 cell 对象 `{ t:'n/s/d', v:value, f:formula }` 与 ExcelJS 的 `cell.value` / `cell.formula` 做了适配映射
- CSS 直接复刻 Electron 版 `index.html` 的暗色主题，保持视觉一致
- 单文件产物约 1 MB（SheetJS 压缩版 930 KB + 业务代码 70 KB）

### 为什么做 HTML 版
- 桌面版 DMG 安装包在公司/同事间分享时，首次打开需要"右键→打开"，有一定理解成本
- Windows 用户无法使用 macOS DMG
- HTML 版双击即用，发微信/邮件/U盘都能直接打开，真正做到免安装、跨平台

### 遗留问题
- SheetJS 社区版对某些复杂的 Excel 样式/图表不完全支持（脱敏场景下影响很小，只涉及单元格值）
- 目前只支持 .xlsx，不支持 .xls（同 Electron 版）

---

## v1.3 — 2026-04-20

### 核心迭代内容
- 修复 `detectColumnType` 对「年-月」格式（`2025-01`、`2025/3`、`2025年1月`）的识别，增加 `^\d{4}[\-\/年]\s?\d{1,2}[月]?$` 正则，命中则走 `date-offset` 策略
- 修复小数值列被 fallback 到文本 mapping 的问题：数字占比高时，无论大小都走 `scale`，不再因为 `max <= 1000` 退出数字分支
- 调整判断顺序：先判断年月 → 再判断完整日期 → 最后判断纯数字，避免"年-月"被 `new Date()` 解析为合法日期后影响其他分支

### 技术要点
- 端到端测试（读→识别→脱敏→导出→再读→还原→对比）验证 0 差异
- `numValues` 过滤增加 `String(v).trim() !== ''` 防止空字符串被 Number() 识别为 0

### 遗留问题
- 无

---

## v1.2 — 2026-04-10

### 核心迭代内容
- 打包格式从 zip 改为 DMG，提供标准 macOS 安装体验（拖入 Applications）
- afterPack.js 中用 `codesign --force --deep --sign -` 做 ad-hoc 签名，替代原来的 `codesign --remove-signature`
- 解决通过微信/邮件/AirDrop 传输后 macOS Gatekeeper 提示"已损坏，应移到废纸篓"的问题

### 技术要点
- macOS 对网络传输文件打 `com.apple.quarantine` 扩展属性，完全去签名的 app 会被直接拦截删除
- ad-hoc 签名（`sign -`）虽不被 Apple 公证，但允许用户右键→打开绕过 Gatekeeper
- DMG 格式比 zip 更标准，自带 Applications 快捷方式引导安装
- 签名顺序：先签内部 Frameworks/Helper → 再签整个 app bundle

### 遗留问题
- 无

---

## v1.1 — 2026-04-08

### 核心迭代内容
- 脱敏和还原完成后，后端在返回结果中附带 `report` 统计信息
- 前端在结果面板中以卡片形式展示操作报告（表格样式）

### 技术要点
- `main.js` 脱敏函数和还原函数各新增报告生成逻辑，遍历 keyData 统计每列处理/跳过的单元格数
- `index.html` 新增报告渲染函数 `renderDesensitizeReport()` / `renderRestoreReport()`
- 报告 CSS 使用与现有主题一致的暗色风格

### 遗留问题
- 无

---

## v1.0 — 2026-03-12

### 核心迭代内容
- 完整的 Excel 数据脱敏与还原流程
- 自动列类型识别引擎
- 8 种脱敏策略实现
- 密钥文件管理（导出 JSON / Excel 映射表）
- 基准密钥复用机制
- macOS 打包与 Gatekeeper 处理

### 技术卡点
- Electron 打包后 EPIPE 错误 → 通过 shell 包装器和 stdout 重定向解决
- macOS Gatekeeper 拦截未签名应用 → afterPack.js 中移除所有代码签名
- ExcelJS 公式单元格处理 → 读取时记录公式，导出时写回

### 遗留问题
- 无
