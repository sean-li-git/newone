// afterPack.js — electron-builder afterPack 钩子
// 功能：用 shell 包装器替换 Electron 可执行文件，
//       在操作系统层面将 stdout/stderr 重定向到 /dev/null，
//       从根源杜绝 EPIPE 错误弹窗。
const fs = require('fs');
const path = require('path');

module.exports = async function (context) {
  // 只处理 macOS
  if (context.electronPlatformName !== 'darwin') return;

  const appName = context.packager.appInfo.productFilename; // "数据脱敏助手"
  const appOutDir = context.appOutDir;
  const exePath = path.join(appOutDir, `${appName}.app`, 'Contents', 'MacOS', appName);

  if (!fs.existsSync(exePath)) return;

  // 将原始可执行文件重命名
  const realExePath = exePath + '.real';
  fs.renameSync(exePath, realExePath);

  // 创建 shell 包装脚本，将 stdout/stderr 重定向到 /dev/null
  const wrapper = `#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$DIR/${appName}.real" "$@" >/dev/null 2>&1
`;

  fs.writeFileSync(exePath, wrapper, { mode: 0o755 });
};
