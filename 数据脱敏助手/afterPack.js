// afterPack.js — electron-builder afterPack 钩子
// 功能：
// 1. 用 shell 包装器替换 Electron 可执行文件，重定向 stdout/stderr 杜绝 EPIPE 错误
// 2. 移除 app 的代码签名，让别人的 Mac 可以通过右键"打开"绕过 Gatekeeper
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

module.exports = async function (context) {
  // 只处理 macOS
  if (context.electronPlatformName !== 'darwin') return;

  const appName = context.packager.appInfo.productFilename; // "数据脱敏助手"
  const appOutDir = context.appOutDir;
  const appPath = path.join(appOutDir, `${appName}.app`);
  const exePath = path.join(appPath, 'Contents', 'MacOS', appName);

  if (!fs.existsSync(exePath)) return;

  // ===== 1. Shell 包装器（防 EPIPE） =====
  const realExePath = exePath + '.real';
  fs.renameSync(exePath, realExePath);

  const wrapper = `#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$DIR/${appName}.real" "$@" >/dev/null 2>&1
`;
  fs.writeFileSync(exePath, wrapper, { mode: 0o755 });

  // ===== 2. 移除代码签名（关键：让 Gatekeeper 放行） =====
  try {
    // 移除 app bundle 的签名
    execSync(`codesign --remove-signature "${appPath}"`, { stdio: 'ignore' });
    // 同时移除内部 Electron Framework 等二进制的签名
    const frameworksPath = path.join(appPath, 'Contents', 'Frameworks');
    if (fs.existsSync(frameworksPath)) {
      const items = fs.readdirSync(frameworksPath);
      for (const item of items) {
        const itemPath = path.join(frameworksPath, item);
        try {
          execSync(`codesign --remove-signature "${itemPath}"`, { stdio: 'ignore' });
        } catch (e) {
          // 某些文件可能不是签名的，忽略
        }
      }
    }
    // 移除 Electron Helper 的签名
    const helpersGlob = path.join(appPath, 'Contents', 'Frameworks', '*.app');
    try {
      const helperApps = execSync(`ls -d "${helpersGlob}" 2>/dev/null || true`, { encoding: 'utf8' }).trim().split('\n').filter(Boolean);
      for (const helperApp of helperApps) {
        execSync(`codesign --remove-signature "${helperApp}"`, { stdio: 'ignore' });
      }
    } catch (e) {}
  } catch (e) {
    // codesign 命令失败不阻断打包
  }
};
