// afterPack.js — electron-builder afterPack 钩子
// 功能：
// 1. 用 shell 包装器替换 Electron 可执行文件，重定向 stdout/stderr 杜绝 EPIPE 错误
// 2. 对 app 做 ad-hoc 签名（免费，不需要开发者账号），让别人的 Mac 可通过右键"打开"绕过 Gatekeeper
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

  // ===== 2. 移除现有签名，然后做 ad-hoc 签名 =====
  try {
    // 先移除所有已有签名
    execSync(`codesign --remove-signature "${appPath}"`, { stdio: 'ignore' });

    // 对内部 Frameworks 逐个签名
    const frameworksPath = path.join(appPath, 'Contents', 'Frameworks');
    if (fs.existsSync(frameworksPath)) {
      const items = fs.readdirSync(frameworksPath);
      for (const item of items) {
        const itemPath = path.join(frameworksPath, item);
        try {
          execSync(`codesign --remove-signature "${itemPath}"`, { stdio: 'ignore' });
        } catch (e) {}
        try {
          execSync(`codesign --force --deep --sign - "${itemPath}"`, { stdio: 'ignore' });
        } catch (e) {}
      }
    }

    // 对 Helper apps 签名
    const helpersDir = path.join(appPath, 'Contents', 'Frameworks');
    try {
      const helperApps = execSync(`find "${helpersDir}" -name "*.app" -type d 2>/dev/null || true`, { encoding: 'utf8' }).trim().split('\n').filter(Boolean);
      for (const helperApp of helperApps) {
        try {
          execSync(`codesign --force --deep --sign - "${helperApp}"`, { stdio: 'ignore' });
        } catch (e) {}
      }
    } catch (e) {}

    // 最后对整个 app bundle 做 ad-hoc 签名
    execSync(`codesign --force --deep --sign - "${appPath}"`, { stdio: 'ignore' });
  } catch (e) {
    // 签名失败不阻断打包，但打印警告
    try {
      process.stderr.write(`⚠️ Ad-hoc signing failed: ${e.message}\n`);
    } catch (_) {}
  }
};
