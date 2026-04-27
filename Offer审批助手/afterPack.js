// afterPack.js — electron-builder afterPack 钩子
// 功能：
// 1. 用 shell 包装器替换 Electron 可执行文件，重定向 stdout/stderr 杜绝 EPIPE 错误
// 2. 对 app 做 ad-hoc 签名（免费，不需要开发者账号）
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

module.exports = async function (context) {
  if (context.electronPlatformName !== 'darwin') return;

  const appName = context.packager.appInfo.productFilename;
  const appOutDir = context.appOutDir;
  const appPath = path.join(appOutDir, `${appName}.app`);
  const exePath = path.join(appPath, 'Contents', 'MacOS', appName);

  if (!fs.existsSync(exePath)) return;

  // Shell 包装器
  const realExePath = exePath + '.real';
  fs.renameSync(exePath, realExePath);

  const wrapper = `#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$DIR/${appName}.real" "$@" >/dev/null 2>&1
`;
  fs.writeFileSync(exePath, wrapper, { mode: 0o755 });

  // Ad-hoc 签名
  try {
    execSync(`codesign --remove-signature "${appPath}"`, { stdio: 'ignore' });

    const frameworksPath = path.join(appPath, 'Contents', 'Frameworks');
    if (fs.existsSync(frameworksPath)) {
      const items = fs.readdirSync(frameworksPath);
      for (const item of items) {
        const itemPath = path.join(frameworksPath, item);
        try { execSync(`codesign --remove-signature "${itemPath}"`, { stdio: 'ignore' }); } catch (e) {}
        try { execSync(`codesign --force --deep --sign - "${itemPath}"`, { stdio: 'ignore' }); } catch (e) {}
      }
    }

    const helpersDir = path.join(appPath, 'Contents', 'Frameworks');
    try {
      const helperApps = execSync(`find "${helpersDir}" -name "*.app" -type d 2>/dev/null || true`, { encoding: 'utf8' }).trim().split('\n').filter(Boolean);
      for (const helperApp of helperApps) {
        try { execSync(`codesign --force --deep --sign - "${helperApp}"`, { stdio: 'ignore' }); } catch (e) {}
      }
    } catch (e) {}

    execSync(`codesign --force --deep --sign - "${appPath}"`, { stdio: 'ignore' });
  } catch (e) {
    try { process.stderr.write(`⚠️ Ad-hoc signing failed: ${e.message}\n`); } catch (_) {}
  }
};
