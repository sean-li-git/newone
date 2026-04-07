// ========== EPIPE 防护 ==========
// 打包为 .app 双击运行时无终端接收 stdout/stderr，任何写入都会触发 EPIPE。
// 必须在 require('electron') 之前完成，因为 Electron 内部初始化也可能写 stdout。

// 1) 将 stdout/stderr.write 替换为安全的空操作
const _safeWrite = function() { return true; };
if (process.stdout) {
  process.stdout.write = _safeWrite;
  if (typeof process.stdout.on === 'function') process.stdout.on('error', () => {});
}
if (process.stderr) {
  process.stderr.write = _safeWrite;
  if (typeof process.stderr.on === 'function') process.stderr.on('error', () => {});
}

// 2) 覆盖 console 方法（它们内部调用 stdout.write，是 EPIPE 的主要触发源）
const _noop = () => {};
console.log = _noop;
console.error = _noop;
console.warn = _noop;
console.info = _noop;
console.debug = _noop;

// 3) 兜底：捕获未捕获的 EPIPE 异常
process.on('uncaughtException', (err) => {
  if (err && (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED')) return;
});

const {
  app,
  BrowserWindow,
  globalShortcut,
  ipcMain,
  clipboard,
  Tray,
  Menu,
  nativeImage,
  screen,
  dialog,
} = require('electron');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');

// 4) 覆盖 Electron 的错误弹窗，拦截漏网的 EPIPE
const _origShowErrorBox = dialog.showErrorBox;
dialog.showErrorBox = (title, content) => {
  if (content && (content.includes('EPIPE') || content.includes('ERR_STREAM_DESTROYED'))) return;
  _origShowErrorBox(title, content);
};

let mainWindow = null;
let bubbleWindow = null;
let tray = null;
let clipboardWatcher = null;
let lastClipboardText = '';
let originalText = '';
let bubbleTimeout = null;
let isQuitting = false;
let lastActiveAppBundleId = ''; // 记录用户触发润色前的活跃应用 bundle ID

// ========== 创建主润色窗口 ==========
function createMainWindow() {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.show();
    mainWindow.focus();
    return;
  }

  mainWindow = new BrowserWindow({
    width: 560,
    height: 500,
    show: true,
    center: true,
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: { x: 12, y: 12 },
    resizable: true,
    minWidth: 420,
    minHeight: 380,
    backgroundColor: '#161623',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  mainWindow.loadFile('index.html');

  // 关闭窗口时隐藏而非退出（macOS 标准行为）
  mainWindow.on('close', (e) => {
    if (!isQuitting) {
      e.preventDefault();
      mainWindow.hide();
      // 确保 Dock 图标保持可见
      if (process.platform === 'darwin' && app.dock) {
        app.dock.show();
      }
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ========== 创建气泡浮窗 ==========
function createBubbleWindow() {
  if (bubbleWindow && !bubbleWindow.isDestroyed()) return;

  bubbleWindow = new BrowserWindow({
    width: 40,
    height: 40,
    show: false,
    frame: false,
    transparent: true,
    resizable: false,
    skipTaskbar: true,
    hasShadow: false,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  bubbleWindow.loadFile('bubble.html');
  bubbleWindow.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  bubbleWindow.setAlwaysOnTop(true, 'screen-saver');

  bubbleWindow.on('closed', () => {
    bubbleWindow = null;
  });
}

// ========== 剪贴板监听 ==========
function checkClipboard() {
  try {
    const currentText = clipboard.readText();

    if (currentText && currentText !== lastClipboardText) {
      const oldText = lastClipboardText;
      lastClipboardText = currentText;

      const lengthOk = currentText.length >= 2 && currentText.length <= 5000;
      const textOk = looksLikeText(currentText);

      if (lengthOk && textOk) {
        // 记录当前活跃应用的 bundle ID（用于采纳后切回）
        // 使用 execFile + stdio 隔离，避免子进程继承父进程无效的 stdout fd 导致 EPIPE
        try {
          execFile('/usr/bin/osascript', ['-e', 'tell application "System Events" to get bundle identifier of first application process whose frontmost is true'], { stdio: ['ignore', 'pipe', 'pipe'] }, (err, stdout) => {
            if (!err && stdout && stdout.trim()) {
              lastActiveAppBundleId = stdout.trim();
            }
          });
        } catch (e) {}

        // 主窗口在前台（有焦点）时不弹气泡，但仍然记录文本变化
        const mainFocused = mainWindow && !mainWindow.isDestroyed() && mainWindow.isFocused();
        if (mainFocused) {
          return;
        }
        showBubble(currentText);
      }
    }
  } catch (e) {
    // 剪贴板读取失败，静默忽略
  }
}

// ========== 显示气泡 ==========
function showBubble(text) {
  // 气泡窗口不存在则重建
  if (!bubbleWindow || bubbleWindow.isDestroyed()) {
    createBubbleWindow();
    // 等待窗口加载完成后再显示
    bubbleWindow.webContents.once('did-finish-load', () => {
      positionAndShowBubble(text);
    });
    return;
  }

  positionAndShowBubble(text);
}

function positionAndShowBubble(text) {
  if (!bubbleWindow || bubbleWindow.isDestroyed()) {
    return;
  }

  if (bubbleTimeout) {
    clearTimeout(bubbleTimeout);
    bubbleTimeout = null;
  }

  originalText = text;

  const mousePos = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(mousePos);
  const { x: dx, y: dy, width: dw, height: dh } = display.workArea;

  let x = mousePos.x + 10;
  let y = mousePos.y - 45;

  if (x + 40 > dx + dw) x = mousePos.x - 50;
  if (y < dy) y = mousePos.y + 10;
  if (y + 40 > dy + dh) y = dy + dh - 50;

  bubbleWindow.setBounds({ x: Math.round(x), y: Math.round(y), width: 40, height: 40 });

  // 确保窗口在最顶层显示
  bubbleWindow.setAlwaysOnTop(true, 'screen-saver');
  bubbleWindow.showInactive();

  // 验证窗口确实显示了
  setTimeout(() => {
    if (bubbleWindow && !bubbleWindow.isDestroyed() && !bubbleWindow.isVisible()) {
      bubbleWindow.showInactive();
    }
  }, 100);

  bubbleTimeout = setTimeout(() => {
    hideBubble();
  }, 8000);
}

function hideBubble() {
  if (bubbleTimeout) {
    clearTimeout(bubbleTimeout);
    bubbleTimeout = null;
  }
  if (bubbleWindow && !bubbleWindow.isDestroyed() && bubbleWindow.isVisible()) {
    bubbleWindow.hide();
  }
}

// ========== 显示主窗口（润色） ==========
async function showPolishWindow(text) {
  hideBubble();
  await ensureMainWindowReady();
  mainWindow.show();
  mainWindow.focus();
  mainWindow.webContents.send('start-polish', text);
}

// ========== 打开手动输入模式 ==========
async function openInputMode() {
  hideBubble();
  await ensureMainWindowReady();
  mainWindow.center();
  mainWindow.show();
  mainWindow.focus();
  mainWindow.webContents.send('open-input-mode');
}

// 确保主窗口存在且渲染进程已完成初始化
let _rendererReadyCallbacks = [];

ipcMain.on('renderer-ready', () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow._rendererReady = true;
  }
  const cbs = _rendererReadyCallbacks;
  _rendererReadyCallbacks = [];
  cbs.forEach(cb => cb());
});

function ensureMainWindowReady() {
  return new Promise((resolve) => {
    if (mainWindow && !mainWindow.isDestroyed() && mainWindow._rendererReady) {
      resolve();
      return;
    }
    if (!mainWindow || mainWindow.isDestroyed()) {
      createMainWindow();
    }
    _rendererReadyCallbacks.push(resolve);
  });
}

// ========== 剪贴板监听启动/停止 ==========
function startSelectionWatcher() {
  lastClipboardText = clipboard.readText();

  clipboardWatcher = setInterval(() => {
    checkClipboard();
  }, 300);
}

function looksLikeText(text) {
  if (/^__TP_\d+__$/.test(text)) return false;
  if (/^https?:\/\/\S+$/.test(text)) return false;
  if (/^[\/~][\w\/\-.]+$/.test(text)) return false;
  if (/^\d+$/.test(text)) return false;
  if (/[\u4e00-\u9fff]/.test(text)) return true;
  if (text.length > 10) return true;
  return false;
}

function stopSelectionWatcher() {
  if (clipboardWatcher) {
    clearInterval(clipboardWatcher);
    clipboardWatcher = null;
  }
}

// ========== 托盘 ==========
function createTray() {
  try {
    // 方案：直接从 icon.png 读取并 resize 为菜单栏尺寸
    const iconPath = path.join(__dirname, 'icon.png');
    const bigIcon = nativeImage.createFromPath(iconPath);

    if (bigIcon.isEmpty()) {
      return;
    }

    // macOS 菜单栏 @2x Retina 推荐 32x32 像素（显示为 16x16 点）
    const trayIcon = bigIcon.resize({ width: 18, height: 18 });

    tray = new Tray(trayIcon);
    tray.setToolTip('语言润色 - 监听中');
    updateTrayMenu();
    tray.on('click', () => openInputMode());
  } catch (e) {
    // 托盘创建失败，静默忽略
  }
}

// ========== 设置 Dock 图标 ==========
function setupDockIcon() {
  if (process.platform === 'darwin' && app.dock) {
    // 确保 Dock 图标始终可见
    app.dock.show();
    const dockIconPath = path.join(__dirname, 'icon.png');
    const dockIcon = nativeImage.createFromPath(dockIconPath);
    if (!dockIcon.isEmpty()) {
      app.dock.setIcon(dockIcon);
    }
  }
}

function updateTrayMenu() {
  if (!tray) return;
  const isWatching = !!clipboardWatcher;

  const contextMenu = Menu.buildFromTemplate([
    {
      label: '打开输入窗口  ⌘⇧L',
      click: () => openInputMode(),
    },
    { type: 'separator' },
    {
      label: isWatching ? '⏸ 暂停剪贴板监听' : '▶ 恢复剪贴板监听',
      click: () => {
        if (isWatching) {
          stopSelectionWatcher();
          tray.setToolTip('语言润色 - 已暂停');
        } else {
          startSelectionWatcher();
          tray.setToolTip('语言润色 - 监听中');
        }
        updateTrayMenu();
      },
    },
    { type: 'separator' },
    {
      label: '退出',
      click: () => {
        isQuitting = true;
        stopSelectionWatcher();
        if (mainWindow && !mainWindow.isDestroyed()) mainWindow.destroy();
        if (bubbleWindow && !bubbleWindow.isDestroyed()) bubbleWindow.destroy();
        app.quit();
      },
    },
  ]);

  tray.setContextMenu(contextMenu);
}

// ========== 采纳替换：复制到剪贴板 + 切回原应用 + 浮动提示 ==========
let toastWindow = null;

function adoptAndReplace(newText) {
  clipboard.writeText(newText);
  lastClipboardText = newText;

  // 隐藏主窗口
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.hide();

  // 切回原应用（使用 execFile + stdio 隔离，避免 EPIPE）
  if (lastActiveAppBundleId) {
    execFile('/usr/bin/open', ['-b', lastActiveAppBundleId], { stdio: ['ignore', 'pipe', 'pipe'], timeout: 3000 }, () => {});
  }

  // 显示浮动提示
  showFloatingToast('已复制，按 ⌘V 粘贴');
}

function showFloatingToast(message) {
  // 清理旧的 toast 窗口
  if (toastWindow && !toastWindow.isDestroyed()) {
    toastWindow.destroy();
    toastWindow = null;
  }

  const mousePos = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(mousePos);
  const { width: dw, y: dy } = display.workArea;

  // 在屏幕顶部居中显示
  const toastWidth = 220;
  const toastHeight = 44;
  const x = Math.round(display.workArea.x + (dw - toastWidth) / 2);
  const y = dy + 12;

  toastWindow = new BrowserWindow({
    width: toastWidth,
    height: toastHeight,
    x,
    y,
    show: false,
    frame: false,
    transparent: true,
    resizable: false,
    skipTaskbar: true,
    hasShadow: true,
    alwaysOnTop: true,
    focusable: false,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  toastWindow.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  toastWindow.setAlwaysOnTop(true, 'screen-saver');

  const html = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8">
    <style>
      * { margin: 0; padding: 0; }
      body {
        -webkit-app-region: no-drag;
        background: transparent;
        display: flex; justify-content: center; align-items: center;
        height: 100vh; font-family: -apple-system, sans-serif;
      }
      .toast {
        background: rgba(0,0,0,0.82);
        color: #fff;
        padding: 10px 22px;
        border-radius: 10px;
        font-size: 14px;
        font-weight: 500;
        letter-spacing: 0.3px;
        white-space: nowrap;
        animation: fadeInOut 2.5s ease-in-out;
      }
      @keyframes fadeInOut {
        0% { opacity: 0; transform: translateY(-8px); }
        12% { opacity: 1; transform: translateY(0); }
        75% { opacity: 1; }
        100% { opacity: 0; }
      }
    </style>
    </head>
    <body><div class="toast">${message}</div></body>
    </html>
  `;

  toastWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  toastWindow.once('ready-to-show', () => {
    toastWindow.showInactive();
  });

  // 2.5秒后自动关闭
  setTimeout(() => {
    if (toastWindow && !toastWindow.isDestroyed()) {
      toastWindow.destroy();
      toastWindow = null;
    }
  }, 2600);
}

// ========== 启动 ==========
app.whenReady().then(() => {
  setupDockIcon();
  createMainWindow();
  createBubbleWindow();
  createTray();
  startSelectionWatcher();

  globalShortcut.register('CommandOrControl+Shift+L', () => {
    if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible()) {
      mainWindow.hide();
    } else {
      openInputMode();
    }
  });
});

app.on('before-quit', () => {
  isQuitting = true;
});

app.on('will-quit', () => {
  stopSelectionWatcher();
  globalShortcut.unregisterAll();
});

app.on('window-all-closed', () => {
  // macOS 上关闭所有窗口不退出应用
  if (process.platform !== 'darwin') app.quit();
});

// 点击 Dock 图标时打开输入窗口
app.on('activate', () => {
  // 如果主窗口存在但被隐藏了，直接 show
  if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isVisible()) {
    mainWindow.center();
    mainWindow.show();
    mainWindow.focus();
    return;
  }
  openInputMode();
});

// ========== IPC 通信 ==========

ipcMain.on('bubble-clicked', () => {
  if (originalText) {
    showPolishWindow(originalText);
  }
});

ipcMain.on('bubble-dismiss', () => {
  hideBubble();
});

ipcMain.on('adopt-and-replace', (event, newText) => {
  adoptAndReplace(newText);
});

ipcMain.on('copy-to-clipboard', (event, text) => {
  clipboard.writeText(text);
  lastClipboardText = text; // 防止复制结果再次触发气泡
});

ipcMain.on('hide-window', () => {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.hide();
});

ipcMain.on('close-window', () => {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.hide();
});

// ========== API 设置持久化（文件存储，替代 localStorage） ==========

function getSettingsPath() {
  return path.join(app.getPath('userData'), 'api-settings.json');
}

function getCustomPromptsPath() {
  return path.join(app.getPath('userData'), 'custom-prompts.json');
}

function readSettings() {
  try {
    const data = fs.readFileSync(getSettingsPath(), 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return {};
  }
}

function writeSettings(settings) {
  try {
    fs.writeFileSync(getSettingsPath(), JSON.stringify(settings, null, 2), 'utf-8');
  } catch (e) {
    // 写入失败静默忽略
  }
}

function readCustomPrompts() {
  try {
    const data = fs.readFileSync(getCustomPromptsPath(), 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return {};
  }
}

function writeCustomPrompts(prompts) {
  try {
    fs.writeFileSync(getCustomPromptsPath(), JSON.stringify(prompts, null, 2), 'utf-8');
  } catch (e) {
    // 写入失败静默忽略
  }
}

// IPC: 读取设置
ipcMain.handle('get-settings', () => {
  return readSettings();
});

// IPC: 保存设置
ipcMain.on('save-settings', (event, settings) => {
  writeSettings(settings);
});

// IPC: 读取自定义 Prompt
ipcMain.handle('get-custom-prompts', () => {
  return readCustomPrompts();
});

// IPC: 保存自定义 Prompt
ipcMain.on('save-custom-prompts', (event, prompts) => {
  writeCustomPrompts(prompts);
});

// ========== 润色历史记录 ==========

const MAX_HISTORY = 100;

function getHistoryPath() {
  return path.join(app.getPath('userData'), 'polish-history.json');
}

function readHistory() {
  try {
    const data = fs.readFileSync(getHistoryPath(), 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return [];
  }
}

function writeHistory(history) {
  try {
    fs.writeFileSync(getHistoryPath(), JSON.stringify(history, null, 2), 'utf-8');
  } catch (e) {
    // 写入失败静默忽略
  }
}

// 保存一条润色记录
ipcMain.on('save-polish-history', (event, record) => {
  const history = readHistory();
  history.push({
    original: record.original,
    result: record.result,
    mode: record.mode,
    starred: false,
    timestamp: Date.now(),
  });
  // 超过上限时淘汰最旧的非加星记录
  while (history.length > MAX_HISTORY) {
    const idx = history.findIndex(h => !h.starred);
    if (idx >= 0) {
      history.splice(idx, 1);
    } else {
      history.shift();
    }
  }
  writeHistory(history);
});

// 获取润色历史（供 few-shot 和历史面板使用）
ipcMain.handle('get-polish-history', () => {
  return readHistory();
});

// 删除一条记录
ipcMain.on('delete-polish-history', (event, timestamp) => {
  let history = readHistory();
  history = history.filter(h => h.timestamp !== timestamp);
  writeHistory(history);
});

// 切换加星状态
ipcMain.on('toggle-star-history', (event, timestamp) => {
  const history = readHistory();
  const item = history.find(h => h.timestamp === timestamp);
  if (item) {
    item.starred = !item.starred;
    writeHistory(history);
  }
});
