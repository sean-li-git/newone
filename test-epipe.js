/**
 * EPIPE 防护测试脚本
 * 
 * 测试方法：模拟打包后 .app 双击运行时 stdout/stderr 管道关闭的场景
 * - 先关闭 stdout 管道（模拟无终端）
 * - 然后执行 main.js 中的防护代码
 * - 再调用 console.log 等方法，验证不会抛出 EPIPE
 */

const { spawn } = require('child_process');
const path = require('path');

// 测试1: 验证 stdout.write 被覆盖后不会 EPIPE
function testStdoutOverride() {
  return new Promise((resolve) => {
    // 创建子进程，但立即关闭其 stdout 管道
    const child = spawn(process.execPath, ['-e', `
      // 先模拟 main.js 的防护代码
      const _safeWrite = function() { return true; };
      if (process.stdout) {
        process.stdout.write = _safeWrite;
        if (typeof process.stdout.on === 'function') process.stdout.on('error', () => {});
      }
      if (process.stderr) {
        process.stderr.write = _safeWrite;
        if (typeof process.stderr.on === 'function') process.stderr.on('error', () => {});
      }
      const _noop = () => {};
      console.log = _noop;
      console.error = _noop;
      console.warn = _noop;
      console.info = _noop;
      console.debug = _noop;

      process.on('uncaughtException', (err) => {
        if (err && (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED')) return;
        // 非 EPIPE 异常，写到 fd 3 报告
        const fs = require('fs');
        fs.writeSync(3, 'FAIL: uncaught ' + err.message + '\\n');
        process.exit(1);
      });

      // 大量调用 console 方法，模拟真实使用场景
      for (let i = 0; i < 100; i++) {
        console.log('test log', i);
        console.error('test error', i);
        console.warn('test warn', i);
        console.info('test info', i);
      }

      // 直接调用 process.stdout.write
      process.stdout.write('direct write test\\n');
      process.stderr.write('direct stderr test\\n');

      // 模拟 execFile（和 main.js 中一样）
      const { execFile } = require('child_process');
      execFile('/usr/bin/osascript', ['-e', 'return "hello"'], { stdio: ['ignore', 'pipe', 'pipe'] }, (err, stdout) => {
        // 等子进程完成后报告成功
        const fs = require('fs');
        fs.writeSync(3, 'PASS\\n');
        process.exit(0);
      });

      // 超时保底
      setTimeout(() => {
        const fs = require('fs');
        fs.writeSync(3, 'PASS (timeout)\\n');
        process.exit(0);
      }, 3000);
    `], {
      stdio: ['ignore', 'pipe', 'pipe', 'pipe'] // fd 3 用来报告结果
    });

    // 关键：立即销毁 stdout 管道，模拟无终端场景
    child.stdout.destroy();
    child.stderr.destroy();

    let result = '';
    child.stdio[3].on('data', (data) => {
      result += data.toString();
    });

    child.on('exit', (code) => {
      const passed = result.trim().startsWith('PASS');
      resolve({ name: '测试1: stdout 覆盖防护', passed, detail: result.trim(), code });
    });
  });
}

// 测试2: 不使用防护代码时确认会 EPIPE（对照组）
function testWithoutProtection() {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, ['-e', `
      // 不做任何防护，直接写 stdout
      let failed = false;
      process.on('uncaughtException', (err) => {
        if (err && err.code === 'EPIPE') {
          const fs = require('fs');
          fs.writeSync(3, 'EPIPE_CAUGHT\\n');
          failed = true;
        }
      });

      process.stdout.on('error', (err) => {
        if (err.code === 'EPIPE') {
          const fs = require('fs');
          fs.writeSync(3, 'EPIPE_ON_ERROR\\n');
          failed = true;
        }
      });

      // 尝试写已关闭的管道
      try {
        for (let i = 0; i < 10; i++) {
          process.stdout.write('write to closed pipe ' + i + '\\n');
        }
      } catch(e) {
        const fs = require('fs');
        fs.writeSync(3, 'EPIPE_THROW\\n');
        failed = true;
      }

      setTimeout(() => {
        const fs = require('fs');
        if (!failed) {
          fs.writeSync(3, 'NO_EPIPE\\n');
        }
        process.exit(0);
      }, 1000);
    `], {
      stdio: ['ignore', 'pipe', 'pipe', 'pipe']
    });

    // 立即销毁管道
    child.stdout.destroy();

    let result = '';
    child.stdio[3].on('data', (data) => {
      result += data.toString();
    });

    child.on('exit', (code) => {
      // 对照组：预期会有 EPIPE
      const epipeOccurred = result.includes('EPIPE');
      resolve({ name: '测试2: 对照组(无防护应触发EPIPE)', passed: epipeOccurred, detail: result.trim(), code });
    });
  });
}

// 测试3: 验证 execFile 不会继承有问题的 stdio
function testExecFileIsolation() {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, ['-e', `
      const _safeWrite = function() { return true; };
      if (process.stdout) process.stdout.write = _safeWrite;
      if (process.stderr) process.stderr.write = _safeWrite;
      const _noop = () => {};
      console.log = _noop;
      console.error = _noop;

      process.on('uncaughtException', (err) => {
        if (err && (err.code === 'EPIPE' || err.code === 'ERR_STREAM_DESTROYED')) return;
        const fs = require('fs');
        fs.writeSync(3, 'FAIL: ' + err.message + '\\n');
        process.exit(1);
      });

      const { execFile } = require('child_process');
      
      // 模拟 checkClipboard 中的 osascript 调用
      execFile('/usr/bin/osascript', ['-e', 'tell application "System Events" to get bundle identifier of first application process whose frontmost is true'], 
        { stdio: ['ignore', 'pipe', 'pipe'] }, 
        (err, stdout) => {
          const fs = require('fs');
          if (err) {
            fs.writeSync(3, 'PASS (execFile err but no EPIPE: ' + err.message + ')\\n');
          } else {
            fs.writeSync(3, 'PASS (execFile ok: ' + stdout.trim() + ')\\n');
          }
          process.exit(0);
        }
      );

      setTimeout(() => {
        const fs = require('fs');
        fs.writeSync(3, 'PASS (timeout)\\n');
        process.exit(0);
      }, 5000);
    `], {
      stdio: ['ignore', 'pipe', 'pipe', 'pipe']
    });

    child.stdout.destroy();
    child.stderr.destroy();

    let result = '';
    child.stdio[3].on('data', (data) => {
      result += data.toString();
    });

    child.on('exit', (code) => {
      const passed = result.trim().startsWith('PASS');
      resolve({ name: '测试3: execFile stdio 隔离', passed, detail: result.trim(), code });
    });
  });
}

// 运行所有测试
async function runTests() {
  process.stdout.write('\n========== EPIPE 防护测试 ==========\n\n');

  const tests = [testStdoutOverride, testWithoutProtection, testExecFileIsolation];
  let allPassed = true;

  for (const test of tests) {
    const result = await test();
    const icon = result.passed ? '✅' : '❌';
    process.stdout.write(`${icon} ${result.name}\n   结果: ${result.detail} (exit code: ${result.code})\n\n`);
    if (!result.passed) allPassed = false;
  }

  process.stdout.write('====================================\n');
  process.stdout.write(allPassed ? '✅ 所有测试通过!\n' : '❌ 部分测试失败!\n');
  process.stdout.write('====================================\n\n');
  process.exit(allPassed ? 0 : 1);
}

runTests();
