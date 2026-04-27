/* ======================================================================
 * Offer审批助手 · AI Provider 抽象层
 * 默认关闭，调用前强制脱敏
 * ====================================================================== */

const AI_PROVIDERS = {
  disabled: {
    name: '禁用（离线模式）',
    call: async () => { throw new Error('AI 功能未启用，请在设置中开启'); },
  },
  // 预留：未来可扩展
  // openai: { name: 'OpenAI', call: async (prompt) => { ... } },
  // deepseek: { name: 'DeepSeek', call: async (prompt) => { ... } },
};

let _currentProvider = 'disabled';
let _aiApiKey = '';

/**
 * 获取当前 AI 状态
 */
function getAIStatus() {
  return {
    enabled: _currentProvider !== 'disabled',
    provider: _currentProvider,
    providerName: AI_PROVIDERS[_currentProvider]?.name || '未知',
  };
}

/**
 * 设置 AI Provider
 */
function setAIProvider(providerId, apiKey) {
  if (AI_PROVIDERS[providerId]) {
    _currentProvider = providerId;
    _aiApiKey = apiKey || '';
    lsSet('ai_provider', providerId);
    // 注意：API Key 不持久化，每次启动需重新输入
  }
}

/**
 * 调用 AI（自动走脱敏管道）
 */
async function callAIProvider(prompt) {
  const status = getAIStatus();
  if (!status.enabled) {
    throw new Error('AI 功能未启用。当前为完全离线模式。');
  }
  const provider = AI_PROVIDERS[_currentProvider];
  return await provider.call(prompt, _aiApiKey);
}
