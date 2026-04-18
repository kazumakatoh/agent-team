/**
 * Amazon Dashboard - Claude API クライアント
 *
 * Anthropic Claude API（Messages API）の薄いラッパー。
 * 週次（sonnet-4-6）/ 月次（opus-4-6）の改善提案・戦略立案で共通利用する。
 *
 * - APIキー: PropertiesService の CLAUDE_API_KEY
 * - 通信: UrlFetchApp（プロンプト+レスポンス）
 * - リトライ: 5xx / 429 のとき指数バックオフで最大3回
 *
 * モデルID（CLAUDE.md 統括指示書 / 環境記載に準拠）:
 *   - claude-sonnet-4-6  : 週次・通常分析
 *   - claude-opus-4-6    : 月次・戦略立案
 */

const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_API_VERSION = '2023-06-01';

const CLAUDE_MODELS = {
  WEEKLY: 'claude-sonnet-4-6',
  MONTHLY: 'claude-opus-4-6',
};

/**
 * Claude API（Messages）にプロンプトを送信してテキスト応答を返す
 *
 * @param {Object} opts
 *   - model {string}        モデルID（必須）
 *   - system {string}       システムプロンプト
 *   - prompt {string}       ユーザープロンプト（messages の単発投入）
 *   - maxTokens {number}    出力上限（既定 4096）
 *   - temperature {number}  0〜1（既定 0.4）
 * @returns {string} 応答テキスト（content[].text を結合）
 */
function callClaude(opts) {
  const model = opts.model;
  if (!model) throw new Error('callClaude: model 必須');
  const apiKey = getCredential('CLAUDE_API_KEY');

  const payload = {
    model: model,
    max_tokens: opts.maxTokens || 4096,
    temperature: opts.temperature != null ? opts.temperature : 0.4,
    messages: [{ role: 'user', content: opts.prompt || '' }],
  };
  if (opts.system) payload.system = opts.system;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': CLAUDE_API_VERSION,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(CLAUDE_API_URL, options);
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code >= 200 && code < 300) {
      const data = JSON.parse(body);
      const parts = (data.content || []).filter(c => c.type === 'text').map(c => c.text);
      return parts.join('\n').trim();
    }

    // リトライ対象: 429 / 5xx
    if ((code === 429 || code >= 500) && attempt < maxAttempts) {
      const wait = Math.pow(2, attempt) * 1000; // 2s / 4s
      Logger.log('Claude API ' + code + ' → ' + wait + 'ms 待機後リトライ');
      Utilities.sleep(wait);
      continue;
    }
    throw new Error('Claude API エラー HTTP ' + code + ': ' + body.substring(0, 500));
  }
  throw new Error('Claude API リトライ上限超過');
}

/**
 * 接続テスト: 短い応答を要求して疎通確認
 */
function testClaudeApi() {
  const text = callClaude({
    model: CLAUDE_MODELS.WEEKLY,
    prompt: '「OK」とだけ返してください。',
    maxTokens: 32,
    temperature: 0,
  });
  Logger.log('Claude応答: ' + text);
}
