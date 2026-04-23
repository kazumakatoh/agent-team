#!/usr/bin/env node
/**
 * GAS デプロイスクリプト（Service Account 認証）
 *
 * 動作:
 *   - SOURCE_DIR 配下の .gs / appsscript.json を読み込み
 *   - Apps Script API の projects.updateContent で一括更新（= clasp push と同等）
 *
 * 環境変数:
 *   GOOGLE_APPLICATION_CREDENTIALS - SA JSON の path（google-github-actions/auth が自動設定）
 *   SOURCE_DIR                     - 対象 GAS プロジェクトのローカルパス
 *   SCRIPT_ID (任意)               - 明示指定。未指定なら {SOURCE_DIR}/.clasp.json から読む
 *
 * 必要な権限:
 *   - GCP プロジェクトで Apps Script API が有効
 *   - SA がターゲット GAS プロジェクトの Editor として共有されている
 *   - （GAS プロジェクトは同じ GCP プロジェクトにリンクされている必要なし）
 */
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

async function main() {
  const sourceDir = process.env.SOURCE_DIR;
  if (!sourceDir) throw new Error('SOURCE_DIR env var 必須');
  if (!fs.existsSync(sourceDir)) throw new Error('SOURCE_DIR が存在しない: ' + sourceDir);

  // scriptId は env > .clasp.json の順で解決
  let scriptId = process.env.SCRIPT_ID;
  if (!scriptId) {
    const claspPath = path.join(sourceDir, '.clasp.json');
    if (!fs.existsSync(claspPath)) {
      throw new Error('SCRIPT_ID 未指定、.clasp.json も見つからない: ' + claspPath);
    }
    scriptId = JSON.parse(fs.readFileSync(claspPath, 'utf8')).scriptId;
  }
  if (!scriptId) throw new Error('scriptId を解決できなかった');

  // Google Auth（SA JSON は google-github-actions/auth が GOOGLE_APPLICATION_CREDENTIALS に設定済み）
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/script.projects'],
  });
  const authClient = await auth.getClient();
  const script = google.script({ version: 'v1', auth: authClient });

  // .gs / appsscript.json を全て収集
  const entries = fs.readdirSync(sourceDir).sort();
  const files = [];
  for (const entry of entries) {
    if (entry.startsWith('.')) continue;  // .clasp.json 等は除外
    const full = path.join(sourceDir, entry);
    if (!fs.statSync(full).isFile()) continue;

    if (entry === 'appsscript.json') {
      files.push({
        name: 'appsscript',
        type: 'JSON',
        source: fs.readFileSync(full, 'utf8'),
      });
    } else if (entry.endsWith('.gs')) {
      files.push({
        name: entry.replace(/\.gs$/, ''),
        type: 'SERVER_JS',
        source: fs.readFileSync(full, 'utf8'),
      });
    } else if (entry.endsWith('.html')) {
      files.push({
        name: entry.replace(/\.html$/, ''),
        type: 'HTML',
        source: fs.readFileSync(full, 'utf8'),
      });
    }
  }

  // appsscript.json が無ければエラー（Apps Script API の必須ファイル）
  if (!files.some(f => f.name === 'appsscript')) {
    throw new Error('appsscript.json が SOURCE_DIR にない（Apps Script API は必須）');
  }

  console.log('🚀 GAS Deploy');
  console.log('  scriptId:  ' + scriptId);
  console.log('  sourceDir: ' + sourceDir);
  console.log('  files:     ' + files.length + ' 件');
  files.forEach(f => console.log('    - ' + f.name + (f.type === 'JSON' ? '.json' : f.type === 'HTML' ? '.html' : '.gs')));

  const t0 = Date.now();
  const res = await script.projects.updateContent({
    scriptId,
    requestBody: { files },
  });

  const elapsed = Math.round((Date.now() - t0) / 1000);
  const returned = (res.data && res.data.files && res.data.files.length) || 0;
  console.log('✅ デプロイ成功（' + elapsed + '秒）: ' + returned + ' ファイル確定');
}

main().catch(err => {
  console.error('❌ デプロイ失敗:', err.message);
  if (err.errors) console.error(JSON.stringify(err.errors, null, 2));
  if (err.response && err.response.data) console.error(JSON.stringify(err.response.data, null, 2));
  process.exit(1);
});
