#!/usr/bin/env node
/**
 * GAS デプロイスクリプト（Service Account + Domain-Wide Delegation）
 *
 * Apps Script API は呼び出しユーザー単位の有効化フラグを持っており、
 * SA は自己 ON できないため、DWD で Workspace ユーザーに成りすまして
 * API を叩く。
 *
 * 認証フロー:
 *   1. WIF で SA に認証（GOOGLE_APPLICATION_CREDENTIALS 経由）
 *   2. IAM Credentials API (signJwt) で user-scope JWT を作成
 *   3. JWT を OAuth2 トークンエンドポイントで access_token に交換
 *   4. user-scope access_token で Apps Script API の updateContent を呼ぶ
 *
 * 必要な GCP 設定:
 *   - SA に roles/iam.serviceAccountTokenCreator（自分自身向け）
 *   - Workspace 管理コンソールで DWD 設定
 *     Client ID = SA の一意ID、Scope = https://www.googleapis.com/auth/script.projects
 *   - GAS プロジェクトが IMPERSONATE_USER で編集可能（所有者 or 編集者）
 *
 * 環境変数:
 *   GOOGLE_APPLICATION_CREDENTIALS - google-github-actions/auth が自動設定
 *   SOURCE_DIR                     - 対象 GAS プロジェクトのローカルパス
 *   SA_EMAIL                       - Service Account email
 *   IMPERSONATE_USER               - 成りすまし対象の Workspace ユーザー email
 *   SCRIPT_ID (任意)               - 未指定なら .clasp.json から解決
 */
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

/**
 * DWD を使って user-scope の OAuth access_token を取得
 */
async function getImpersonatedToken(saEmail, impersonateUser, scopes) {
  // SA として GCP 認証（IAM Credentials API 呼び出し用）
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/iam'],
  });
  const authClient = await auth.getClient();
  const iamCredentials = google.iamcredentials({ version: 'v1', auth: authClient });

  // 成りすまし JWT を SA に署名させる
  const now = Math.floor(Date.now() / 1000);
  const jwtPayload = {
    iss: saEmail,
    sub: impersonateUser,
    scope: scopes.join(' '),
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  };

  const signRes = await iamCredentials.projects.serviceAccounts.signJwt({
    name: `projects/-/serviceAccounts/${saEmail}`,
    requestBody: { payload: JSON.stringify(jwtPayload) },
  });

  // JWT を access_token に交換
  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: signRes.data.signedJwt,
    }),
  });
  const tokenJson = await tokenRes.json();
  if (!tokenJson.access_token) {
    throw new Error('Token exchange failed: ' + JSON.stringify(tokenJson));
  }
  return tokenJson.access_token;
}

async function main() {
  const sourceDir = process.env.SOURCE_DIR;
  const saEmail = process.env.SA_EMAIL;
  const impersonateUser = process.env.IMPERSONATE_USER;

  if (!sourceDir) throw new Error('SOURCE_DIR env var 必須');
  if (!saEmail) throw new Error('SA_EMAIL env var 必須');
  if (!impersonateUser) throw new Error('IMPERSONATE_USER env var 必須');
  if (!fs.existsSync(sourceDir)) throw new Error('SOURCE_DIR が存在しない: ' + sourceDir);

  // scriptId
  let scriptId = process.env.SCRIPT_ID;
  if (!scriptId) {
    const claspPath = path.join(sourceDir, '.clasp.json');
    if (!fs.existsSync(claspPath)) {
      throw new Error('SCRIPT_ID 未指定 / .clasp.json も無い: ' + claspPath);
    }
    scriptId = JSON.parse(fs.readFileSync(claspPath, 'utf8')).scriptId;
  }
  if (!scriptId) throw new Error('scriptId を解決できなかった');

  console.log('🚀 GAS Deploy (DWD)');
  console.log('  scriptId:        ' + scriptId);
  console.log('  sourceDir:       ' + sourceDir);
  console.log('  SA:              ' + saEmail);
  console.log('  Impersonate as:  ' + impersonateUser);

  // .gs / appsscript.json / .html を全て収集
  const entries = fs.readdirSync(sourceDir).sort();
  const files = [];
  for (const entry of entries) {
    if (entry.startsWith('.')) continue;
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

  if (!files.some(f => f.name === 'appsscript')) {
    throw new Error('appsscript.json が SOURCE_DIR にない（Apps Script API では必須）');
  }

  console.log('  files:           ' + files.length + ' 件');
  files.forEach(f => {
    const ext = f.type === 'JSON' ? '.json' : f.type === 'HTML' ? '.html' : '.gs';
    console.log('    - ' + f.name + ext);
  });

  // DWD でユーザー成りすまし access_token を取得
  console.log('  🔑 user-token を取得中...');
  const accessToken = await getImpersonatedToken(
    saEmail,
    impersonateUser,
    ['https://www.googleapis.com/auth/script.projects']
  );
  console.log('  ✓ user-token 取得完了');

  // ユーザートークンで Apps Script API 呼び出し
  const userAuth = new google.auth.OAuth2();
  userAuth.setCredentials({ access_token: accessToken });
  const script = google.script({ version: 'v1', auth: userAuth });

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
