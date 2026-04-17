# Amazon Dashboard - セットアップ手順

## 前提条件

- Google アカウント（スプレッドシート使用可）
- Node.js がインストール済み（clasp用）
- Amazon SP-API の認証情報
- Amazon Ads API の認証情報（サポート回答後）

---

## Step 1: Google スプレッドシートを作成

1. Google Drive で新しいスプレッドシートを作成
2. 名前: 「Amazon Dashboard」
3. 以下のシートを作成（タブ名を正確に）:
   - `事業ダッシュボード`
   - `カテゴリ分析`
   - `商品分析`
   - `日次データ`
   - `経費明細`
   - `広告詳細`
   - `商品マスター`
   - `月次仕入単価`
   - `販促費マスター`
4. スプレッドシートのURLからIDをコピー
   - URL: `https://docs.google.com/spreadsheets/d/XXXXXXX/edit`
   - `XXXXXXX` の部分がID

---

## Step 2: GASプロジェクトを作成

1. スプレッドシートを開く
2. メニュー → `拡張機能` → `Apps Script`
3. プロジェクト名を「Amazon Dashboard」に変更
4. 表示されたスクリプトIDをメモ（URL内の `/projects/XXXXXXX/edit` の部分）

---

## Step 3: clasp セットアップ

```bash
# clasp をインストール（初回のみ）
npm install -g @google/clasp

# Googleアカウントでログイン（初回のみ）
clasp login

# プロジェクトを接続
cd 10_amazon/gas
cp .clasp.json.example .clasp.json
# .clasp.json の scriptId を Step 2 のスクリプトIDに書き換え

# コードをGASに反映
clasp push
```

---

## Step 4: 認証情報を設定

1. GASエディタで `Config.gs` の `setupCredentials()` を開く
2. 各認証情報を入力:
   - SP-API: SELLER_ID, CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN
   - Ads API: CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN（新規取得済み）
   - MAIN_SHEET_ID: Step 1 のスプレッドシートID
   - CF_SHEET_ID: キャッシュフロー管理シートのID
   - CLAUDE_API_KEY: Anthropic API Key
   - GMAIL_TO: 社長のメールアドレス
3. `setupCredentials()` を実行
4. **実行後、関数内の値を空文字に戻す**（セキュリティ）
5. `checkCredentials()` を実行して設定状態を確認

---

## Step 5: 接続テスト

GASエディタで以下の関数を順に実行:

1. `testSpApiAuth()` - SP-API認証テスト
2. `testGetOrders()` - 注文データ取得テスト
3. `testAdsApiAuth()` - Ads API認証テスト（サポート回答後）

---

## clasp での開発フロー

### 自動デプロイ（推奨）

`main` または `claude/**` ブランチに `10_amazon/gas/**` への変更が push されると、
GitHub Actions が自動で `clasp push --force` を実行する。
ワークフロー: `.github/workflows/deploy-gas-amazon.yml`

→ **ローカルに clasp 環境を作らなくても、push するだけで GAS に反映される**

#### トークン切れ時の対応

`CLASPRC_JSON` の RAPT が期限切れすると `invalid_grant` エラーでデプロイが
失敗する。その時は以下の手順で更新：

1. ローカル PC で `clasp login` 実行 → 新しいトークン発行
2. `~/.clasprc.json` の内容をクリップボードにコピー
3. GitHub Secrets の `CLASPRC_JSON` を更新
4. トリガー用に `10_amazon/gas/` 内のファイルを変更して push

> ⚠️ トークンの中身（refresh_token / client_secret）は絶対に外部に見せない。
> 誤って露出した場合は https://myaccount.google.com/permissions で即座に無効化し再発行する。
> 更新後は `10_amazon/gas/` 内のファイルを変更して push すると自動デプロイが起動。

### 手動デプロイ（ローカル作業時）

```bash
# ローカルで編集 → GASに反映
clasp push

# GASから最新を取得
clasp pull

# GASエディタをブラウザで開く
clasp open
```
