# MoneyForward → Google Drive 自動同期 セットアップ手順

## 概要

MoneyForward クラウド会計のレポート（仕訳一覧・試算表・損益計算書）を  
Google Apps Script で定期取得し、Google Drive に自動保存するスクリプトです。

---

## 前提条件

- Google Workspace / Gmail アカウント
- MoneyForward クラウド会計の API 利用権限（法人プラン）
- MoneyForward API のクライアントID/シークレット

---

## Step 1: MoneyForward API の利用申請

1. [MoneyForward クラウド API](https://biz.moneyforward.com/api/) にアクセス
2. API利用の申請を行い、以下を取得する:
   - **クライアントID**
   - **クライアントシークレット**
3. リダイレクトURIには後で GAS のコールバックURLを設定する（Step 3で取得）

---

## Step 2: Google Apps Script プロジェクトを作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」を作成
3. プロジェクト名を「MoneyForward-GDrive-Sync」に変更
4. `Code.gs` の内容をすべて削除し、本リポジトリの `Code.gs` の内容を貼り付け

---

## Step 3: OAuth2 ライブラリを追加

GAS エディタで以下の操作を行う:

1. 左サイドバーの「ライブラリ」の「＋」をクリック
2. スクリプトID に以下を入力:
   ```
   1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
   ```
3. 「検索」をクリック → **OAuth2** ライブラリが表示される
4. 最新バージョンを選択して「追加」

---

## Step 4: スクリプトプロパティを設定

GAS エディタで:

1. 左サイドバーの ⚙（プロジェクトの設定）をクリック
2. 「スクリプト プロパティ」セクションで以下を追加:

| プロパティ名 | 値 | 必須 |
|---|---|---|
| `MF_CLIENT_ID` | MoneyForward APIのクライアントID | ✅ |
| `MF_CLIENT_SECRET` | MoneyForward APIのクライアントシークレット | ✅ |
| `MF_OFFICE_ID` | MoneyForward の事業所ID | ✅ |
| `GDRIVE_FOLDER_ID` | Google Drive保存先フォルダID | 任意（未設定なら自動作成） |
| `NOTIFICATION_EMAIL` | 通知先メールアドレス | 任意 |

### 事業所ID の確認方法

MoneyForward クラウド会計にログインし、URL の `office_id=` に続く文字列が事業所IDです。  
例: `https://accounting.moneyforward.com/xxxx/journals` の `xxxx` 部分

### Google Drive フォルダID の確認方法

Google Drive でフォルダを開いた際のURLの末尾がフォルダIDです。  
例: `https://drive.google.com/drive/folders/XXXXX` の `XXXXX` 部分

---

## Step 5: OAuth2 認証を実行

1. GAS エディタで関数選択ドロップダウンから `getAuthorizationUrl` を選択
2. 「実行」をクリック
3. 「ログ」（表示 → ログ）を開き、表示されたURLにアクセス
4. MoneyForward のアカウントでログインして認証を許可
5. 「認証成功！」と表示されれば完了

### リダイレクトURIについて

認証URL生成後、MoneyForward API の設定画面で  
GASのコールバックURL をリダイレクトURIとして登録する必要があります。

コールバックURLの形式:
```
https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
```

`{SCRIPT_ID}` は GAS エディタの「プロジェクトの設定」で確認できます。

---

## Step 6: 動作確認

1. `checkSetup` を実行 → すべて「OK」または「認証済み」になっていることを確認
2. `syncReportsManual` を実行 → Google Drive にCSVファイルが保存されることを確認

---

## Step 7: 自動実行トリガーを設定

用途に応じて以下のいずれかを実行:

- **月次実行**（毎月1日 AM9時）: `setupMonthlyTrigger` を実行
- **週次実行**（毎週月曜 AM9時）: `setupWeeklyTrigger` を実行

### トリガーの確認・削除

- GAS エディタ左サイドバーの「トリガー」（時計アイコン）で確認可能
- `removeAllTriggers` を実行すると全トリガーが削除される

---

## Google Drive の保存先構造

```
MoneyForward_Reports/         ← 親フォルダ（GDRIVE_FOLDER_ID で指定 or 自動作成）
  └── 2026-03/               ← 年月サブフォルダ（自動作成）
      ├── 仕訳一覧_2026_03_20260401_090012.csv
      ├── 試算表BS_2026_03_20260401_090012.csv
      └── 損益計算書PL_2026_03_20260401_090012.csv
```

---

## トラブルシューティング

### 「MoneyForward未認証です」エラー
→ `getAuthorizationUrl` を再実行して認証してください

### 「MF API エラー (401)」
→ アクセストークンの期限切れ。`resetAuth` → `getAuthorizationUrl` で再認証

### 「MF API エラー (403)」
→ APIの利用権限がない可能性。MoneyForward のプランとAPI設定を確認

### 「MF API エラー (404)」
→ 事業所IDが正しいか確認。`MF_OFFICE_ID` を再設定

### Google Drive にファイルが作成されない
→ `GDRIVE_FOLDER_ID` が正しいか確認。空にすれば自動作成される

---

## 注意事項

- MoneyForward API には**レート制限**があります。短時間に大量リクエストを送らないでください
- APIの仕様変更により動作しなくなる場合があります。MoneyForward API のドキュメントを定期的に確認してください
- 認証トークンの有効期限は通常30日程度です。期限切れ時は再認証が必要です
