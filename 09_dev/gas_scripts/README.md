# GAS Scripts - SageMaster統合運用ダッシュボード

## ファイル構成

| ファイル | 役割 |
|---|---|
| `01_mt4_html_parser.gs` | BigBoss MT4 HTML詳細レポートをパース、FX実績を取得 |
| `02_mexc_api_integration.gs` | MEXC取引所のAPI経由で現物保有資産・取引履歴を取得 |
| `03_weekly_review_generator.gs` | 週次レビューを自動生成しスプシ＋Gmail通知 |

## セットアップ手順

### 1. スプシにコードを貼り付け
1. 対象スプシを開く
2. メニュー「拡張機能」→「Apps Script」
3. 各 `.gs` ファイルの内容をコピペ

### 2. APIキー設定（MEXC）
1. MEXC取引所で **読み取り専用** のAPIキーを発行
2. Apps Script の「プロジェクトの設定」→「スクリプトプロパティ」
3. 以下2項目を追加：
   - `MEXC_API_KEY`：APIキー
   - `MEXC_SECRET_KEY`：シークレットキー

### 3. トリガー設定（初回のみ手動実行）
```javascript
// Apps Scriptエディタで以下を1回実行
setupTriggers();
```
- 月曜 7:00 JST：週次レビュー自動生成
- 毎日 6:00 JST：MT4/MEXC データ集計

### 4. Google Drive フォルダ準備
- `SageMaster/FX/MT4_Reports/` フォルダを作成
- 社長が毎週月曜朝にMT4から `DetailedStatement.htm` をここへ保存

### 5. 動作テスト
```javascript
// 手動実行でテスト
runMT4Aggregation();
runMEXCAggregation();
generateWeeklyReview();
```

## 権限

初回実行時にGoogleから以下の権限許可を求められる：
- スプレッドシートへのアクセス
- Google Driveへのアクセス
- 外部URL（MEXC API）へのアクセス
- Gmail送信

すべて社長個人のGoogleアカウント配下で動作。

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| MT4 HTML が読み込まれない | ファイル名が `DetailedStatement.htm` になっているか確認 |
| MEXC API 401エラー | APIキー/シークレットキーを再確認、IPホワイトリスト設定 |
| 週次レビューが届かない | トリガー設定確認、`setupTriggers()` 再実行 |
| 数値がずれる | タイムゾーン設定（Asia/Tokyo）を確認 |

## 実装ステータス

- [x] MT4 HTMLパーサー（骨子）
- [x] MEXC API連携（骨子）
- [x] 週次レビュー生成ロジック（骨子）
- [ ] スプシ書き込み処理の詳細実装（Day 2）
- [ ] 将来シミュレーション（モンテカルロ）の実装（Day 3）
- [ ] Looker Studio連携（Day 3）
- [ ] 動作テスト・チューニング（Day 3）
