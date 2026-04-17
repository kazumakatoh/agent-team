# BigBoss MT4 取引履歴 自動集計システム 実装仕様書

## 概要
BigBoss（MT4口座）でSageMaster経由のコピートレード運用を行っており、その取引結果を自動集計・可視化するシステムを構築する。

- **対象口座**：BigBoss MT4（SageMaster連携口座）
- **想定工数**：2〜3日（09_dev チーム / 外注エンジニア）
- **方針**：MT4の標準HTMLレポートをパース → Google Sheetsへ自動反映 → ダッシュボード化
- **設計原則**：SageMaster本体には一切干渉しない（読み取り専用）

---

## 1. システム構成

```
[MT4端末/VPS]
   ↓ 詳細レポートHTML出力（日次）
[ローカル/クラウドストレージ]（Google Drive または S3）
   ↓ Pythonスクリプト（pandas.read_html）
[集計エンジン]
   ↓ Google Sheets API
[Google Sheets ダッシュボード]
   ↓
[Looker Studio 可視化]
```

---

## 2. 機能要件

### 2.1 データ取得
- MT4「口座履歴」→ 右クリック → 「詳細レポート保存」でHTML出力
- 日次でレポートを所定フォルダに保存（手動 or AutoHotkey/タスクスケジューラで半自動化）
- ファイル命名規則：`bigboss_YYYYMMDD.html`

### 2.2 パース・集計（Python）
- `pandas.read_html()` で約定履歴テーブルを抽出
- 以下のカラムを正規化：
  - 約定日時、銘柄、Buy/Sell、ロット、価格、SL、TP、決済価格、損益（¥/pips）、スワップ、手数料
- 重複排除（チケット番号ベース）
- 累積データとしてGoogle Sheets「raw_trades」シートに追記

### 2.3 KPI算出
集計シート「kpi_daily」「kpi_monthly」を自動生成：

| KPI | 計算式 |
|---|---|
| 日次/月次損益（¥） | 損益合計 + スワップ + 手数料 |
| 勝率（%） | 勝ちトレード数 / 総トレード数 |
| プロフィットファクター | 総利益 / 総損失 |
| 平均RR比 | 平均利益 / 平均損失 |
| 最大ドローダウン | 累積損益曲線から算出 |
| 銘柄別損益 | 銘柄ごとに集計 |
| 時間帯別損益 | JST時間でビン分割（東京/欧州/NY） |

### 2.4 ダッシュボード
- Google Sheets を Looker Studio に接続
- 表示項目：
  - 累積損益カーブ
  - 月次損益バーチャート
  - 勝率・PF・最大DD（KPIカード）
  - 銘柄別損益（横棒グラフ）
  - 時間帯別損益（ヒートマップ）

### 2.5 アラート（任意）
- 日次損失が閾値（例：-5%）を超えたらSlack/メール通知
- 最大DDが過去最大を更新したら通知

---

## 3. 技術スタック

| レイヤ | 採用技術 | 備考 |
|---|---|---|
| 言語 | Python 3.11+ | |
| パース | pandas, beautifulsoup4 | |
| 連携 | gspread, google-auth | Google Sheets API |
| 実行環境 | GitHub Actions（cron）or ローカルcron | 月額課金回避 |
| 可視化 | Google Sheets + Looker Studio | 無料 |
| 通知 | Slack Incoming Webhook | 既存ワークスペース流用 |

---

## 4. 開発タスク（2〜3日想定）

### Day 1：基盤構築
- [ ] GitHubリポジトリ作成（`kazumakatoh/bigboss-mt4-aggregator`）
- [ ] Google Cloud プロジェクト作成、Sheets APIサービスアカウント発行
- [ ] テスト用Sheetsテンプレート作成（raw_trades / kpi_daily / kpi_monthly）
- [ ] サンプルHTMLレポートでパーサー実装

### Day 2：集計ロジック
- [ ] 重複排除・累積データ管理
- [ ] KPI計算関数群
- [ ] Google Sheetsへの書き込み処理
- [ ] エラーハンドリング・ログ

### Day 3：自動化・ダッシュボード
- [ ] GitHub Actions ワークフロー（日次cron）
- [ ] Looker Studio ダッシュボード構築
- [ ] Slackアラート（任意）
- [ ] 運用手順書作成

---

## 5. 運用フロー（社長アクション）

### 日次
1. MT4を開く → 口座履歴 → 詳細レポート保存（30秒）
2. 保存先フォルダ（Google Drive同期）に置く
3. GitHub Actionsが自動実行（朝7時）→ Sheets更新

### 週次
- Looker Studioダッシュボードで成績確認（10分）
- SageMaster配信者の継続/停止判断

### 月次
- 月次レポートPDF出力
- 第8期決算（2026年4月）に向け、伊藤さんへ実績データ共有

---

## 6. セキュリティ

- MT4の**Investor Password**は使用しない（HTMLレポートには口座番号のみ含まれる）
- Google Cloud サービスアカウントキーは GitHub Secrets で管理
- リポジトリは **Private**
- 取引履歴データは社外共有禁止

---

## 7. 将来拡張案

- **Phase 2**：複数SageMaster配信者を並行運用 → 配信者別スコアリング自動化
- **Phase 3**：Myfxbook連携で他社FX口座（XM等）も統合管理
- **Phase 4**：Amazon物販・民泊事業の損益と統合した「全事業ダッシュボード」化

---

## 8. 確認事項（社長レビュー必須）

- [ ] MT4稼働端末はどれか？（Mac / Windows / VPS）
- [ ] Google Workspace のどのアカウントで運用するか
- [ ] アラート通知先（Slack / LINE / メール）
- [ ] 外注エンジニア発注 or 09_dev 内で対応するか
- [ ] 本格運用前に2週間のテスト運用期間を設けるか

---

**作成日**：2026-04-17
**作成者**：09_dev エージェント
**承認**：社長レビュー待ち
