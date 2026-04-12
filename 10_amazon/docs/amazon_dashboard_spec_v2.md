# Amazon収益ダッシュボード 仕様書 v2.1

*作成日: 2026-04-12*
*最終更新: 2026-04-12（社長フィードバック反映）*
*ステータス: 社長レビュー済み*
*依頼者: 加藤一馬（株式会社LEVEL1 代表）*

---

## 1. プロジェクト概要

Amazon EC販売事業の収益・広告・販売データを自動集計し、改善提案まで一気通貫で行うシステム。

### ゴール

社長が把握したいこと:
- **Amazon事業全体**がうまくいっているか（俯瞰）
- **カテゴリ単位**でどうか（中解像度）
- **個別商品(ASIN)**の収益・広告効果（高解像度）

→ 3層ダッシュボードで「俯瞰→深堀り」を実現

### 技術スタック

| 技術 | 用途 |
|---|---|
| Google Apps Script (GAS) | データ取得・自動化・通知 |
| Google スプレッドシート | データ保存・集計・可視化 |
| Claude API (claude-sonnet-4-6 / claude-opus-4-6) | 改善提案の自動生成 |
| GitHub + clasp | コード管理・デプロイ |
| LINE Messaging API | 緊急アラート通知 |

### 方針

- Make は使わない（GAS のみで完結）
- Python 不要（GAS の JavaScript で完結）
- 外部エンジニアは介さない（Claude Code で開発）
- 認証情報はコードに直書きしない（PropertiesService で管理）

---

## 2. システム全体構成

```
Amazon SP-API
  ├── Orders API（売上・注文: CV/点数分離）
  ├── Reports API（Settlement Report）
  ├── Reports API（GET_SALES_AND_TRAFFIC_REPORT: PV/セッション/CVR）
  ├── Inventory API（在庫）
  ├── Catalog Items API（VINE/A+/商品情報チェック）
  ├── Product Pricing API（競合価格）
  └── Account Health API（アカウント健全性）

Amazon Advertising API
  ├── キャンペーン実績（ASIN別・M19連携）
  ├── キーワード別実績
  └── 検索用語レポート

       ↓（GAS が毎日自動取得）

Google スプレッドシート
  【見るシート（3枚）】
  ├── L1 事業ダッシュボード（全体+カテゴリ+注意商品）
  ├── L2 カテゴリ分析（ドロップダウンで選択）
  └── L3 商品分析（ドロップダウンで選択・日次〜年次）

  【データシート（裏側）】
  ├── D1 日次データ（全商品x全日付）
  ├── D2 経費明細（Settlement Report確定値）
  └── D3 広告詳細（キャンペーン/キーワード/検索用語）

  【マスター（手入力）】
  ├── M1 商品マスター（ASIN/カテゴリ/仕入単価）
  ├── M2 仕入履歴（単価変遷ログ）
  └── M3 販促費マスター（Amence/M19/人件費）

       ↓（GAS → Claude API / Gmail / LINE）

改善提案（5カテゴリ + 戦略立案）
  ├── ① 商品登録チェック（VINE/A+/広告設定）
  ├── ② 販売改善（KPI比較・レビュー監視）
  ├── ③ 競合チェック（価格・広告強度・施策）
  ├── ④ セール対策（セール情報・対策・施策）
  ├── ⑤ アカウント健全性（警告・在庫切れ・エラー）
  └── + AI戦略立案（利益最大化の逆算分析）

通知
  ├── 週次改善提案 → Gmail (Sonnet 4.6)
  ├── 月次改善提案 + 戦略 → Gmail (Opus 4.6)
  └── 緊急アラート → LINE
```

---

## 3. 詳細設計ドキュメント

| ドキュメント | 内容 |
|---|---|
| [data_sources.md](./data_sources.md) | API認証・取得レポート・暫定/確定ルール |
| [sheet_design.md](./sheet_design.md) | 3層ダッシュボード + データ/マスターシート設計 |
| [profit_model.md](./profit_model.md) | 2層利益モデル・仕入履歴管理・3軸表示 |
| [kpi_and_operations.md](./kpi_and_operations.md) | KPI全指標・通知設計・改善提案5カテゴリ・運用 |
| [implementation_plan.md](./implementation_plan.md) | Phase 0〜5・未解決事項・セキュリティ方針 |
