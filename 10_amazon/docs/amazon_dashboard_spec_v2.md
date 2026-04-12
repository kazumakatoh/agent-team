# Amazon収益ダッシュボード 仕様書 v2

*作成日: 2026-04-12*
*ステータス: 社長レビュー待ち*
*依頼者: 加藤一馬（株式会社LEVEL1 代表）*

---

## 1. プロジェクト概要

Amazon EC販売事業の収益・広告・販売データを自動集計し、改善提案まで一気通貫で行うシステム。

### 技術スタック

| 技術 | 用途 |
|---|---|
| Google Apps Script (GAS) | データ取得・自動化・通知 |
| Google スプレッドシート | データ保存・集計・可視化 |
| Claude API (claude-sonnet-4-6) | 改善提案の自動生成 |
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
  ├── Orders API（売上・注文）
  ├── Reports API（Settlement Report / FBA Fee / Return）
  ├── Reports API（GET_SALES_AND_TRAFFIC_REPORT）← PV・訪問数
  └── Inventory API（在庫）

Amazon Advertising API
  ├── キャンペーン実績
  ├── キーワード別実績
  └── 検索用語レポート

       ↓（GAS が毎日自動取得）

Google スプレッドシート
  ├── ① 商品マスター（手入力）
  ├── ② 日次データ（自動・暫定値）
  ├── ③ 経費明細（Settlement Report・確定値）
  ├── ④ 商品別収益（自動集計）
  ├── ⑤ 週次レポート + グラフ
  ├── ⑥ 月次レポート + グラフ
  ├── ⑦ 広告分析
  ├── ⑧ 販売分析
  └── ⑨ 販促費マスター（手入力・レイヤー2用）

       ↓（GAS → Claude API / Gmail / LINE）

通知
  ├── 週次改善提案 → Gmail
  └── 緊急アラート → LINE
```

---

## 3. データソース

詳細: [data_sources.md](./data_sources.md)

## 4. スプレッドシート構成

詳細: [sheet_design.md](./sheet_design.md)

## 5. 利益計算ロジック（2層モデル）

詳細: [profit_model.md](./profit_model.md)

## 6. KPI・通知・運用設計

詳細: [kpi_and_operations.md](./kpi_and_operations.md)

## 7. 実装フェーズ・未解決事項

詳細: [implementation_plan.md](./implementation_plan.md)
