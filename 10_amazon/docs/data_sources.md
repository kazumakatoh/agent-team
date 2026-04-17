# データソース定義

## SP-API（売上・経費・在庫・トラフィック）

### 認証情報

| 項目 | 値 | 保管場所 |
|---|---|---|
| SELLER_ID | 設定済み | PropertiesService |
| MARKETPLACE_ID (JP) | A1VC38T7YXB528 | PropertiesService |
| LWA Client ID | 設定済み | PropertiesService |
| LWA Client Secret | そのまま維持 | PropertiesService |
| LWA Refresh Token | 設定済み | PropertiesService |

### 取得レポート一覧

| レポート種別 | API / レポートID | 取得内容 | 更新頻度 |
|---|---|---|---|
| 売上・注文 | Orders API | 注文数・売上金額・ASIN別売上 | 日次 |
| 収益明細（確定） | `GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2` | 売上・全手数料・返金の確定値 | 14日サイクル |
| FBA手数料（速報） | `GET_FBA_ESTIMATED_FBA_FEES_TXT_DATA` | 配送料・保管料の予測値 | 日次 |
| 返品情報 | `GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA` | 返品数・返金額 | 日次 |
| 在庫 | `GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA` | 現在庫数 | 日次 |
| トラフィック | `GET_SALES_AND_TRAFFIC_REPORT` | PV・訪問数・CVR・Buy Box率 | 日次 |

### Settlement Report の経費カテゴリ（すべて集計対象）

トランザクション画面に表示されるすべてのカテゴリを自動集計する:

- 注文（売上）
- 返金
- サービス料（販売手数料・FBA手数料等）
- 税金
- Amazon Easy Ship 料金
- FBA在庫の払戻
- 保証請求
- 調整金
- パススルー
- 補償
- プロモーション（クーポン費用等）
- Amazon Shipping
- Inventory Reimbursement
- その他

### 暫定値→確定値の運用ルール

- **日次データ（シート②）**: Orders API + FBA Fee Preview の暫定値で速報表示
- **経費明細（シート③）**: Settlement Report 到着時に確定値で上書き
- **月次レポート**: 該当月の Settlement Report が全て到着後にファイナライズ
- 暫定行には「※暫定」マークを付与、確定後に「確定」マークに変更

---

## Amazon Advertising API（広告データ）

### 認証情報

| 項目 | 値 | 保管場所 |
|---|---|---|
| ADS_CLIENT_ID | 設定済み | PropertiesService |
| ADS_CLIENT_SECRET | そのまま維持 | PropertiesService |
| ADS_REFRESH_TOKEN | **再発行予定** | PropertiesService |
| ADS_PROFILE_ID | `/v2/profiles` から数値IDを再取得 | PropertiesService |

### エンドポイント

- リージョン: **Far East**
- URL: `https://advertising-api-fe.amazon.com`

### 取得データ

| データ種別 | 取得内容 | 更新頻度 |
|---|---|---|
| キャンペーン実績 | 広告費・IMP・CT・広告売上・ACOS・ROAS | 日次 |
| キーワード別実績 | キーワードごとのACOS/ROAS/CTR | 日次 |
| 検索用語レポート | 実際の検索キーワードと成果 | 日次 |

### 現在のブロッカー

- `/v2/profiles` が **0件**を返す問題
- 原因: OAuth認可時のアカウント紐付け不足（マーケットプレイス関連チェック未設定）
- 対応: Refresh Token 再発行で解決予定

---

## トラフィックデータの取得方法

`GET_SALES_AND_TRAFFIC_REPORT` (SP-API) を使用。

| 指標 | 日本語 | 取得可否 |
|---|---|---|
| sessions | セッション数（訪問数） | OK |
| pageViews | ページビュー（PV） | OK |
| unitSessionPercentage | ユニットセッション率（CVR） | OK |
| buyBoxPercentage | Buy Box 獲得率 | OK |
| unitsOrdered | 注文点数 | OK |
| orderedProductSales | 売上 | OK |

- ASIN別・日次の粒度で取得可能
- 最大2年分の過去データを取得可能
- SP-API認証情報で取得可能（追加のBrand Analytics権限が必要な場合は実装時に検証）
