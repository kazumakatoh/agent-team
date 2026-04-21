# 実装フェーズ・未解決事項

*v2.7 - Phase 3（Amazon Ads API）実装完了（2026-04-21）*

## 実装フェーズ

### Phase 0: 事前準備（API稼働確認・環境構築）

- [x] Amazon Ads API の Refresh Token 再発行（2026-04-21・+noads エイリアス経由で別アカウント衝突を回避）
- [x] `/v2/profiles` で Profile ID (数値) を取得・確認（JP Seller: 3514200134733929 / トレジャーマーケット 本店）
- [x] GitHubリポジトリ `kazumakatoh/agent-team` で管理（PR #4）
- [x] clasp セットアップ（ローカル ↔ GAS 同期）
- [x] Google スプレッドシート作成・GAS プロジェクト初期化
- [ ] PropertiesService に認証情報を設定（社長作業）
- [ ] 既存Excel資産の共有・移行マッピング設計（社長から共有待ち）

### Phase 1: SP-APIデータ取得 + マスター構築

- [x] LWA認証（Access Token取得）の実装 ← `Auth.gs` / `SpApi.gs`
- [x] Orders API で売上・注文データ取得（CV/点数を分離）← `DailyFetch.gs`
- [x] `GET_SALES_AND_TRAFFIC_REPORT` でPV・訪問数・CVR取得 ← `ReportFetch.gs`
- [x] `GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2` で確定経費取得 ← `SettlementFetch.gs`
- [ ] `GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA` で在庫取得
- [x] M1（商品マスター）作成・カテゴリ設定・CFシートからインポート ← `ProductMaster.gs`
- [x] M2（仕入履歴）作成・CFシート自動連携
- [x] D1（日次データ）への書き込み ← `SheetWriter.gs`
- [x] D2（経費明細）への Settlement Report 取り込み + 月次集計(D2S)
- [ ] 暫定→確定の自動更新ロジック

### Phase 2: 3層ダッシュボード構築

- [x] L1（事業ダッシュボード）: 全体サマリー + カテゴリ別 + 注意商品 ← `Dashboard.gs`
- [x] L2（カテゴリ分析）: カテゴリごとの月次推移 + ASIN別 ← `Dashboard.gs`
- [x] L3（商品分析）: ASIN別一覧・前月比色分け（広告サマリ4指標のみ、詳細はD3）← `L3Product.gs`
- [x] M3（販促費マスター）+ レイヤー2（最終利益）計算 ← `PromoCost.gs`
- [x] 条件付き書式（粗利率・ACOS・TACOS・ROAS のハイライト）← `Formatting.gs`
- [x] 既存Excel過去3年データの一括インポート ← `HistoricalImport.gs`

### Phase 3: 広告データ統合（Ads API稼働後）

- [x] Ads API 認証（Access Token / Profile ID）← `AdsAuth.gs` / `AdsApi.gs`
- [x] ASIN別広告実績の取得（spAdvertisedProduct）← `AdsReport.gs`
- [x] キーワード/ターゲティング別実績の取得（spTargeting）← `AdsReport.gs`
- [x] 検索用語レポートの取得（spSearchTerm）← `AdsReport.gs`
- [x] D3（広告詳細）の実装 - 3シート構成（キャンペーン別 / 検索用語 / ターゲティング）
- [x] D1 日次データへの広告指標統合（広告費 / 広告売上 / IMP / CT）← `updateDailyAdsFromAdvertisedProduct`
- [x] L3 商品分析の広告セクション有効化（既存の広告4指標が自動で埋まるようになる）

**運用開始前チェック**（社長作業）:
1. GASエディタで `addPhase3Triggers()` を▶実行（毎日 11:00 のトリガー追加）
2. `menuFetchAdsReports()` または `testAdsEndToEnd()` を手動実行してログを確認
3. スプシ上の D3_ADS_CAMPAIGN / D3_ADS_SEARCHTERM / D3_ADS_TARGET に行が書き込まれるか確認
4. D1 の広告費 / 広告売上 / IMP / CT 列が更新されるか確認
5. 必要なら `backfillAdsReportsRange(1, 7)` 等で過去日バックフィル

### Phase 4: 改善提案・通知（5カテゴリ + 戦略）

#### Phase 4a: 基本通知
- [x] Claude API 連携（claude-sonnet-4-6）← `ClaudeApi.gs`
- [x] 週次改善提案の生成ロジック（KPI比較・TOP/BOTTOM・5セクション）← `WeeklyAiReport.gs`
- [x] Gmail送信（GmailApp）← `WeeklyAiReport.gs`
- [x] LINE Messaging API 連携 ← `LineAlert.gs`
- [x] 緊急アラートのトリガー条件実装（広告費スパイク / 大型返金 / 重複抑止）← `LineAlert.gs`

#### Phase 4b: 高度分析
- [x] ③競合チェック: Product Pricing API連携（D4 競合価格シート + BuyBox喪失通知）← `Competitor.gs`
- [x] ④セール対策: セールカレンダー + Claude分析（M4 + 6/4/2週前の自動レポート）← `SaleCalendar.gs`
- [x] ⑤アカウント健全性: 返品率 + sellers/v1/account + 在庫切れ比率 → 総合スコア（D5）← `AccountHealth.gs`
- [x] 月次改善提案（claude-opus-4-6）: 戦略立案・逆算分析 ← `MonthlyAiReport.gs`
- [x] 在庫トラッキング + 在庫切れアラート（D6 在庫シート, LINE通知）← `Inventory.gs`
- [x] GASトリガーの全自動化設定 ← `Config.gs#setupDailyTriggers`

### Phase 5: 安定化・本番運用・将来改良

- [ ] エラーハンドリング・リトライ処理
- [ ] GAS実行時間制限への対応（処理分割）
- [ ] 1週間のテスト運用
- [ ] 既存Excelとの数値突合
- [ ] 本番運用開始
- [ ] （将来）FIFO仕入原価追跡
- [ ] （将来）レビュー自動監視
- [ ] （将来）競合価格の定期モニタリング

---

## 未解決事項

### 解決待ち

| # | 事項 | 状態 | 必要なアクション |
|---|---|---|---|
| 1 | ~~Ads API Profiles: 0 問題~~ | ✅ 解決（2026-04-21） | - |
| 2 | 既存Excelの共有 | インポート枠組み完成 / データ共有待ち | `setupHistoricalImportSheet()` 実行後、Excelデータをコピペ → `importHistoricalData()` |

### 実装時に検証が必要

| # | 事項 | リスク | フォールバック |
|---|---|---|---|
| 1 | `GET_SALES_AND_TRAFFIC_REPORT` の権限 | Brand Analytics権限が必要な場合あり | 権限追加申請 or 手動CSV |
| 2 | GAS 6分制限 | レポート取得が6分超える可能性 | 処理分割 |
| 3 | Product Pricing API アクセス | 競合価格取得に追加権限が必要な場合 | 手動確認 + Claude分析 |
| 4 | Account Health API | 権限範囲の確認 | Seller Central画面の目視確認 |

---

## セキュリティ方針

| 項目 | ルール |
|---|---|
| 認証情報の保管 | GAS: `PropertiesService` / ローカル: `.env` |
| GitHubへのコミット | 認証情報は絶対にコミットしない |
| Refresh Token再発行後 | 旧トークン即時無効化・エンジニアに通知 |
| Claude API キー | PropertiesService に保管 |
| LINE トークン | PropertiesService に保管 |

---

*このドキュメントは実装の進捗に合わせて更新する*
