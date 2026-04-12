# 実装フェーズ・未解決事項

## 実装フェーズ

### Phase 0: 事前準備（API稼働確認・環境構築）

- [ ] Amazon Ads API の Refresh Token 再発行
- [ ] `/v2/profiles` で Profile ID (数値) を取得・確認
- [ ] GitHubリポジトリ `kazumakatoh/amazon-dashboard` 作成 (Private)
- [ ] clasp セットアップ（ローカル ↔ GAS 同期）
- [ ] Google スプレッドシート作成・GAS プロジェクト初期化
- [ ] PropertiesService に認証情報を設定
- [ ] 既存Excel資産の共有・移行マッピング設計

### Phase 1: SP-APIデータ取得（Ads API不要で動く範囲）

- [ ] LWA認証（Access Token取得）の実装
- [ ] Orders API で売上・注文データ取得
- [ ] `GET_SALES_AND_TRAFFIC_REPORT` でPV・訪問数取得
- [ ] `GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2` で確定経費取得
- [ ] `GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA` で在庫取得
- [ ] シート①（商品マスター）の作成・ASINマスター移行
- [ ] シート②（日次データ）への書き込み
- [ ] シート③（経費明細）への Settlement Report 取り込み
- [ ] 暫定→確定の自動更新ロジック

### Phase 2: 集計・可視化

- [ ] シート④（商品別収益）の自動集計（レイヤー1）
- [ ] シート⑤（週次レポート）+ グラフ生成
- [ ] シート⑥（月次レポート）+ グラフ生成
- [ ] シート⑧（販売分析）
- [ ] シート⑨（販促費マスター）+ レイヤー2計算
- [ ] 条件付き書式（粗利率・ACOS・TACOS のハイライト）
- [ ] 既存Excel過去データの一括インポート

### Phase 3: 広告データ統合（Ads API稼働後）

- [ ] Ads API 認証（Access Token / Profile ID）
- [ ] キャンペーン実績の取得
- [ ] キーワード別実績の取得
- [ ] 検索用語レポートの取得
- [ ] シート⑦（広告分析）の実装
- [ ] 日次データへの広告指標統合（IMP・CT・広告売上等）

### Phase 4: 改善提案・通知

- [ ] Claude API 連携（claude-sonnet-4-6）
- [ ] 週次改善提案の生成ロジック
- [ ] Gmail送信（GmailApp）
- [ ] 月次改善提案の生成（claude-opus-4-6）
- [ ] LINE Messaging API 連携
- [ ] 緊急アラートのトリガー条件実装
- [ ] GASトリガーの全自動化設定

### Phase 5: 安定化・本番運用

- [ ] エラーハンドリング・リトライ処理
- [ ] GAS実行時間制限への対応（処理分割）
- [ ] 1週間のテスト運用
- [ ] 既存Excelとの数値突合
- [ ] 本番運用開始

---

## 未解決事項

### 解決待ち

| # | 事項 | 状態 | 必要なアクション |
|---|---|---|---|
| 1 | Ads API Profiles: 0 問題 | Refresh Token再発行で解決予定 | 社長がadvertising.amazon.comでJPプロファイル確認 → Token再発行 |
| 2 | 既存Excelの共有 | 社長から後日共有 | 共有後に移行マッピング設計 |

### 実装時に検証が必要

| # | 事項 | リスク | フォールバック |
|---|---|---|---|
| 1 | `GET_SALES_AND_TRAFFIC_REPORT` の権限 | Brand Analytics権限が必要な場合あり | 権限追加申請 or 週1手動CSV（最終手段） |
| 2 | GAS 6分制限 | レポート取得が6分超える可能性 | 処理分割（createReport / download を別トリガー） |
| 3 | Settlement Report の全カテゴリ網羅 | 未知のカテゴリが存在する可能性 | 「その他」カテゴリで吸収 |

---

## セキュリティ方針

| 項目 | ルール |
|---|---|
| 認証情報の保管 | GAS: `PropertiesService.getScriptProperties()` / ローカル: `.env` |
| GitHubへのコミット | 認証情報は絶対にコミットしない。`.gitignore` で除外 |
| Refresh Token再発行後 | 旧トークンは即時無効化。エンジニアに旧トークン無効化を通知 |
| Claude API キー | PropertiesService に保管 |
| LINE トークン | PropertiesService に保管 |

---

*このドキュメントは実装の進捗に合わせて更新する*
