# セットアップ手順書（社長作業ガイド）

> フェーズ1A〜1Cで社長に手元で行っていただく作業の手順。
> リポジトリ側のドキュメント・テンプレ・seedデータは作成済（v0.3）。

---

## 全体の流れ

```
[1A] Drive構造作成 + Sheets作成（seedsインポート）
       │
       ▼
[1A後半] exam_protocols.md 過去履歴入力
       │
       ▼
[1B] GAS実装（私が並行作業、社長は動作確認のみ）
       │
       ▼
[1C] Google Cloud Console + Fitbit過去データ救済
     ★2026年7月15日が期限★
       │
       ▼
[2] 過去PDFの順次流し込み（1〜3年分）
       │
       ▼
[3] 秘書朝レポート統合 → 通常運用
```

---

## フェーズ1A：Drive構造とSheets作成

### A-1. Drive 構造を作る

会社Drive `12_加藤家/` 配下に以下を作成。**共有設定は「制限付き」、社長個人アカウントのみ閲覧可**に設定すること。

```
12_加藤家/
└ 01_健康/                       ← 新規作成
   ├ 01_inbox/                   ← 新規作成（PDFを置く場所）
   ├ 02_processed/               ← 新規作成
   ├ analysis/                   ← 新規作成
   │  └ _log/                    ← 新規作成
   └ _backup/                    ← 新規作成
```

### A-2. Google Sheets を作成

`12_加藤家/01_健康/` 直下に以下のSheetsを作成し、**それぞれに `seeds/` のCSVをインポート**：

| Sheets名 | seedファイル | 備考 |
|---|---|---|
| `master_health` | （空） | ヘッダ行のみ：`data_model.md` §2 参照 |
| `master_vitals` | （空） | 同 §3 |
| `item_dictionary` | `seeds/item_dictionary_seed.csv` | **45項目をインポート** |
| `pending_dictionary` | （空） | 同 §5 |
| `exam_schedule` | `seeds/exam_schedule_seed.csv` | **22検査をインポート** |
| `exam_history` | （空） | 同 §7 |

**インポート方法**：
1. Sheetsを開く → ファイル → インポート → アップロード
2. `seeds/xxx_seed.csv` を選択
3. 「現在のシートを置換する」を選択

### A-3. exam_protocols.md に過去履歴を入力

`99_personal_health/exam_protocols.md` の **§3「過去受診履歴 入力欄」** に、思い出せる範囲で過去の受診を記入。**特に重要：**

- ✅ **ピロリ菌検査**（陰性確定なら胃がんリスク大幅減）
- ✅ **肝炎ウイルス検査**（陰性確定なら肝がんリスク大幅減）
- 過去3年の定期健診（3機関分）
- テストステロン検査の過去結果
- その他覚えている検査

→ 入力後、私が `exam_history` と `exam_schedule.last_done` に転記します。

---

## フェーズ1B：GAS実装

ここは私（Claude）が実装します。社長の作業は最小限：

- [ ] GASプロジェクトをDriveに紐付け（手順は実装時に案内）
- [ ] OAuth認可画面で許可（初回1回のみ）
- [ ] 1件のテストPDFで動作確認

---

## フェーズ1C：Google Cloud Console + Fitbit過去データ救済 ★最優先★

**期限：2026年7月15日**（Fitbit旧データ救済の最終期限）

### C-1. Google Cloud Consoleでプロジェクト作成

1. `kazuma.katoh.0406@gmail.com` でログイン
2. https://console.cloud.google.com/ にアクセス
3. 新規プロジェクト作成：プロジェクト名「kazuma-health-data」
4. 以下のAPIを有効化：
   - **Google Health API**（旧Fitbit Web APIの後継）
   - Google Drive API
   - Google Sheets API
   - Gmail API
5. OAuth同意画面を「外部」で設定、テストユーザーに `kazuma.katoh.0406@gmail.com` を追加
6. OAuth 2.0クライアントID（種類：ウェブアプリケーション）を作成

### C-2. Fitbit過去データのバックアップ

Google Health API での取得とは別に、**最終救済として旧Fitbit Web APIから生データをCSVで取得**しておく。

1. https://www.fitbit.com/settings/data/export にアクセス
2. 「データのアーカイブをダウンロード」を選択
3. ダウンロードしたZIPを `12_加藤家/01_健康/_backup/fitbit_archive_2026-XX.zip` に保存

**期限厳守：2026年7月15日**

### C-3. Google Health APIの認証情報設定

GASのスクリプトプロパティに以下を設定（私が実装時にサポート）：
- `CLIENT_ID`
- `CLIENT_SECRET`
- `REFRESH_TOKEN`

---

## フェーズ2：過去PDFの流し込み

1. 過去1〜3年分の健診/検査PDFを `01_inbox/` に順次保存
2. GASが自動抽出 → `master_health` 追記
3. 新項目候補が出たらGmailで通知 → 社長承認
4. 抽出ミスは `master_health` で手動修正

---

## フェーズ3：秘書朝レポート統合・通常運用

- 月次バイタルレビューを秘書エージェントに連携
- 四半期レビューの初回生成
- 必要に応じてテンプレート微調整

---

## トラブル時の連絡

各フェーズで詰まったら、いつでも私（Claude）に「○○がうまくいかない」とお伝えください。実装側の修正で対応します。

---

*v0.1 - 2026-04-27 作成*
