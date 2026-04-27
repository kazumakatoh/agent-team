# データモデル詳細仕様（v0.1）

> Google Sheets各シートの列定義・型・主キー・参照関係・バリデーションルール。
> フェーズ1A成果物。GAS実装はこの仕様に従う。

---

## 1. シート一覧

| シート名 | 役割 | レコード単位 | 主キー |
|---|---|---|---|
| `master_health` | 健診/検査結果（全項目時系列） | 1検査日×1項目 | (exam_date, facility, item_code) |
| `master_vitals` | バイタル日次データ | 1日 | date |
| `item_dictionary` | 項目正規化辞書 | 1項目 | item_code |
| `pending_dictionary` | 新項目候補 | 1項目 | tmp_id |
| `exam_schedule` | 予防検査スケジュール | 1検査種別 | exam_code |
| `exam_history` | 予防検査受診履歴 | 1受診 | (exam_code, exam_date) |
| `_log` | GAS実行ログ | 1イベント | log_id |

---

## 2. `master_health`（健診/検査マスター）

### 2.1 列定義

| # | 列名 | 型 | 必須 | 例 | 備考 |
|---|---|---|---|---|---|
| 1 | exam_date | DATE | ✓ | 2026-04-15 | YYYY-MM-DD |
| 2 | facility | TEXT | ✓ | ○○クリニック | 受診医療機関 |
| 3 | category | ENUM | ✓ | 定期健診 | 定期健診/人間ドック/専門検査 |
| 4 | specialty | ENUM |  | 内科 | 内科/眼科/整形/泌尿器/その他 |
| 5 | item_code | TEXT | ✓ | hba1c | `item_dictionary.item_code` 参照 |
| 6 | item_name_raw | TEXT | ✓ | HbA1c | PDF表記そのまま |
| 7 | value | NUMBER | ✓ | 5.4 | 数値（文字列の場合は note へ） |
| 8 | unit | TEXT |  | % | 単位 |
| 9 | ref_low | NUMBER |  | 4.6 | 基準値下限 |
| 10 | ref_high | NUMBER |  | 6.2 | 基準値上限 |
| 11 | judgment | ENUM |  | A | A/B/C/D/E/その他 |
| 12 | risk_link | TEXT |  | 糖代謝/心疾患 | `item_dictionary` から自動補完 |
| 13 | source_pdf | TEXT | ✓ | 20260415_xx.pdf | 出典PDFファイル名 |
| 14 | extracted_by | ENUM | ✓ | AI | AI / 手動 |
| 15 | confirmed | BOOL | ✓ | TRUE | 社長確認済みフラグ |
| 16 | note | TEXT |  |  | 任意メモ |
| 17 | created_at | DATETIME | ✓ | 2026-04-27T10:00 | 自動 |
| 18 | updated_at | DATETIME | ✓ | 2026-04-27T10:00 | 自動 |

### 2.2 重複チェック
主キー (exam_date, facility, item_code) で UPSERT。同一キー存在時は `updated_at` のみ更新し、値が変わったら警告ログ。

### 2.3 バリデーション
- value は数値必須。文字列値（"陰性"等）は判定文字列として `note` へ
- ref_low ≤ value ≤ ref_high なら `judgment` 自動的にA推定（ただし PDF判定優先）
- 辞書未登録の `item_code` は書き込み拒否、`pending_dictionary` 経由で承認後に書き込み

---

## 3. `master_vitals`（バイタルマスター）

### 3.1 列定義

| # | 列名 | 型 | 必須 | 例 | 備考 |
|---|---|---|---|---|---|
| 1 | date | DATE | ✓ | 2026-04-01 | 主キー |
| 2 | steps | INTEGER |  | 8520 | 歩数 |
| 3 | resting_hr | INTEGER |  | 62 | 安静時心拍 (bpm) |
| 4 | max_hr | INTEGER |  | 142 | 最大心拍 |
| 5 | sleep_hours | NUMBER |  | 6.8 | 睡眠時間 |
| 6 | sleep_score | INTEGER |  | 78 | 睡眠スコア (0-100) |
| 7 | weight_kg | NUMBER |  | 72.3 | 体重 |
| 8 | body_fat_pct | NUMBER |  | 18.5 | 体脂肪率 |
| 9 | active_minutes | INTEGER |  | 35 | アクティブ分 |
| 10 | calories_burned | INTEGER |  | 2450 | 消費カロリー |
| 11 | hrv | NUMBER |  | 45 | 心拍変動 (ms) |
| 12 | spo2 | NUMBER |  | 96.5 | 血中酸素 |
| 13 | source | ENUM | ✓ | google_health_api | google_health_api/manual |
| 14 | created_at | DATETIME | ✓ |  | 自動 |

### 3.2 取得タイミング
- 月次プル：前月分を月初に一括取得
- 必要に応じて日次プルへ移行可能（API制限を考慮）

---

## 4. `item_dictionary`（項目辞書）

### 4.1 列定義

| # | 列名 | 型 | 必須 | 例 | 備考 |
|---|---|---|---|---|---|
| 1 | item_code | TEXT | ✓ | hba1c | 主キー、英小文字+アンダースコア |
| 2 | canonical_name | TEXT | ✓ | HbA1c (NGSP) | 正規名称 |
| 3 | aliases | TEXT |  | HbA1c\|ヘモグロビンA1c | パイプ区切り |
| 4 | category | ENUM | ✓ | 糖代謝 | 脂質/血液/肺機能/糖代謝/肝機能/腎機能/血圧/身体計測/ホルモン/その他 |
| 5 | risk_link | TEXT |  | 糖尿病/心疾患 | スラッシュ区切り |
| 6 | unit_canonical | TEXT |  | % |  |
| 7 | ref_low_default | NUMBER |  | 4.6 | デフォルト基準値下限 |
| 8 | ref_high_default | NUMBER |  | 6.2 | デフォルト基準値上限 |
| 9 | description | TEXT |  | 過去1〜2ヶ月の平均血糖値の指標 |  |
| 10 | status | ENUM | ✓ | confirmed | confirmed/pending |
| 11 | created_at | DATETIME | ✓ |  | 自動 |

---

## 5. `pending_dictionary`（新項目候補）

PDF抽出時に辞書未登録の項目が見つかった場合、ここに自動追加。社長承認後 `item_dictionary` に正式登録。

| # | 列名 | 型 | 必須 | 例 |
|---|---|---|---|---|
| 1 | tmp_id | TEXT | ✓ | tmp_20260427_001 |
| 2 | item_name_raw | TEXT | ✓ | 25-OH ビタミンD |
| 3 | suggested_code | TEXT |  | vitamin_d_25oh |
| 4 | suggested_category | TEXT |  | ビタミン |
| 5 | suggested_unit | TEXT |  | ng/mL |
| 6 | suggested_ref_low | NUMBER |  | 30 |
| 7 | suggested_ref_high | NUMBER |  | 100 |
| 8 | source_pdf | TEXT | ✓ |  |
| 9 | source_value | TEXT |  | 28.5 ng/mL |
| 10 | reviewed | BOOL | ✓ | FALSE |
| 11 | approved | BOOL |  |  |
| 12 | created_at | DATETIME | ✓ |  |

---

## 6. `exam_schedule`（予防検査スケジュール）

### 6.1 列定義

| # | 列名 | 型 | 必須 | 例 |
|---|---|---|---|---|
| 1 | exam_code | TEXT | ✓ | h_pylori |
| 2 | exam_name | TEXT | ✓ | ピロリ菌検査 |
| 3 | target_risk | TEXT | ✓ | 胃がん |
| 4 | risk_priority | ENUM | ✓ | 高 (高/中/低) |
| 5 | frequency_label | TEXT | ✓ | 一生に1回（陰性なら以降不要） |
| 6 | frequency_months | INTEGER |  | NULL（一生に1回） / 12 / 60 |
| 7 | last_done | DATE |  | 2025-06-15 |
| 8 | last_result | TEXT |  | 陰性 |
| 9 | next_due | DATE |  | (frequency_monthsから算出) |
| 10 | status | ENUM | ✓ | 完了/期限内/期限超過/未実施 |
| 11 | facility_recommendation | TEXT |  | 内科 |
| 12 | note | TEXT |  | 陰性確定のため追加検査不要 |
| 13 | updated_at | DATETIME | ✓ |  |

### 6.2 ステータス算出ロジック
```
if last_done is NULL:
    status = "未実施"
elif "陰性確定" in note:
    status = "完了"
elif next_due < today:
    status = "期限超過"
elif next_due - today < 30 days:
    status = "もうすぐ予定日"
else:
    status = "期限内"
```

---

## 7. `exam_history`（予防検査受診履歴）

`exam_schedule` の `last_done` だけでは過去履歴を保持できないため、別シートで全受診を記録。

| # | 列名 | 型 | 必須 | 例 |
|---|---|---|---|---|
| 1 | exam_code | TEXT | ✓ | h_pylori |
| 2 | exam_date | DATE | ✓ | 2025-06-15 |
| 3 | facility | TEXT |  | ○○内科 |
| 4 | result | TEXT |  | 陰性 |
| 5 | result_value | NUMBER |  |  |
| 6 | source_pdf | TEXT |  |  |
| 7 | note | TEXT |  | 抗体法 |
| 8 | created_at | DATETIME | ✓ |  |

→ 受診のたびにレコード追加、`exam_schedule.last_done` は最新日付で更新。

---

## 8. シート間の参照関係

```
item_dictionary
   ▲ item_code (FK)
   │
master_health ──── source_pdf
                       ▲
master_vitals          │
                  pending_dictionary（item_dictionaryに昇格前のバッファ）

exam_schedule ──── exam_code (FK)
   ▲                     ▲
   └──────── exam_history
```

---

## 9. バックアップ・履歴管理

- 全シートに `created_at` `updated_at` を持つ
- 重要シート（`master_health` `exam_history`）は月次でCSVバックアップを `_backup/` に出力
- 削除は論理削除（`deleted_at` 列）ではなく**物理削除**（社長の削除権を尊重）

---

## 10. GAS実装メモ

- Sheets API v4 経由でアクセス
- 一括書き込みは `batchUpdate` を使用
- 重複チェックは事前にキー集合をメモリに読み込んでから行う
- 競合検出のため `updated_at` をチェックして楽観ロック

---

*v0.1 - 2026-04-27 作成*
