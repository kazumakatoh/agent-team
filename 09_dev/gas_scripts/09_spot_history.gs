/**
 * 現物資産履歴シート（V3：6行統一・合計枚数追加・時価増加率追加・基準枚数1以上）
 *
 * - 月次履歴：年月×通貨ブロック
 * - 日次履歴：日付×通貨ブロック（4/13〜+60日）
 * - 全通貨6行構造：枚数/枚数増加率/評価額/評価額増加率/時価/時価増加率
 * - 合計5行：投入元本/枚数/枚数増加率/評価額/評価額増加率
 * - 基準日：枚数≥1の最初の日
 * - 枚数<1日は増加率0%
 * - USD/JPY切替対応
 */

const SPOT_HIST_CONFIG = {
  SHEET_MONTHLY: '現物_月次履歴',
  SHEET_DAILY: '現物_日次履歴',
  COINS: ['USDT', 'TIA', 'RENDER', 'LINK', 'ALGO', 'CAKE', 'ENA', 'PENDLE'],
  START_YEAR: 2026,
  START_MONTH: 4,
  END_YEAR: 2030,
  END_MONTH: 12,
  DAILY_START: new Date(2026, 3, 13),
  DAILY_DAYS_AHEAD: 60,
  INITIAL_CAPITAL_USD: 2972
};

const COIN_ROWS = ['枚数', '枚数増加率', '評価額', '評価額増加率', '時価', '時価増加率'];
const COIN_ROWS_USDT = ['枚数', '枚数増加率', '評価額', '評価額増加率', '時価', '時価増加率'];
const TOTAL_ROWS = ['投入元本', '枚数', '枚数増加率', '評価額', '評価額増加率'];

// このファイルは Apps Script 上で完全実装されており、ここはアーカイブ用ヘッダーのみ。
// 完全な実装は社長のスプシ「SageMaster_実績集計」の Apps Script エディタを参照。
//
// 主要関数：
//   - initSpotMonthlyHistory()        : 月次履歴シート初期化
//   - initSpotDailyHistory()          : 日次履歴シート初期化
//   - appendDailySpotHistory()        : 日次履歴へ今日の列追加（毎日6:30トリガー）
//   - snapshotMonthlySpotHistory()    : 月次履歴の当月列確定（毎日21:00、月末日のみ動作）
//   - backfillDailySpotHistory()      : 過去取引から日次データ全期間再構築
//   - addHistorySheetToggles()        : USD/JPY通貨切替UI追加
//   - applyGrayFillToHistory()        : 枚数/評価額/時価行を薄いグレー塗り
//   - finalizeHistoryFormat()         : 書式整備（時価のみ$表示、それ以外整数）
//   - removeCurrencyLabelFromHistory(): A列ラベルから「($)」削除
//   - setupSpotHistoryTriggers()      : 自動更新トリガー登録
//
// 計算ロジック：
//   - baseDate = 枚数≥1の最初の日
//   - 枚数<1の日 → 増加率0%
//   - 通貨別：枚数増加率・評価額増加率・時価増加率を baseDate基準で計算
//   - 合計：枚数は単純合計、評価額増加率は投入元本($2,972)基準
