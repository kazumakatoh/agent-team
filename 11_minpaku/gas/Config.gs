/**
 * 民泊自動化システム - 設定ファイル
 * 和成也ハイム浅草 運営管理
 *
 * ※ SPREADSHEET_ID のみ実際の値に変更してください
 */

const CONFIG = {

  // ==============================
  // スプレッドシート設定
  // ==============================
  SPREADSHEET_ID: '1fEMhUhGdZcIDuspTsf3_TUrjXoD0z8lgxD58rWdANA8', // スプレッドシートID

  SHEETS: {
    RESERVATIONS: '予約リスト',   // 予約データ一覧
    COSTS:        '経費入力',     // 月次経費入力
    MONTHLY:      '月別集計',     // 月別KPI集計
    ANNUAL:        '年間集計',    // 年間合計
    DASHBOARD:    'ダッシュボード' // 可視化ダッシュボード
  },

  // ==============================
  // 物件情報
  // ==============================
  PROPERTY: {
    NAME:              '和成也ハイム浅草',
    MAX_ANNUAL_DAYS:   180,  // 民泊法 年間上限日数（4月〜翌3月）
    FISCAL_YEAR_MONTH: 4,    // 事業年度開始月（4月）
    ROOMS:             1     // 部屋数（稼働率計算用）
  },

  // ==============================
  // Gmail 設定
  // ==============================
  GMAIL: {
    PROCESSED_LABEL: '民泊_処理済み', // 処理済みメールに付けるラベル
    SEARCH_DAYS:     180,              // 何日前までのメールを検索するか
    SUBJECTS: {
      AIRBNB: [
        '予約確定',
        'Reservation confirmed',
        'New reservation',
        '予約リクエストが承認されました'
      ],
      BOOKING: [
        '新しい予約',
        'New booking',
        'Booking confirmation',
        '予約確認'
      ]
    },
    FROM: {
      AIRBNB:  ['automated@airbnb.com', 'express@airbnb.com', 'no-reply@airbnb.com', 'email.master@co-reception.com'],
      BOOKING: ['noreply@booking.com', 'customer.service@booking.com', 'noreply-partner@booking.com', 'email.master@co-reception.com'],
      BEDS24:  ['email.master@co-reception.com']
    }
  },

  // ==============================
  // 手数料率
  // ==============================
  COMMISSION_RATE: {
    AIRBNB:  0.03, // ホスト手数料3%
    BOOKING: 0.15, // Booking.com手数料15%
    OTHER:   0.10
  },

  // ==============================
  // 予約リストのカラム定義（A列=1始まり）
  // ==============================
  RESERVATION_COLS: {
    ID:           1,  // A: 予約ID
    PLATFORM:     2,  // B: プラットフォーム
    BOOKED_DATE:  3,  // C: 予約受付日
    CHECKIN:      4,  // D: チェックイン日
    CHECKOUT:     5,  // E: チェックアウト日
    NIGHTS:       6,  // F: 宿泊数
    GUESTS:       7,  // G: 人数
    GUEST_NAME:   8,  // H: ゲスト名
    REVENUE:      9,  // I: 売上（税込）
    COMMISSION:  10,  // J: 手数料
    CLEANING_FEE:11,  // K: 清掃費
    NET_REVENUE: 12,  // L: 純売上（=売上-手数料-清掃費）
    STATUS:      13,  // M: ステータス
    NOTES:       14,  // N: 備考
    EMAIL_ID:    15   // O: GmailメッセージID（重複防止）
  },

  // ==============================
  // 経費入力シートのカラム定義
  // ==============================
  COST_COLS: {
    YEAR_MONTH:    1, // A: 年月（YYYY-MM）
    CLEANING:      2, // B: 清掃費合計
    SUPPLIES:      3, // C: 備品・消耗品費
    UTILITIES:     4, // D: 水光熱費
    RENT:          5, // E: 家賃
    OTHER:         6, // F: その他経費
    NOTES:         7  // G: 備考
  }
};
