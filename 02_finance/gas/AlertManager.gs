/**
 * キャッシュフロー管理システム - アラート管理モジュール
 *
 * PayPay 005口座の残高を監視し、キャッシュアウトリスクを検知する。
 *  - 500万円以下 → 🔴 危険
 *  - 1,000万円以下 → 🟡 注意
 *  - 1,000万円超 → 🟢 正常
 *
 * ※ 各Dailyシートの残高セル色付けは DailySheet.gs の applyAlertColors_ で実装
 */

/**
 * 全口座のキャッシュアウトリスクをチェックする
 * PayPay 005（監視対象口座）の全日付の残高を確認
 * @return {Array<Object>} アラート配列
 */
function checkCashOutRisk() {
  const ss = getCfSpreadsheet();
  const alertAccount = CF_CONFIG.ALERT.ALERT_ACCOUNT;
  const sheetName = CF_CONFIG.ACCOUNTS[alertAccount].dailySheet;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return [];

  const numRows = lastRow - headerRows;
  const dates = sheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();
  const balances = sheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

  const alerts = [];

  for (let i = 0; i < numRows; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;

    const balance = Number(balances[i][0]) || 0;
    if (balance === 0) continue;

    const dateStr = Utilities.formatDate(d, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd');

    if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
      alerts.push({ level: 'DANGER', date: dateStr, balance: balance });
    } else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
      alerts.push({ level: 'WARNING', date: dateStr, balance: balance });
    }
  }

  return alerts;
}
