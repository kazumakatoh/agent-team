/**
 * キャッシュフロー管理システム - アラート管理モジュール
 *
 * ALERT_ACCOUNTSに指定した口座の残高を監視し、キャッシュアウトリスクを検知する。
 *  - 500万円以下 → 🔴 危険
 *  - 1,000万円以下 → 🟡 注意
 *  - 1,000万円超 → 🟢 正常
 *
 * ※ 各Dailyシートの残高セル色付けは DailySheet.gs の applyAlertColors_ で実装
 */

/**
 * 監視対象口座のキャッシュアウトリスクをチェックする
 * ALERT_ACCOUNTSに指定された全口座の全日付の残高を確認
 * @return {Array<Object>} アラート配列（account情報付き）
 */
function checkCashOutRisk() {
  const ss = getCfSpreadsheet();
  const alertAccounts = CF_CONFIG.ALERT.ALERT_ACCOUNTS || [];
  const alerts = [];

  alertAccounts.forEach(accountKey => {
    const account = CF_CONFIG.ACCOUNTS[accountKey];
    if (!account) return;

    const sheet = ss.getSheetByName(account.dailySheet);
    if (!sheet) return;

    const C = CF_CONFIG.DAILY_COLS;
    const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRows) return;

    const numRows = lastRow - headerRows;
    const dates = sheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();
    const balances = sheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const d = dates[i][0];
      if (!(d instanceof Date)) continue;

      const balance = Number(balances[i][0]) || 0;
      if (balance === 0) continue;

      const dateStr = Utilities.formatDate(d, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd');

      if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
        alerts.push({
          level: 'DANGER', date: dateStr, balance: balance,
          accountKey: accountKey, accountName: account.shortName
        });
      } else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
        alerts.push({
          level: 'WARNING', date: dateStr, balance: balance,
          accountKey: accountKey, accountName: account.shortName
        });
      }
    }
  });

  return alerts;
}
