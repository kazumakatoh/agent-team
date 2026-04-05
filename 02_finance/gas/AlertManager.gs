/**
 * キャッシュフロー管理システム - アラート管理モジュール
 *
 * PayPay 005口座の残高を監視し、キャッシュアウトリスクを検知する。
 *  - 500万円以下 → 🔴 危険
 *  - 1,000万円以下 → 🟡 注意
 *  - 1,000万円超 → 🟢 正常
 */

/**
 * Dailyシートのキャッシュアウトリスクをチェックする
 * 予定を含む将来の残高推移から、危険水準を検知する。
 */
function checkCashOutRisk() {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return;

  const alertAccount = CF_CONFIG.ALERT.ALERT_ACCOUNT;
  const cols = CF_CONFIG.ACCOUNTS[alertAccount].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const dates = sheet.getRange(headerRows + 1, cols.DATE, numRows, 1).getValues();
  const balances = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();

  const alerts = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 0; i < numRows; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;

    const balance = Number(balances[i][0]) || 0;
    if (balance === 0) continue;

    const dateStr = Utilities.formatDate(d, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd');

    if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
      alerts.push({
        level: 'DANGER',
        date: dateStr,
        balance: balance,
        row: i + headerRows + 1
      });
    } else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
      alerts.push({
        level: 'WARNING',
        date: dateStr,
        balance: balance,
        row: i + headerRows + 1
      });
    }
  }

  // Dailyシートの残高セルに色付け
  applyAlertColors_(sheet, cols, headerRows, numRows);

  // アラートがあれば通知
  if (alerts.length > 0) {
    notifyAlerts_(alerts);
  }

  return alerts;
}

/**
 * Dailyシートの残高セルにアラート色を適用する
 */
function applyAlertColors_(sheet, cols, headerRows, numRows) {
  const balances = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();
  const balanceRange = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1);

  // 一旦リセット
  balanceRange.setBackground(null).setFontColor('#000000');

  for (let i = 0; i < numRows; i++) {
    const balance = Number(balances[i][0]) || 0;
    if (balance === 0) continue;

    const cell = sheet.getRange(headerRows + 1 + i, cols.BALANCE);

    if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
      // 🔴 危険: 赤背景
      cell.setBackground('#ffcdd2').setFontColor('#b71c1c').setFontWeight('bold');
    } else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
      // 🟡 注意: 黄背景
      cell.setBackground('#fff9c4').setFontColor('#f57f17').setFontWeight('bold');
    } else {
      // 🟢 正常
      cell.setFontWeight('normal');
    }
  }

  // A列（合計）にも色付け
  const totalCol = CF_CONFIG.DAILY_TOTAL_COL;
  const totals = sheet.getRange(headerRows + 1, totalCol, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const total = Number(totals[i][0]) || 0;
    if (total === 0) continue;

    const cell = sheet.getRange(headerRows + 1 + i, totalCol);

    // 合計も同じ基準でチェック（3口座合計で見るとより安全）
    if (total <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
      cell.setBackground('#ffcdd2').setFontColor('#b71c1c').setFontWeight('bold');
    } else if (total <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
      cell.setBackground('#fff9c4').setFontColor('#f57f17');
    }
  }
}

/**
 * アラート通知を表示する
 * @param {Array<Object>} alerts
 */
function notifyAlerts_(alerts) {
  const dangerAlerts = alerts.filter(a => a.level === 'DANGER');
  const warningAlerts = alerts.filter(a => a.level === 'WARNING');

  let msg = '⚠️ キャッシュアウトリスク検知\n\n';
  msg += `監視口座: ${CF_CONFIG.ACCOUNTS[CF_CONFIG.ALERT.ALERT_ACCOUNT].shortName}\n\n`;

  if (dangerAlerts.length > 0) {
    msg += '🔴【危険】500万円以下\n';
    dangerAlerts.slice(0, 5).forEach(a => {
      msg += `  ${a.date}: ¥${Number(a.balance).toLocaleString()}\n`;
    });
    if (dangerAlerts.length > 5) msg += `  ...他 ${dangerAlerts.length - 5}件\n`;
    msg += '\n';
  }

  if (warningAlerts.length > 0) {
    msg += '🟡【注意】1,000万円以下\n';
    warningAlerts.slice(0, 5).forEach(a => {
      msg += `  ${a.date}: ¥${Number(a.balance).toLocaleString()}\n`;
    });
    if (warningAlerts.length > 5) msg += `  ...他 ${warningAlerts.length - 5}件\n`;
  }

  // UIが利用可能な場合のみダイアログ表示
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    // トリガー実行時はUIなし → ログのみ
    Logger.log(msg);
  }
}

/**
 * 現在残高のサマリーを表示する
 */
function showBalanceSummary() {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return;

  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) {
    SpreadsheetApp.getUi().alert('Dailyシートにデータがありません。');
    return;
  }

  let msg = '📊 現在の口座残高\n\n';
  let totalBalance = 0;

  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    // 最新の残高を取得（最終行から逆方向に探索）
    const balance = getLatestBalance_(sheet, account.daily, headerRows, lastRow);
    totalBalance += balance;

    // アラートレベル判定（CF005のみ）
    let indicator = '';
    if (key === CF_CONFIG.ALERT.ALERT_ACCOUNT) {
      if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) indicator = ' 🔴';
      else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) indicator = ' 🟡';
      else indicator = ' 🟢';
    }

    msg += `${account.shortName}: ¥${Number(balance).toLocaleString()}${indicator}\n`;
  });

  msg += `\n━━━━━━━━━━━━━━━\n`;
  msg += `合計: ¥${Number(totalBalance).toLocaleString()}\n`;

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 指定口座の最新残高を取得
 */
function getLatestBalance_(sheet, cols, headerRows, lastRow) {
  const numRows = lastRow - headerRows;
  const balances = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();

  for (let i = numRows - 1; i >= 0; i--) {
    const val = Number(balances[i][0]);
    if (val && val !== 0) return val;
  }
  return 0;
}

/**
 * 日次トリガーで実行: アラートチェック＋MFデータ同期
 */
function dailyCashFlowCheck() {
  try {
    Logger.log('=== 日次キャッシュフローチェック開始 ===');

    if (isMfConnected()) {
      // 当月のMFデータを同期
      const today = new Date();
      const dateFrom = Utilities.formatDate(
        new Date(today.getFullYear(), today.getMonth(), 1),
        CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd'
      );
      const dateTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

      const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);
      const ss = getCfSpreadsheet();
      const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);

      if (sheet) {
        for (const [accountKey, txns] of Object.entries(allTxns)) {
          if (txns.length > 0) writeTransactionsToDaily_(sheet, accountKey, txns);
        }

        Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => recalculateBalances(key));
        updateDailyTotals_();
      }
    }

    // アラートチェック
    checkCashOutRisk();

    Logger.log('=== 日次チェック完了 ===');
  } catch (e) {
    Logger.log(`日次チェックエラー: ${e.message}\n${e.stack}`);
  }
}
