/**
 * キャッシュフロー管理システム - Dailyシート管理モジュール
 *
 * 各口座ごとに専用のDailyシートを持つ。
 * 列構成: A:日付 / B:内容 / C:入金 / D:出金 / E:残高 / F:ソース
 *
 * ■ 運用フロー
 *   1. MFから入出金実績を取得 → シートの下に追加
 *   2. 手入力で将来の入出金予定を追加（MFデータの下に日付順で追加）
 *   3. 残高は1行目（前月繰越）から順に自動計算
 */

/**
 * スプレッドシートを取得する
 */
function getCfSpreadsheet() {
  if (CF_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CF_CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ==============================
// MFからDailyシートを更新
// ==============================

/**
 * MFから当月の入出金を取得してDailyシートを更新する
 */
function syncCurrentMonth() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
  const dateTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

  syncToDaily_(dateFrom, dateTo);
}

/**
 * 期間指定でMFデータをDailyシートに同期（ダイアログ）
 */
function syncWithDateRange() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultFrom = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), 1),
    CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd'
  );
  const defaultTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd');

  const html = HtmlService.createHtmlOutput(`
    <html><head><base target="_top">
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { display: block; margin: 8px 0 4px; font-weight: bold; }
      input { width: 100%; padding: 6px; box-sizing: border-box; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; margin-bottom: 8px; }
      .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 12px; }
      button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
      .ok { background: #1a73e8; color: white; border: none; }
      .cancel { background: white; border: 1px solid #ccc; }
      #status { display: none; margin-top: 8px; color: #1a73e8; font-size: 12px; }
    </style></head><body>
      <label>開始日</label><input type="text" id="dateFrom" value="${defaultFrom}">
      <label>終了日</label><input type="text" id="dateTo" value="${defaultTo}">
      <div class="buttons">
        <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
        <button class="ok" id="okBtn" onclick="submit()">同期開始</button>
      </div>
      <div id="status">⏳ MFからデータ取得中...</div>
      <script>
        function submit() {
          document.getElementById('okBtn').disabled = true;
          document.getElementById('status').style.display = 'block';
          var from = document.getElementById('dateFrom').value.replace(/\\//g, '-');
          var to = document.getElementById('dateTo').value.replace(/\\//g, '-');
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) {
              document.getElementById('status').textContent = '❌ ' + e.message;
              document.getElementById('okBtn').disabled = false;
            })
            .syncDateRangeToDaily(from, to);
        }
      </script>
    </body></html>
  `).setWidth(380).setHeight(250);
  ui.showModalDialog(html, 'MFデータ同期（期間指定）');
}

/**
 * 期間指定でMFデータをDailyに同期（ダイアログから呼ばれる）
 */
function syncDateRangeToDaily(dateFrom, dateTo) {
  syncToDaily_(dateFrom, dateTo);
  SpreadsheetApp.getUi().alert(
    `✅ MFデータ同期完了\n\n・期間: ${dateFrom} 〜 ${dateTo}`
  );
}

/**
 * MFデータ同期の共通処理
 */
function syncToDaily_(dateFrom, dateTo) {
  if (!isMfConnected()) {
    throw new Error('MF未連携です。先にMF連携を実行してください。');
  }

  Logger.log(`=== Daily同期開始: ${dateFrom} 〜 ${dateTo} ===`);

  const ss = getCfSpreadsheet();
  const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);

  Object.entries(allTxns).forEach(([accountKey, txns]) => {
    if (txns.length === 0) return;
    const sheetName = CF_CONFIG.ACCOUNTS[accountKey].dailySheet;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`⚠️ シート ${sheetName} が見つかりません`);
      return;
    }
    writeTransactionsToSheet_(sheet, txns);
    recalculateBalances_(sheet);
  });

  // 現残高シート更新
  updateCurrentBalanceSheet_();

  // 日別サマリー更新
  updateDailySummary();

  Logger.log('=== Daily同期完了 ===');
}

// ==============================
// Dailyシートへの書き込み
// ==============================

/**
 * トランザクションをDailyシートに書き込む
 * MFの仕訳を素直に全件追加する（重複チェックなし）
 * ダブりの管理はMF会計側で行う方針
 */
function writeTransactionsToSheet_(sheet, transactions) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;

  transactions.forEach(tx => {
    // 日付順の正しい位置に挿入
    const insertRow = findInsertRow_(sheet, tx.date);

    if (insertRow <= sheet.getLastRow()) {
      sheet.insertRowBefore(insertRow);
    }

    sheet.getRange(insertRow, C.DATE).setValue(tx.date).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(insertRow, C.CONTENT).setValue(tx.content);
    if (tx.deposit > 0) sheet.getRange(insertRow, C.DEPOSIT).setValue(tx.deposit).setNumberFormat('#,##0');
    if (tx.withdrawal > 0) sheet.getRange(insertRow, C.WITHDRAWAL).setValue(tx.withdrawal).setNumberFormat('#,##0');
    sheet.getRange(insertRow, C.SOURCE).setValue(tx.source);
  });
}

/**
 * 日付順の挿入位置を見つける（予定データの手前に挿入）
 */
function findInsertRow_(sheet, targetDate) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return headerRows + 1;

  const numRows = lastRow - headerRows;
  const dates = sheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    if (dates[i][0] instanceof Date && dates[i][0] > targetDate) {
      return i + headerRows + 1;
    }
  }

  return lastRow + 1;
}

// ==============================
// 残高計算
// ==============================

/**
 * Dailyシートの残高を上から再計算する
 * 1行目に前月繰越（残高のみ）がある場合、それを起点にする
 */
function recalculateBalances_(sheet) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const deposits = sheet.getRange(headerRows + 1, C.DEPOSIT, numRows, 1).getValues();
  const withdrawals = sheet.getRange(headerRows + 1, C.WITHDRAWAL, numRows, 1).getValues();
  const balances = sheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

  // 最初の行の残高を前月繰越として使用
  let balance = Number(balances[0][0]) || 0;

  for (let i = 0; i < numRows; i++) {
    const dep = Number(deposits[i][0]) || 0;
    const wth = Number(withdrawals[i][0]) || 0;

    // 前月繰越行（入出金なし・残高のみ）はスキップ
    if (i === 0 && dep === 0 && wth === 0 && balance > 0) continue;

    if (dep === 0 && wth === 0) continue;

    balance = balance + dep - wth;
    sheet.getRange(headerRows + 1 + i, C.BALANCE).setValue(balance).setNumberFormat('#,##0');
  }

  // アラート色付け（PayPay 005のみ）
  applyAlertColors_(sheet);
}

/**
 * アラート色付け
 */
function applyAlertColors_(sheet) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  // このシートがアラート対象口座かチェック
  const alertAccount = CF_CONFIG.ALERT.ALERT_ACCOUNT;
  const alertSheetName = CF_CONFIG.ACCOUNTS[alertAccount].dailySheet;
  if (sheet.getName() !== alertSheetName) return;

  const numRows = lastRow - headerRows;
  const balances = sheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const bal = Number(balances[i][0]) || 0;
    if (bal === 0) continue;

    const cell = sheet.getRange(headerRows + 1 + i, C.BALANCE);
    if (bal <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
      cell.setBackground('#ffcdd2').setFontColor('#b71c1c').setFontWeight('bold');
    } else if (bal <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
      cell.setBackground('#fff9c4').setFontColor('#f57f17').setFontWeight('bold');
    } else {
      cell.setBackground(null).setFontColor('#000000').setFontWeight('normal');
    }
  }
}

// ==============================
// 手入力の予定管理
// ==============================

/**
 * 入出金予定を手入力するダイアログ
 */
function addPlannedTransaction() {
  const accountOptions = Object.entries(CF_CONFIG.ACCOUNTS)
    .map(([key, acct]) => `<option value="${key}">${acct.shortName}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
    <html><head><base target="_top">
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { display: block; margin: 8px 0 4px; font-weight: bold; color: #555; font-size: 12px; }
      input, select { width: 100%; padding: 6px; box-sizing: border-box; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; }
      .row { display: flex; gap: 8px; }
      .row > div { flex: 1; }
      .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
      button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
      .ok { background: #1a73e8; color: white; border: none; }
      .cancel { background: white; border: 1px solid #ccc; }
    </style></head><body>
      <h3 style="margin:0 0 12px; color:#1a73e8;">📝 入出金予定の登録</h3>
      <label>口座</label>
      <select id="account">${accountOptions}</select>
      <label>日付</label>
      <input type="text" id="date" placeholder="yyyy/MM/dd">
      <label>内容</label>
      <input type="text" id="content" placeholder="例: Amazon売上、家賃、JALカード">
      <div class="row">
        <div><label>入金額</label><input type="text" id="deposit" placeholder="0"></div>
        <div><label>出金額</label><input type="text" id="withdrawal" placeholder="0"></div>
      </div>
      <div class="buttons">
        <button class="cancel" onclick="google.script.host.close()">閉じる</button>
        <button class="ok" onclick="submit()">登録</button>
      </div>
      <script>
        function submit() {
          var data = {
            account: document.getElementById('account').value,
            date: document.getElementById('date').value,
            content: document.getElementById('content').value,
            deposit: document.getElementById('deposit').value || '0',
            withdrawal: document.getElementById('withdrawal').value || '0'
          };
          google.script.run
            .withSuccessHandler(function() {
              document.getElementById('date').value = '';
              document.getElementById('content').value = '';
              document.getElementById('deposit').value = '';
              document.getElementById('withdrawal').value = '';
              alert('✅ 登録しました');
            })
            .withFailureHandler(function(e) { alert('❌ ' + e.message); })
            .savePlannedTransaction(data);
        }
      </script>
    </body></html>
  `).setWidth(400).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, '入出金予定の登録');
}

/**
 * 予定データをDailyシートに保存する
 */
function savePlannedTransaction(data) {
  const ss = getCfSpreadsheet();
  const sheetName = CF_CONFIG.ACCOUNTS[data.account].dailySheet;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート ${sheetName} が見つかりません。`);

  const C = CF_CONFIG.DAILY_COLS;
  const date = new Date(data.date.replace(/\//g, '-'));
  if (isNaN(date.getTime())) throw new Error('日付の形式が正しくありません。');

  const deposit = parseInt(String(data.deposit).replace(/[,、]/g, '')) || 0;
  const withdrawal = parseInt(String(data.withdrawal).replace(/[,、]/g, '')) || 0;
  if (deposit === 0 && withdrawal === 0) throw new Error('入金額または出金額を入力してください。');

  // 日付順の正しい位置に挿入
  const insertRow = findInsertRow_(sheet, date);
  if (insertRow <= sheet.getLastRow()) {
    sheet.insertRowBefore(insertRow);
  }

  sheet.getRange(insertRow, C.DATE).setValue(date).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(insertRow, C.CONTENT).setValue(data.content);
  if (deposit > 0) sheet.getRange(insertRow, C.DEPOSIT).setValue(deposit).setNumberFormat('#,##0');
  if (withdrawal > 0) sheet.getRange(insertRow, C.WITHDRAWAL).setValue(withdrawal).setNumberFormat('#,##0');
  sheet.getRange(insertRow, C.SOURCE).setValue(CF_CONFIG.SOURCE.PLANNED);

  // 予定行は薄い黄色
  sheet.getRange(insertRow, C.DATE, 1, 6).setBackground('#fff9c4');

  // 残高再計算
  recalculateBalances_(sheet);
}

// ==============================
// ヘルパー関数
// ==============================

/**
 * 指定口座のDailyシートから前月繰越残高を取得
 */
function getCarryForwardBalance_(accountKey) {
  const ss = getCfSpreadsheet();
  const sheetName = CF_CONFIG.ACCOUNTS[accountKey].dailySheet;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  if (sheet.getLastRow() <= headerRows) return 0;

  const firstBalance = sheet.getRange(headerRows + 1, C.BALANCE).getValue();
  const firstDeposit = sheet.getRange(headerRows + 1, C.DEPOSIT).getValue();
  const firstWithdrawal = sheet.getRange(headerRows + 1, C.WITHDRAWAL).getValue();

  if ((!firstDeposit || firstDeposit === 0) && (!firstWithdrawal || firstWithdrawal === 0) && firstBalance > 0) {
    return Number(firstBalance);
  }
  return 0;
}

/**
 * 指定口座のDailyシートから最新残高を取得
 */
function getLatestBalance_(accountKey) {
  const ss = getCfSpreadsheet();
  const sheetName = CF_CONFIG.ACCOUNTS[accountKey].dailySheet;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return 0;

  const numRows = lastRow - headerRows;
  const balances = sheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

  for (let i = numRows - 1; i >= 0; i--) {
    const val = Number(balances[i][0]);
    if (val && val !== 0) return val;
  }
  return 0;
}

// ==============================
// 日別サマリーシート
// ==============================

/**
 * 日別サマリーシートを更新する
 *
 * 3口座のDailyシートから全日付を収集し、日付ごとに
 * 入金合計・出金合計・各口座残高・全体残高を表示する。
 *
 * 列: A:日付 / B:入金合計 / C:出金合計 / D:全体残高 / E:PayPay005残高 / F:PayPay003残高 / G:西武信金残高
 */
function updateDailySummary() {
  const ss = getCfSpreadsheet();
  let sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY_SUMMARY);

  if (!sheet) {
    sheet = ss.insertSheet(CF_CONFIG.SHEETS.DAILY_SUMMARY);
  }

  const headers = ['日付', '入金合計', '出金合計', '全体残高',
    'PayPay 005\n残高', 'PayPay 003\n残高', '西武信用金庫\n残高'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true);

  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const accountKeys = Object.keys(CF_CONFIG.ACCOUNTS);

  // 全口座のデータを日付→{入金, 出金, 残高}のマップに集約
  const dateMap = {};
  const allDates = new Set();

  accountKeys.forEach(key => {
    const acctSheet = ss.getSheetByName(CF_CONFIG.ACCOUNTS[key].dailySheet);
    if (!acctSheet || acctSheet.getLastRow() <= headerRows) return;

    const numRows = acctSheet.getLastRow() - headerRows;
    const dates = acctSheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();
    const deposits = acctSheet.getRange(headerRows + 1, C.DEPOSIT, numRows, 1).getValues();
    const withdrawals = acctSheet.getRange(headerRows + 1, C.WITHDRAWAL, numRows, 1).getValues();
    const balances = acctSheet.getRange(headerRows + 1, C.BALANCE, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const d = dates[i][0];
      if (!(d instanceof Date)) continue;

      const dateKey = Utilities.formatDate(d, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');
      allDates.add(dateKey);

      if (!dateMap[dateKey]) {
        dateMap[dateKey] = { date: d, deposits: {}, withdrawals: {}, balances: {} };
      }

      const dep = Number(deposits[i][0]) || 0;
      const wth = Number(withdrawals[i][0]) || 0;
      dateMap[dateKey].deposits[key] = (dateMap[dateKey].deposits[key] || 0) + dep;
      dateMap[dateKey].withdrawals[key] = (dateMap[dateKey].withdrawals[key] || 0) + wth;

      const bal = Number(balances[i][0]) || 0;
      if (bal !== 0) dateMap[dateKey].balances[key] = bal;
    }
  });

  // 日付順にソート
  const sortedDates = Array.from(allDates).sort();
  if (sortedDates.length === 0) return;

  // 各口座の前回残高を追跡
  const latestBalance = {};
  accountKeys.forEach(key => { latestBalance[key] = 0; });

  const rows = [];

  sortedDates.forEach(dateKey => {
    const data = dateMap[dateKey];
    let totalDeposit = 0;
    let totalWithdrawal = 0;

    accountKeys.forEach(key => {
      totalDeposit += data.deposits[key] || 0;
      totalWithdrawal += data.withdrawals[key] || 0;
      if (data.balances[key] !== undefined) {
        latestBalance[key] = data.balances[key];
      }
    });

    const totalBalance = accountKeys.reduce((sum, key) => sum + latestBalance[key], 0);

    rows.push([
      data.date,
      totalDeposit || '',
      totalWithdrawal || '',
      totalBalance,
      latestBalance.CF005,
      latestBalance.CF003,
      latestBalance.SEIBU
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(2, 2, rows.length, headers.length - 1).setNumberFormat('#,##0');

    // アラート色付け（PayPay005残高ベース）
    for (let i = 0; i < rows.length; i++) {
      const cf005Bal = rows[i][4];
      const totalCell = sheet.getRange(i + 2, 4);

      if (cf005Bal > 0 && cf005Bal <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
        totalCell.setBackground('#ffcdd2').setFontColor('#b71c1c').setFontWeight('bold');
      } else if (cf005Bal > 0 && cf005Bal <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
        totalCell.setBackground('#fff9c4').setFontColor('#f57f17').setFontWeight('bold');
      }
    }
  }

  sheet.setColumnWidth(1, 90);
  for (let i = 2; i <= headers.length; i++) sheet.setColumnWidth(i, 110);
  sheet.setFrozenRows(1);

  Logger.log(`日別サマリー更新完了: ${rows.length}日分`);
}

// ==============================
// 実口座残高の自動更新
// ==============================

/**
 * 実口座残高シートの指定月にMFの残高試算表データを自動入力する
 *
 * MF API: GET /reports/trial_balance_bs
 * 取得項目: 普通預金, 売掛金, 未払金, 未払費用, 預り金, 商品（棚卸資産）, 長期借入金
 */
function updateRealBalance() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultYear = today.getFullYear();
  const defaultMonth = today.getMonth(); // 前月

  const result = ui.prompt(
    '実口座残高の更新',
    `更新する年月を入力してください（例: 2026/03）\n\n全期間一括更新する場合は「ALL」と入力\n\nMF会計の残高試算表から自動取得します。`,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const input = result.getResponseText().trim();

  if (input.toUpperCase() === 'ALL') {
    updateRealBalanceAll_();
    return;
  }

  const parts = input.split(/[\/\-\.]/);
  if (parts.length !== 2) {
    ui.alert('⚠️ 年月の形式が正しくありません（例: 2026/03 または ALL）');
    return;
  }
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);

  if (!isMfConnected()) {
    ui.alert('❌ MF未連携です。先にMF連携を実行してください。');
    return;
  }

  try {
    // MFの残高試算表（BS）を取得
    const bsData = mfApiRequest_('/reports/trial_balance_bs', {
      start_date: `${year}-${String(month).padStart(2, '0')}-01`,
      end_date: `${year}-${String(month).padStart(2, '0')}-${new Date(year, month, 0).getDate()}`
    });

    const balanceMap = extractBalancesFromRows_(bsData.rows || []);
    Logger.log('BS残高マップ: ' + JSON.stringify(balanceMap).substring(0, 500));

    // 必要な科目の残高を取得
    const deposits = (balanceMap['普通預金'] || 0);
    const receivables = (balanceMap['売掛金'] || 0);
    const payables = (balanceMap['未払金'] || 0);
    const accruedExpenses = (balanceMap['未払費用'] || 0);
    const depositsReceived = (balanceMap['預り金'] || 0);
    const inventory = (balanceMap['商品'] || balanceMap['商品及び製品'] || 0);
    const longTermDebt = (balanceMap['長期借入金'] || 0);

    // 実口座残高シートの該当行を特定
    const ss = getCfSpreadsheet();
    const sheet = ss.getSheetByName('実口座残高');
    if (!sheet) throw new Error('実口座残高シートが見つかりません。');

    const targetRow = findRealBalanceRow_(sheet, year, month);
    if (targetRow === 0) throw new Error(`${year}年${month}月の行が見つかりません。`);

    // 自動入力（MFデータ）※Amazon残高列削除後の列番号
    sheet.getRange(targetRow, 3).setValue(deposits);         // C: 普通預金
    sheet.getRange(targetRow, 4).setValue(receivables);      // D: 売掛金
    sheet.getRange(targetRow, 5).setValue(payables);         // E: 未払金
    sheet.getRange(targetRow, 6).setValue(accruedExpenses);  // F: 未払費用
    sheet.getRange(targetRow, 7).setValue(depositsReceived); // G: 預り金
    sheet.getRange(targetRow, 9).setValue(inventory);        // I: 商品在庫
    sheet.getRange(targetRow, 14).setValue(longTermDebt);    // N: 融資残高

    ui.alert(
      `✅ ${year}年${month}月の実口座残高を更新しました\n\n` +
      `・普通預金: ¥${deposits.toLocaleString()}\n` +
      `・売掛金: ¥${receivables.toLocaleString()}\n` +
      `・未払金: ¥${payables.toLocaleString()}\n` +
      `・未払費用: ¥${accruedExpenses.toLocaleString()}\n` +
      `・預り金: ¥${depositsReceived.toLocaleString()}\n` +
      `・商品在庫: ¥${inventory.toLocaleString()}\n` +
      `・融資残高: ¥${longTermDebt.toLocaleString()}\n\n` +
      `※ Amazon残高・想定売上・入金割合は手入力してください。`
    );

  } catch (e) {
    ui.alert(`❌ エラー\n\n${e.message}`);
    Logger.log(`実口座残高更新エラー: ${e.message}\n${e.stack}`);
  }
}

/**
 * 実口座残高シートから指定年月の行を検索
 */
/**
 * 全期間の実口座残高を一括更新する
 * MFの会計年度ごとに残高試算表を取得して各月に反映
 */
function updateRealBalanceAll_() {
  const ui = SpreadsheetApp.getUi();
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName('実口座残高');
  if (!sheet) throw new Error('実口座残高シートが見つかりません。');

  // 実口座残高シートの全行を走査して年月を取得
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  const labels = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  let updated = 0;

  for (let i = 0; i < labels.length; i++) {
    const label = String(labels[i][0]);
    const match = label.match(/^(\d{4})\.(\d{2})$/);
    if (!match) continue;

    const year = parseInt(match[1]);
    const month = parseInt(match[2]);
    const row = i + 3;

    try {
      const lastDay = new Date(year, month, 0).getDate();
      const bsData = mfApiRequest_('/reports/trial_balance_bs', {
        start_date: `${year}-${String(month).padStart(2, '0')}-01`,
        end_date: `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`
      });

      const balanceMap = extractBalancesFromRows_(bsData.rows || []);

      sheet.getRange(row, 3).setValue(balanceMap['普通預金'] || 0);         // C
      sheet.getRange(row, 4).setValue(balanceMap['売掛金'] || 0);         // D
      sheet.getRange(row, 5).setValue(balanceMap['未払金'] || 0);         // E
      sheet.getRange(row, 6).setValue(balanceMap['未払費用'] || 0);       // F
      sheet.getRange(row, 7).setValue(balanceMap['預り金'] || 0);         // G
      sheet.getRange(row, 9).setValue(balanceMap['商品'] || balanceMap['商品及び製品'] || 0);  // I
      sheet.getRange(row, 14).setValue(balanceMap['長期借入金'] || 0);    // N

      updated++;
      Logger.log(`${label}: 更新完了`);

      // レート制限対策
      Utilities.sleep(500);
    } catch (e) {
      Logger.log(`${label}: エラー - ${e.message}`);
      // 会計年度外の月はスキップ
    }
  }

  ui.alert(`✅ 実口座残高を一括更新しました\n\n・更新: ${updated}ヶ月分\n\n※ Amazon残高・入金割合は手入力してください。`);
}

function findRealBalanceRow_(sheet, year, month) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return 0;

  // 年月ラベル（2025.04形式）で検索
  const targetLabel = `${year}.${String(month).padStart(2, '0')}`;

  const colA = sheet.getRange(3, 1, lastRow - 2, 1).getValues();

  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === targetLabel) {
      return i + 3;
    }
  }

  return 0;
}
