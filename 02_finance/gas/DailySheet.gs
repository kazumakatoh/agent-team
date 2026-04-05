/**
 * キャッシュフロー管理システム - Dailyシート管理モジュール
 *
 * MFから取得した入出金実績と、手入力の予定を統合して
 * 日次のキャッシュフローを管理する。
 *
 * ■ 列構成（3口座が横に並ぶ）
 *   A: 3口座合計残高
 *   B-H: PayPay 005（日付/内容/入金/出金/残高/ソース）
 *   I: （空白）
 *   J-O: PayPay 003
 *   P: （空白）
 *   Q-V: 西武信金
 */

/**
 * スプレッドシートを取得する
 * @return {Spreadsheet}
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
  syncMonthToDaily(year, month);
}

/**
 * 指定月のMFデータでDailyシートを更新する
 * @param {number} year
 * @param {number} month
 */
function syncMonthToDaily(year, month) {
  if (!isMfConnected()) {
    SpreadsheetApp.getUi().alert('❌ MF未連携です。先にMF連携を実行してください。');
    return;
  }

  const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const dateTo = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;

  Logger.log(`=== Daily同期開始: ${dateFrom} 〜 ${dateTo} ===`);

  // 全口座の入出金を取得
  const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);

  // 各口座をDailyシートに反映
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) throw new Error('Dailyシートが見つかりません。Setupを実行してください。');

  let totalAdded = 0;
  let totalUpdated = 0;

  for (const [accountKey, txns] of Object.entries(allTxns)) {
    if (txns.length === 0) continue;
    const result = writeTransactionsToDaily_(sheet, accountKey, txns);
    totalAdded += result.added;
    totalUpdated += result.updated;
  }

  // 全口座の残高を再計算
  Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => {
    recalculateBalances(key);
  });

  // 3口座合計を更新
  updateDailyTotals_();

  // アラートチェック
  checkCashOutRisk();

  Logger.log(`=== Daily同期完了: 追加${totalAdded}件 / 更新${totalUpdated}件 ===`);

  SpreadsheetApp.getUi().alert(
    `✅ MFデータ同期完了\n\n` +
    `・期間: ${year}年${month}月\n` +
    `・新規追加: ${totalAdded}件\n` +
    `・上書更新: ${totalUpdated}件`
  );
}

/**
 * MFデータ同期（ダイアログで期間指定）
 */
function syncWithDateRange() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultFrom = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), 1),
    CF_CONFIG.DISPLAY.TIMEZONE,
    'yyyy/MM/dd'
  );
  const defaultTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd');

  const html = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; font-size: 13px; }
        label { display: block; margin-bottom: 4px; font-weight: bold; color: #555; }
        input { width: 100%; padding: 8px; box-sizing: border-box; font-size: 14px;
                border: 1px solid #ccc; border-radius: 4px; margin-bottom: 12px; }
        .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px; }
        button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
        .ok { background: #1a73e8; color: white; border: none; }
        .cancel { background: white; border: 1px solid #ccc; }
        #status { display: none; margin-top: 12px; color: #1a73e8; font-size: 12px; }
      </style>
    </head>
    <body>
      <label>開始日</label>
      <input type="text" id="dateFrom" value="${defaultFrom}">
      <label>終了日</label>
      <input type="text" id="dateTo" value="${defaultTo}">
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
    </body>
    </html>
  `).setWidth(380).setHeight(280);

  ui.showModalDialog(html, 'MFデータ同期（期間指定）');
}

/**
 * 期間指定でMFデータをDailyシートに同期（ダイアログから呼ばれる）
 */
function syncDateRangeToDaily(dateFrom, dateTo) {
  const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);

  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) throw new Error('Dailyシートが見つかりません。');

  let totalAdded = 0;
  let totalUpdated = 0;

  for (const [accountKey, txns] of Object.entries(allTxns)) {
    if (txns.length === 0) continue;
    const result = writeTransactionsToDaily_(sheet, accountKey, txns);
    totalAdded += result.added;
    totalUpdated += result.updated;
  }

  Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => recalculateBalances(key));
  updateDailyTotals_();
  checkCashOutRisk();

  SpreadsheetApp.getUi().alert(
    `✅ MFデータ同期完了\n\n` +
    `・期間: ${dateFrom} 〜 ${dateTo}\n` +
    `・新規追加: ${totalAdded}件\n` +
    `・上書更新: ${totalUpdated}件`
  );
}

// ==============================
// Dailyシートへの書き込み
// ==============================

/**
 * トランザクション配列をDailyシートの指定口座列に書き込む
 * @param {Sheet} sheet - Dailyシート
 * @param {string} accountKey - 口座キー
 * @param {Array<Object>} transactions - 入出金データ
 * @return {Object} { added, updated }
 */
function writeTransactionsToDaily_(sheet, accountKey, transactions) {
  const cols = CF_CONFIG.ACCOUNTS[accountKey].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  let added = 0;
  let updated = 0;

  transactions.forEach(tx => {
    // 同じ日付・内容・ソースのデータがあれば上書き
    const existingRow = findExistingDailyRow_(sheet, cols, tx.date, tx.content);

    if (existingRow > 0) {
      // 既存行を更新
      sheet.getRange(existingRow, cols.CONTENT).setValue(tx.content);
      if (tx.deposit > 0) {
        sheet.getRange(existingRow, cols.DEPOSIT).setValue(tx.deposit);
        sheet.getRange(existingRow, cols.WITHDRAWAL).setValue('');
      } else {
        sheet.getRange(existingRow, cols.DEPOSIT).setValue('');
        sheet.getRange(existingRow, cols.WITHDRAWAL).setValue(tx.withdrawal);
      }
      sheet.getRange(existingRow, cols.SOURCE).setValue(CF_CONFIG.SOURCE.MF);
      updated++;
    } else {
      // 日付順の正しい位置に挿入
      const insertRow = findInsertRowForDate_(sheet, cols.DATE, tx.date, headerRows);
      sheet.insertRowAfter(Math.max(insertRow - 1, headerRows));

      const newRow = insertRow;
      sheet.getRange(newRow, cols.DATE).setValue(tx.date).setNumberFormat('yyyy/MM/dd');
      sheet.getRange(newRow, cols.CONTENT).setValue(tx.content);
      if (tx.deposit > 0) {
        sheet.getRange(newRow, cols.DEPOSIT).setValue(tx.deposit).setNumberFormat('#,##0');
      }
      if (tx.withdrawal > 0) {
        sheet.getRange(newRow, cols.WITHDRAWAL).setValue(tx.withdrawal).setNumberFormat('#,##0');
      }
      sheet.getRange(newRow, cols.SOURCE).setValue(CF_CONFIG.SOURCE.MF);
      added++;
    }
  });

  return { added, updated };
}

/**
 * 既存のDailyデータから同一日付・内容の行を検索
 */
function findExistingDailyRow_(sheet, cols, date, content) {
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return 0;

  const numRows = lastRow - headerRows;
  const dateRange = sheet.getRange(headerRows + 1, cols.DATE, numRows, 1).getValues();
  const contentRange = sheet.getRange(headerRows + 1, cols.CONTENT, numRows, 1).getValues();
  const sourceRange = sheet.getRange(headerRows + 1, cols.SOURCE, numRows, 1).getValues();

  const targetDate = Utilities.formatDate(date, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

  for (let i = 0; i < numRows; i++) {
    if (!(dateRange[i][0] instanceof Date)) continue;
    const rowDate = Utilities.formatDate(dateRange[i][0], CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');
    const rowSource = String(sourceRange[i][0]);

    if (rowDate === targetDate
        && String(contentRange[i][0]).trim() === String(content).trim()
        && rowSource === CF_CONFIG.SOURCE.MF) {
      return i + headerRows + 1;
    }
  }

  return 0;
}

/**
 * 日付順の挿入位置を取得
 */
function findInsertRowForDate_(sheet, dateCol, targetDate, headerRows) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return headerRows + 1;

  const dates = sheet.getRange(headerRows + 1, dateCol, lastRow - headerRows, 1).getValues();

  for (let i = 0; i < dates.length; i++) {
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
 * 指定口座の残高を上から再計算する
 * @param {string} accountKey - 口座キー
 */
function recalculateBalances(accountKey) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return;

  const cols = CF_CONFIG.ACCOUNTS[accountKey].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const deposits = sheet.getRange(headerRows + 1, cols.DEPOSIT, numRows, 1).getValues();
  const withdrawals = sheet.getRange(headerRows + 1, cols.WITHDRAWAL, numRows, 1).getValues();
  const balances = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();

  // 最初の行に残高がある場合はそれを初期値として使用
  // なければ0から計算
  let balance = 0;

  // 最初のMFデータの残高から逆算して初期残高を求める
  const firstDeposit = Number(deposits[0][0]) || 0;
  const firstWithdrawal = Number(withdrawals[0][0]) || 0;
  const firstBalance = Number(balances[0][0]) || 0;

  if (firstBalance > 0) {
    // MFから残高が取れている場合、初期残高を逆算
    balance = firstBalance;
    sheet.getRange(headerRows + 1, cols.BALANCE).setValue(balance).setNumberFormat('#,##0');

    // 2行目以降を計算
    for (let i = 1; i < numRows; i++) {
      const dep = Number(deposits[i][0]) || 0;
      const wth = Number(withdrawals[i][0]) || 0;
      if (dep === 0 && wth === 0) continue;
      balance = balance + dep - wth;
      sheet.getRange(headerRows + 1 + i, cols.BALANCE).setValue(balance).setNumberFormat('#,##0');
    }
  } else {
    // 残高情報がない場合は入出金の差分のみ
    for (let i = 0; i < numRows; i++) {
      const dep = Number(deposits[i][0]) || 0;
      const wth = Number(withdrawals[i][0]) || 0;
      if (dep === 0 && wth === 0) continue;
      balance = balance + dep - wth;
      sheet.getRange(headerRows + 1 + i, cols.BALANCE).setValue(balance).setNumberFormat('#,##0');
    }
  }
}

/**
 * 3口座の合計残高（A列）を更新する
 */
function updateDailyTotals_() {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return;

  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const accountKeys = Object.keys(CF_CONFIG.ACCOUNTS);

  // 各口座の残高列を取得
  const balanceCols = {};
  accountKeys.forEach(key => {
    const cols = CF_CONFIG.ACCOUNTS[key].daily;
    balanceCols[key] = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();
  });

  // A列に3口座合計を書き込み
  const totals = [];
  for (let i = 0; i < numRows; i++) {
    let total = 0;
    let hasData = false;
    accountKeys.forEach(key => {
      const val = Number(balanceCols[key][i][0]) || 0;
      if (val !== 0) hasData = true;
      total += val;
    });
    totals.push([hasData ? total : '']);
  }

  sheet.getRange(headerRows + 1, CF_CONFIG.DAILY_TOTAL_COL, numRows, 1)
    .setValues(totals)
    .setNumberFormat('#,##0');
}

// ==============================
// 手入力の予定管理
// ==============================

/**
 * 入出金予定を手入力するダイアログを表示する
 */
function addPlannedTransaction() {
  const html = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 16px; font-size: 13px; }
        label { display: block; margin: 8px 0 4px; font-weight: bold; color: #555; font-size: 12px; }
        input, select { width: 100%; padding: 6px 8px; box-sizing: border-box;
                        font-size: 13px; border: 1px solid #ccc; border-radius: 4px; }
        .row { display: flex; gap: 8px; }
        .row > div { flex: 1; }
        .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
        .ok { background: #1a73e8; color: white; border: none; }
        .cancel { background: white; border: 1px solid #ccc; }
      </style>
    </head>
    <body>
      <h3 style="margin:0 0 12px; color:#1a73e8;">📝 入出金予定の登録</h3>

      <label>口座</label>
      <select id="account">
        <option value="CF005">PayPay 005（ビジネス営業部）</option>
        <option value="CF003">PayPay 003（はやぶさ支店）</option>
        <option value="SEIBU">西武信金（阿佐ヶ谷支店）</option>
      </select>

      <label>日付</label>
      <input type="text" id="date" placeholder="yyyy/MM/dd">

      <label>内容</label>
      <input type="text" id="content" placeholder="例: Amazon売上、家賃、JALカード">

      <div class="row">
        <div>
          <label>入金額</label>
          <input type="text" id="deposit" placeholder="0">
        </div>
        <div>
          <label>出金額</label>
          <input type="text" id="withdrawal" placeholder="0">
        </div>
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
              // フォームをリセットして続けて入力可能に
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
    </body>
    </html>
  `).setWidth(400).setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, '入出金予定の登録');
}

/**
 * 予定データをDailyシートに保存する（ダイアログから呼ばれる）
 * @param {Object} data - { account, date, content, deposit, withdrawal }
 */
function savePlannedTransaction(data) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) throw new Error('Dailyシートが見つかりません。');

  const cols = CF_CONFIG.ACCOUNTS[data.account].daily;
  const date = new Date(data.date.replace(/\//g, '-'));
  if (isNaN(date.getTime())) throw new Error('日付の形式が正しくありません。');

  const deposit = parseInt(String(data.deposit).replace(/[,、]/g, '')) || 0;
  const withdrawal = parseInt(String(data.withdrawal).replace(/[,、]/g, '')) || 0;

  if (deposit === 0 && withdrawal === 0) throw new Error('入金額または出金額を入力してください。');

  // 日付順の正しい位置に挿入
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const insertRow = findInsertRowForDate_(sheet, cols.DATE, date, headerRows);
  sheet.insertRowAfter(Math.max(insertRow - 1, headerRows));

  sheet.getRange(insertRow, cols.DATE).setValue(date).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(insertRow, cols.CONTENT).setValue(data.content);
  if (deposit > 0) sheet.getRange(insertRow, cols.DEPOSIT).setValue(deposit).setNumberFormat('#,##0');
  if (withdrawal > 0) sheet.getRange(insertRow, cols.WITHDRAWAL).setValue(withdrawal).setNumberFormat('#,##0');
  sheet.getRange(insertRow, cols.SOURCE).setValue(CF_CONFIG.SOURCE.PLANNED);

  // 背景色を予定用に変更（薄い黄色）
  sheet.getRange(insertRow, cols.DATE, 1, cols.SOURCE - cols.DATE + 1)
    .setBackground('#fff9c4');

  // 残高再計算
  recalculateBalances(data.account);
  updateDailyTotals_();
}
