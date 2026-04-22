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
 * 日付が逆順の場合は自動で入れ替える
 * 会計年度をまたぐ場合は自動で分割して取得する
 */
function syncDateRangeToDaily(dateFrom, dateTo) {
  // 日付順序チェック（逆順なら入れ替え）
  if (dateFrom > dateTo) {
    const tmp = dateFrom;
    dateFrom = dateTo;
    dateTo = tmp;
  }

  // 会計年度をまたぐかチェック（3/1が年度境界）
  const periods = splitByFiscalYear_(dateFrom, dateTo);

  let totalInserted = 0;
  periods.forEach(p => {
    Logger.log(`同期: ${p.from} 〜 ${p.to}`);
    syncToDaily_(p.from, p.to);
  });

  const msg = periods.length > 1
    ? `✅ MFデータ同期完了\n\n${periods.length}期間に分けて取得:\n` + periods.map(p => `・${p.from} 〜 ${p.to}`).join('\n')
    : `✅ MFデータ同期完了\n\n・期間: ${dateFrom} 〜 ${dateTo}`;

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 期間を会計年度（3/1〜翌2/末）ごとに分割する
 * 例: 2026-01-01 〜 2026-04-30
 *     → [{from: 2026-01-01, to: 2026-02-28}, {from: 2026-03-01, to: 2026-04-30}]
 */
function splitByFiscalYear_(dateFrom, dateTo) {
  const periods = [];
  const start = new Date(dateFrom + 'T00:00:00');
  const end = new Date(dateTo + 'T00:00:00');

  let current = new Date(start);

  while (current <= end) {
    const y = current.getFullYear();
    const m = current.getMonth() + 1;

    // 今の日付が属する会計年度の最終日（2月末）を計算
    // 3月〜12月 → 翌年2月末が年度末
    // 1月〜2月 → 今年2月末が年度末
    const fyEndYear = m < 3 ? y : y + 1;
    // 3月1日から1日引いて2月末日を取得（閏年も自動対応）
    const fyEnd = new Date(fyEndYear, 2, 1);  // 3月1日
    fyEnd.setDate(fyEnd.getDate() - 1);       // → 2月末日

    // 今期間の終了日 = min(fyEnd, end)
    const periodEnd = fyEnd < end ? fyEnd : end;

    periods.push({
      from: Utilities.formatDate(current, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd'),
      to: Utilities.formatDate(periodEnd, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd')
    });

    // 次の期間の開始 = periodEnd + 1日
    current = new Date(periodEnd);
    current.setDate(current.getDate() + 1);
  }

  return periods;
}

/**
 * MFデータ同期の共通処理
 * 期間内の既存「MF」ソース行を削除してから新しいデータを挿入する
 * （手入力・予定は削除されない）
 */
function syncToDaily_(dateFrom, dateTo) {
  if (!isMfConnected()) {
    throw new Error('MF未連携です。先にMF連携を実行してください。');
  }

  Logger.log(`=== Daily同期開始: ${dateFrom} 〜 ${dateTo} ===`);

  const ss = getCfSpreadsheet();
  const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);

  // 期間の Date オブジェクトを準備
  const periodStart = new Date(dateFrom + 'T00:00:00');
  const periodEnd = new Date(dateTo + 'T23:59:59');

  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([accountKey, account]) => {
    const sheet = ss.getSheetByName(account.dailySheet);
    if (!sheet) {
      Logger.log(`⚠️ シート ${account.dailySheet} が見つかりません`);
      return;
    }

    // 期間内の既存MF行を削除（ダブり防止）
    deleteMfRowsInPeriod_(sheet, periodStart, periodEnd);

    // 新しいMFデータを挿入
    const txns = allTxns[accountKey] || [];
    if (txns.length > 0) {
      writeTransactionsToSheet_(sheet, txns);
    }

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
 * 指定期間内の「MF」ソース行を削除する（ダブり防止）
 */
function deleteMfRowsInPeriod_(sheet, dateFrom, dateTo) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  // 下から走査して削除（行番号のずれ防止）
  for (let row = lastRow; row > headerRows; row--) {
    const date = sheet.getRange(row, C.DATE).getValue();
    const source = sheet.getRange(row, C.SOURCE).getValue();

    if (source !== CF_CONFIG.SOURCE.MF) continue;
    if (!(date instanceof Date)) continue;

    if (date >= dateFrom && date <= dateTo) {
      sheet.deleteRow(row);
    }
  }
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
// 予定マスタからの一括展開
// ==============================

/**
 * 予定マスタから定期予定を各Dailyシートに展開する
 *
 * 処理:
 *  1. 各Dailyシートから既存の「予定」ソースの行を削除（実績は残す）
 *  2. 予定マスタを読み込む
 *  3. 各予定を期間内の発生日に展開してDailyシートに挿入
 *  4. 残高再計算 + 日別サマリー更新
 */
function expandPlannedTransactions() {
  const ui = SpreadsheetApp.getUi();
  const ss = getCfSpreadsheet();

  const masterSheet = ss.getSheetByName('予定マスタ');
  if (!masterSheet) {
    ui.alert('❌ 予定マスタシートが見つかりません。セットアップを実行してください。');
    return;
  }

  const result = ui.alert(
    '予定を一括展開',
    '予定マスタから全Dailyシートに予定を展開します。\n\n' +
    '・既存の「予定」行は削除されます（実績は残ります）\n' +
    '・マスタの全項目を期間内の発生日に展開します\n' +
    '・残高と日別サマリーが自動更新されます\n\n' +
    '実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  // 1. 既存の予定行を削除
  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    const sheet = ss.getSheetByName(account.dailySheet);
    if (sheet) removePlannedRows_(sheet);
  });

  // 2. 予定マスタを読み込み
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ 予定マスタにデータがありません。');
    return;
  }

  const masterData = masterSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  let totalInserted = 0;

  // 3. 各予定を展開
  masterData.forEach(row => {
    const [accountKey, content, amount, type, frequency, dayStr, startYm, endYm, note] = row;

    if (!accountKey || !content || !startYm || !endYm) return;
    if (!CF_CONFIG.ACCOUNTS[accountKey]) return;

    const sheet = ss.getSheetByName(CF_CONFIG.ACCOUNTS[accountKey].dailySheet);
    if (!sheet) return;

    // 各発生日に展開
    const dates = generatePlannedDates_(String(frequency), String(dayStr), String(startYm), String(endYm));
    dates.forEach(d => {
      insertPlannedRow_(sheet, {
        date: d,
        content: String(content),
        deposit: type === '入' ? Number(amount) || 0 : 0,
        withdrawal: type === '出' ? Number(amount) || 0 : 0
      });
      totalInserted++;
    });
  });

  // 4. 残高再計算
  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    const sheet = ss.getSheetByName(account.dailySheet);
    if (sheet) recalculateBalances_(sheet);
  });

  // 5. 現残高と日別サマリー更新
  updateCurrentBalanceSheet_();
  updateDailySummary();

  ui.alert(`✅ 予定を展開しました\n\n・展開件数: ${totalInserted}件`);
}

/**
 * Dailyシートから「予定」ソースの行を全て削除する
 */
function removePlannedRows_(sheet) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  // 下から走査して削除
  for (let row = lastRow; row > headerRows + 1; row--) {
    const source = sheet.getRange(row, C.SOURCE).getValue();
    if (source === CF_CONFIG.SOURCE.PLANNED) {
      sheet.deleteRow(row);
    }
  }
}

/**
 * 頻度・発生日・期間から実際の日付リストを生成する
 * @param {string} frequency - monthly/bimonthly/yearly
 * @param {string} dayStr - 発生日（数字/last/end/MM/DD）
 * @param {string} startYm - 開始年月 (2026.01形式)
 * @param {string} endYm - 終了年月 (2027.12形式)
 * @return {Array<Date>}
 */
function generatePlannedDates_(frequency, dayStr, startYm, endYm) {
  const dates = [];
  const startMatch = String(startYm).match(/^(\d{4})[\.\/](\d{1,2})$/);
  const endMatch = String(endYm).match(/^(\d{4})[\.\/](\d{1,2})$/);
  if (!startMatch || !endMatch) return dates;

  const startY = parseInt(startMatch[1]);
  const startM = parseInt(startMatch[2]);
  const endY = parseInt(endMatch[1]);
  const endM = parseInt(endMatch[2]);

  if (frequency === 'yearly') {
    // yearly: dayStrは "MM/DD" 形式
    const ymdMatch = String(dayStr).match(/^(\d{1,2})[\/\-](\d{1,2})$/);
    if (!ymdMatch) return dates;
    const mo = parseInt(ymdMatch[1]);
    const d = parseInt(ymdMatch[2]);

    for (let y = startY; y <= endY; y++) {
      const date = new Date(y, mo - 1, d);
      if (isWithinPeriod_(date, startY, startM, endY, endM)) {
        dates.push(date);
      }
    }
    return dates;
  }

  // monthly / bimonthly
  const step = frequency === 'bimonthly' ? 2 : 1;
  let y = startY, m = startM;

  while (y < endY || (y === endY && m <= endM)) {
    const date = resolvePlannedDate_(y, m, dayStr);
    if (date) dates.push(date);

    m += step;
    while (m > 12) { m -= 12; y++; }
  }

  return dates;
}

/**
 * 年月と発生日文字列から実際の日付を決定する
 */
function resolvePlannedDate_(year, month, dayStr) {
  const s = String(dayStr).toLowerCase().trim();

  // "end" or "月末" = 月末日
  if (s === 'end' || s === '月末') {
    return new Date(year, month, 0);  // 翌月の0日 = 当月末日
  }

  // "last" = 最終営業日
  if (s === 'last' || s === '最終営業日') {
    let d = new Date(year, month, 0);  // 月末日から開始
    while (d.getDay() === 0 || d.getDay() === 6) {
      d.setDate(d.getDate() - 1);
    }
    return d;
  }

  // 数字
  const n = parseInt(s);
  if (isNaN(n) || n < 1 || n > 31) return null;

  // 月の最終日を超える場合は月末日に丸める
  const lastDay = new Date(year, month, 0).getDate();
  const day = Math.min(n, lastDay);
  return new Date(year, month - 1, day);
}

/**
 * 日付が期間内か判定
 */
function isWithinPeriod_(date, startY, startM, endY, endM) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  if (y < startY || (y === startY && m < startM)) return false;
  if (y > endY || (y === endY && m > endM)) return false;
  return true;
}

/**
 * 予定行をDailyシートに挿入する（日付順）
 */
function insertPlannedRow_(sheet, tx) {
  const C = CF_CONFIG.DAILY_COLS;
  const insertRow = findInsertRow_(sheet, tx.date);

  if (insertRow <= sheet.getLastRow()) {
    sheet.insertRowBefore(insertRow);
  }

  sheet.getRange(insertRow, C.DATE).setValue(tx.date).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(insertRow, C.CONTENT).setValue(tx.content);
  if (tx.deposit > 0) sheet.getRange(insertRow, C.DEPOSIT).setValue(tx.deposit).setNumberFormat('#,##0');
  if (tx.withdrawal > 0) sheet.getRange(insertRow, C.WITHDRAWAL).setValue(tx.withdrawal).setNumberFormat('#,##0');
  sheet.getRange(insertRow, C.SOURCE).setValue(CF_CONFIG.SOURCE.PLANNED);

  // 薄い黄色で予定行を視覚的に区別
  sheet.getRange(insertRow, 1, 1, 6).setBackground('#fff9c4');
}

// ==============================
// Dailyシートのクリーンアップ
// ==============================

/**
 * 全Dailyシートから「日付なしの行」を削除する
 * 残高の累積計算で異常値になっている時の復旧用
 */
function cleanupAllDailySheets() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Dailyシートのクリーンアップ',
    '全Dailyシート（Daily_005/Daily_003/Daily_西武）から、\n' +
    '日付が入っていない行を削除します。\n\n' +
    '※ 前月繰越行（2行目）は残ります。\n' +
    '※ この操作は元に戻せません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  const ss = getCfSpreadsheet();
  let totalDeleted = 0;

  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    const sheet = ss.getSheetByName(account.dailySheet);
    if (!sheet) return;
    const deleted = cleanupDailySheet_(sheet);
    totalDeleted += deleted;
    Logger.log(`${account.dailySheet}: ${deleted}行削除`);
  });

  // 残高再計算
  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    const sheet = ss.getSheetByName(account.dailySheet);
    if (sheet) recalculateBalances_(sheet);
  });

  // 現残高と日別サマリー更新
  updateCurrentBalanceSheet_();
  updateDailySummary();

  ui.alert(`✅ クリーンアップ完了\n\n合計 ${totalDeleted}行を削除しました。`);
}

/**
 * 1つのDailyシートから日付がない行を削除する
 * @return {number} 削除した行数
 */
function cleanupDailySheet_(sheet) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows + 1) return 0;

  // 下から走査（削除で行番号がずれるため）
  // 2行目（前月繰越）は残す
  let deleted = 0;
  for (let row = lastRow; row > headerRows + 1; row--) {
    const date = sheet.getRange(row, C.DATE).getValue();
    if (!(date instanceof Date)) {
      sheet.deleteRow(row);
      deleted++;
    }
  }
  return deleted;
}

// ==============================
// 残高計算
// ==============================

/**
 * Dailyシートの残高を数式で再構築する
 *
 * 最初の行（前月繰越）: 値のまま
 * 以降の行: =前行の残高 + 当行の入金 - 当行の出金
 *
 * これにより、金額を手動で変更した時に自動で残高が再計算される。
 *
 * 異常行の自動クリーンアップ:
 * - 日付なしの行 → 残高をクリア
 */
function recalculateBalances_(sheet) {
  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const numRows = lastRow - headerRows;
  const dates = sheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();
  const deposits = sheet.getRange(headerRows + 1, C.DEPOSIT, numRows, 1).getValues();
  const withdrawals = sheet.getRange(headerRows + 1, C.WITHDRAWAL, numRows, 1).getValues();

  // 列文字（A=1, B=2, ...）
  const depCol = 'C';   // C: 入金
  const wthCol = 'D';   // D: 出金
  const balCol = 'E';   // E: 残高

  // 直前の有効な残高セル（数式参照用）
  let prevBalanceRow = 0;

  for (let i = 0; i < numRows; i++) {
    const dep = Number(deposits[i][0]) || 0;
    const wth = Number(withdrawals[i][0]) || 0;
    const date = dates[i][0];
    const hasDate = date instanceof Date;
    const currentRow = headerRows + 1 + i;
    const balanceCell = sheet.getRange(currentRow, C.BALANCE);

    // 最初の行が前月繰越行（日付あり・入出金なし・残高あり）ならスキップ（値のまま）
    if (i === 0 && hasDate && dep === 0 && wth === 0) {
      prevBalanceRow = currentRow;
      continue;
    }

    // 日付なしの行は残高をクリア
    if (!hasDate) {
      balanceCell.setValue('');
      continue;
    }

    // 数式を設定: =前行の残高 + 入金 - 出金
    if (prevBalanceRow === 0) {
      // 前月繰越がまだ見つかっていない場合、入出金の差で開始
      balanceCell.setFormula(`=IFERROR(${depCol}${currentRow},0)-IFERROR(${wthCol}${currentRow},0)`);
    } else {
      balanceCell.setFormula(
        `=IFERROR(${balCol}${prevBalanceRow},0)+IFERROR(${depCol}${currentRow},0)-IFERROR(${wthCol}${currentRow},0)`
      );
    }
    balanceCell.setNumberFormat('#,##0');

    prevBalanceRow = currentRow;
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
// ==============================
// 在庫数の最新化
// ==============================

/**
 * 発注管理表の在庫数を在庫残高シートの当月「在庫」行に反映する
 *
 * 処理:
 *  1. 発注管理表（別スプシ）の「在庫一覧」シートからASIN→在庫数マップを作成
 *  2. 在庫残高シートの1行目（ASIN行）と照合
 *  3. 当月の「在庫」行に在庫数を書き込む
 *  4. 見込売上 = 在庫数 × 販売単価 を自動再計算
 *  5. 棚卸原価 = 在庫数 × 仕入原価 を自動再計算
 */
function updateInventoryFromOrderMgmt() {
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. 発注管理表を開く
    const orderSs = SpreadsheetApp.openById(CF_CONFIG.ORDER_MGMT.SPREADSHEET_ID);
    const orderSheet = orderSs.getSheetByName(CF_CONFIG.ORDER_MGMT.SHEET_NAME);
    if (!orderSheet) throw new Error(`発注管理表の「${CF_CONFIG.ORDER_MGMT.SHEET_NAME}」シートが見つかりません。`);

    const orderCols = CF_CONFIG.ORDER_MGMT.COLS;
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 2) throw new Error('発注管理表にデータがありません。');

    // ASIN → FBA在庫数マップを作成（F列: FBA在庫を使用）
    const asinData = orderSheet.getRange(2, orderCols.ASIN, lastRow - 1, 1).getValues();
    const stockData = orderSheet.getRange(2, orderCols.FBA_STOCK, lastRow - 1, 1).getValues();

    const stockMap = {};
    for (let i = 0; i < asinData.length; i++) {
      const asin = String(asinData[i][0]).trim();
      if (asin) {
        stockMap[asin] = (stockMap[asin] || 0) + (Number(stockData[i][0]) || 0);
      }
    }

    Logger.log(`発注管理表: ${Object.keys(stockMap).length}件のASINを取得`);

    // 2. 在庫残高シートを開く
    const ss = getCfSpreadsheet();
    const invSheet = ss.getSheetByName(CF_CONFIG.SHEETS.INVENTORY);
    if (!invSheet) throw new Error('在庫残高シートが見つかりません。');

    // 1行目のASINを読み込み（E列=5列目以降）
    const lastCol = invSheet.getLastColumn();
    if (lastCol < 5) throw new Error('在庫残高シートに商品列がありません。');

    const asinRow = invSheet.getRange(1, 5, 1, lastCol - 4).getValues()[0];

    // 3. 当月の「在庫」行を特定
    const today = new Date();
    const yearMonth = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, '0')}`;
    // 在庫残高シートの列構造: A=空, B=年月, C=項目(在庫/見込売上/...)
    const invLastRow = invSheet.getLastRow();
    const colYm = invSheet.getRange(1, 2, invLastRow, 1).getValues();   // B列: 年月
    const colItem = invSheet.getRange(1, 3, invLastRow, 1).getValues(); // C列: 項目

    // 5行構成: 在庫, 見込売上, 棚卸原価, 仕入原価, 販売単価
    // yearMonthの在庫行を見つける
    let stockRow = 0;
    const targetYm = Number(yearMonth); // 2026.04 as number
    for (let i = 0; i < invLastRow; i++) {
      const rawYm = colYm[i][0];
      const item = String(colItem[i][0]).trim();

      // 年月マッチ（テキスト/数値どちらでも対応）
      const ymStr = String(rawYm).trim();
      const ymNum = Number(rawYm);
      const ymMatch = ymStr === yearMonth || ymNum === targetYm;

      if (ymMatch && item === '在庫') {
        stockRow = i + 1;
        break;
      }
    }

    if (stockRow === 0) throw new Error(`${yearMonth}の在庫行が見つかりません。`);

    salesRow = stockRow + 1;  // 見込売上
    costRow = stockRow + 2;   // 棚卸原価
    const unitCostRow = stockRow + 3; // 仕入原価
    const unitPriceRow = stockRow + 4; // 販売単価

    // 4. ASINでマッチして在庫数を書き込み
    let matched = 0;
    let unmatched = 0;

    for (let col = 0; col < asinRow.length; col++) {
      const asin = String(asinRow[col]).trim();
      if (!asin) continue;

      const colIdx = col + 5; // E列=5から

      if (stockMap[asin] !== undefined) {
        const qty = stockMap[asin];
        invSheet.getRange(stockRow, colIdx).setValue(qty);

        // 見込売上 = 在庫数 × 販売単価
        invSheet.getRange(salesRow, colIdx)
          .setFormula(`=${invSheet.getRange(stockRow, colIdx).getA1Notation()}*${invSheet.getRange(unitPriceRow, colIdx).getA1Notation()}`);

        // 棚卸原価 = 在庫数 × 仕入原価
        invSheet.getRange(costRow, colIdx)
          .setFormula(`=${invSheet.getRange(stockRow, colIdx).getA1Notation()}*${invSheet.getRange(unitCostRow, colIdx).getA1Notation()}`);

        matched++;
      } else {
        unmatched++;
      }
    }

    // D列の合計（見込売上合計、棚卸原価合計）を更新
    const startCol = columnToLetter_(5);
    const endCol = columnToLetter_(lastCol);
    invSheet.getRange(salesRow, 4)
      .setFormula(`=SUM(${startCol}${salesRow}:${endCol}${salesRow})`);
    invSheet.getRange(costRow, 4)
      .setFormula(`=SUM(${startCol}${costRow}:${endCol}${costRow})`);

    ui.alert(
      `✅ 在庫数を最新化しました\n\n` +
      `・対象月: ${yearMonth}\n` +
      `・マッチ: ${matched}商品\n` +
      `・未マッチ: ${unmatched}商品\n\n` +
      `見込売上・棚卸原価も自動更新されました。`
    );

  } catch (e) {
    ui.alert(`❌ エラー\n\n${e.message}`);
    Logger.log(`在庫更新エラー: ${e.message}\n${e.stack}`);
  }
}

/**
 * 列番号をA1表記のアルファベットに変換
 */
function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
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
    // BS残高データ取得完了

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
