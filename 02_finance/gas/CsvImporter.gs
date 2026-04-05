/**
 * キャッシュフロー管理システム - CSV取込モジュール
 *
 * 銀行からダウンロードしたCSVファイルを読み込み、
 * Dailyシートに入出金データを反映する。
 *
 * 対応銀行:
 *  - PayPay銀行（ビジネス営業部 / はやぶさ支店）
 *  - 西武信用金庫 阿佐ヶ谷支店
 */

/**
 * CSVファイルをアップロードするHTMLダイアログを表示
 * @param {string} accountKey - 口座キー（CF005 / CF003 / SEIBU）
 */
function showCsvUploadDialog(accountKey) {
  const account = CF_CONFIG.ACCOUNTS[accountKey];
  if (!account) {
    SpreadsheetApp.getUi().alert('❌ 口座が見つかりません: ' + accountKey);
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; font-size: 13px; }
        h3 { margin: 0 0 12px; color: #1a73e8; }
        .info { color: #666; margin-bottom: 16px; font-size: 12px; }
        .file-area { border: 2px dashed #ccc; border-radius: 8px; padding: 24px;
                     text-align: center; margin-bottom: 16px; cursor: pointer; }
        .file-area:hover { border-color: #1a73e8; background: #f8f9ff; }
        .file-area.has-file { border-color: #34a853; background: #f0fff4; }
        input[type="file"] { display: none; }
        .file-name { margin-top: 8px; font-weight: bold; color: #333; }
        .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
        .ok { background: #1a73e8; color: white; border: none; }
        .ok:disabled { opacity: 0.5; cursor: not-allowed; }
        .cancel { background: white; border: 1px solid #ccc; }
        #status { display: none; margin-top: 12px; color: #1a73e8; font-size: 12px; }
        .option-row { margin-bottom: 8px; }
        label { font-size: 12px; color: #555; }
        select { padding: 4px 8px; font-size: 13px; }
      </style>
    </head>
    <body>
      <h3>📄 ${account.name} CSV取込</h3>
      <div class="info">
        銀行からダウンロードしたCSVファイルを選択してください。<br>
        既にDailyシートにある日付のデータは上書きされます。
      </div>

      <div class="option-row">
        <label>文字コード: </label>
        <select id="encoding">
          <option value="Shift_JIS" selected>Shift_JIS（デフォルト）</option>
          <option value="UTF-8">UTF-8</option>
        </select>
      </div>

      <div class="file-area" id="dropArea" onclick="document.getElementById('csvFile').click()">
        📁 クリックしてCSVファイルを選択<br>
        <span style="font-size:11px; color:#999;">またはドラッグ＆ドロップ</span>
        <div class="file-name" id="fileName"></div>
      </div>
      <input type="file" id="csvFile" accept=".csv,.txt">

      <div class="buttons">
        <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
        <button class="ok" id="okBtn" onclick="upload()" disabled>取込開始</button>
      </div>
      <div id="status">⏳ CSV取込中...</div>

      <script>
        var fileContent = null;
        var accountKey = '${accountKey}';

        document.getElementById('csvFile').addEventListener('change', function(e) {
          handleFile(e.target.files[0]);
        });

        // ドラッグ＆ドロップ
        var dropArea = document.getElementById('dropArea');
        dropArea.addEventListener('dragover', function(e) { e.preventDefault(); e.stopPropagation(); });
        dropArea.addEventListener('drop', function(e) {
          e.preventDefault(); e.stopPropagation();
          if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
        });

        function handleFile(file) {
          if (!file) return;
          document.getElementById('fileName').textContent = file.name;
          dropArea.classList.add('has-file');
          var encoding = document.getElementById('encoding').value;
          var reader = new FileReader();
          reader.onload = function(e) {
            fileContent = e.target.result;
            document.getElementById('okBtn').disabled = false;
          };
          if (encoding === 'Shift_JIS') {
            reader.readAsText(file, 'Shift_JIS');
          } else {
            reader.readAsText(file, 'UTF-8');
          }
        }

        function upload() {
          if (!fileContent) return;
          document.getElementById('okBtn').disabled = true;
          document.getElementById('status').style.display = 'block';
          google.script.run
            .withSuccessHandler(function(result) {
              google.script.host.close();
            })
            .withFailureHandler(function(e) {
              document.getElementById('status').textContent = '❌ エラー: ' + e.message;
              document.getElementById('okBtn').disabled = false;
            })
            .processCsvUpload(accountKey, fileContent);
        }
      </script>
    </body>
    </html>
  `).setWidth(480).setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, `CSV取込 - ${account.shortName}`);
}

/**
 * CSVアップロードの実処理（HTMLダイアログから呼ばれる）
 * @param {string} accountKey - 口座キー
 * @param {string} csvContent - CSVファイルの内容
 * @return {string} 結果メッセージ
 */
function processCsvUpload(accountKey, csvContent) {
  const formatKey = CF_CONFIG.ACCOUNT_CSV_MAP[accountKey];
  const format = CF_CONFIG.CSV_FORMATS[formatKey];
  if (!format) throw new Error('CSV形式が未定義: ' + formatKey);

  // CSVをパース
  const transactions = parseCsv_(csvContent, format);
  if (transactions.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 取込対象のデータが見つかりませんでした。\nCSVの形式を確認してください。');
    return;
  }

  // Dailyシートに反映
  const result = writeToDailySheet_(accountKey, transactions);

  // 残高を再計算
  recalculateBalances(accountKey);

  // アラートチェック
  checkCashOutRisk();

  SpreadsheetApp.getUi().alert(
    `✅ CSV取込完了（${CF_CONFIG.ACCOUNTS[accountKey].shortName}）\n\n` +
    `・取込件数: ${transactions.length}件\n` +
    `・新規追加: ${result.added}件\n` +
    `・上書更新: ${result.updated}件\n` +
    `・期間: ${result.dateRange}`
  );
}

/**
 * CSV文字列をパースしてトランザクション配列に変換
 * @param {string} content - CSV内容
 * @param {Object} format - CSV形式定義
 * @return {Array<Object>} トランザクション配列
 */
function parseCsv_(content, format) {
  const lines = content.split(/\r?\n/).filter(line => line.trim() !== '');

  // ヘッダー行のスキップ
  const startIdx = format.hasHeader ? 1 : format.skipRows;
  const transactions = [];

  for (let i = startIdx; i < lines.length; i++) {
    const cols = parseCsvLine_(lines[i], format.delimiter);
    const c = format.columns;

    // 日付の解析
    const dateStr = (cols[c.date] || '').trim();
    if (!dateStr) continue;

    const date = parseDateString_(dateStr);
    if (!date) continue;

    // 金額の解析（カンマ・円記号を除去）
    const deposit    = parseAmount_(cols[c.deposit]);
    const withdrawal = parseAmount_(cols[c.withdrawal]);
    const balance    = parseAmount_(cols[c.balance]);
    const content_   = (cols[c.content] || '').trim();

    // 入金も出金も0の行はスキップ
    if (deposit === 0 && withdrawal === 0) continue;

    transactions.push({
      date: date,
      content: content_,
      deposit: deposit,
      withdrawal: withdrawal,
      balance: balance,
      source: CF_CONFIG.SOURCE.CSV
    });
  }

  // 日付の昇順でソート
  transactions.sort((a, b) => a.date - b.date);

  return transactions;
}

/**
 * CSV行をパース（ダブルクォート内のカンマに対応）
 */
function parseCsvLine_(line, delimiter) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

/**
 * 日付文字列をDateオブジェクトに変換
 */
function parseDateString_(str) {
  // yyyy/MM/dd or yyyy-MM-dd or yyyyMMdd
  const cleaned = str.replace(/[年月]/g, '/').replace(/日/g, '');

  let match = cleaned.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }

  match = cleaned.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }

  return null;
}

/**
 * 金額文字列を数値に変換
 */
function parseAmount_(str) {
  if (!str) return 0;
  const cleaned = String(str).replace(/[¥￥,、\s]/g, '').trim();
  if (cleaned === '' || cleaned === '-') return 0;
  const num = parseInt(cleaned, 10);
  return isNaN(num) ? 0 : Math.abs(num);
}

/**
 * CSVデータをDailyシートに書き込む
 * @param {string} accountKey - 口座キー
 * @param {Array<Object>} transactions - トランザクション配列
 * @return {Object} { added, updated, dateRange }
 */
function writeToDailySheet_(accountKey, transactions) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) throw new Error('Dailyシートが見つかりません。Setupを実行してください。');

  const account = CF_CONFIG.ACCOUNTS[accountKey];
  const cols = account.daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = Math.max(sheet.getLastRow(), headerRows);

  // 既存データの日付マップを構築（日付 → 行番号）
  const existingDateRows = {};
  if (lastRow > headerRows) {
    const dates = sheet.getRange(headerRows + 1, cols.DATE, lastRow - headerRows, 1).getValues();
    dates.forEach((row, i) => {
      if (row[0] instanceof Date) {
        const key = Utilities.formatDate(row[0], CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');
        // 同じ日付の行が複数ある場合は最初の行を使用
        if (!existingDateRows[key]) {
          existingDateRows[key] = i + headerRows + 1;
        }
      }
    });
  }

  let added = 0;
  let updated = 0;
  let minDate = null;
  let maxDate = null;

  transactions.forEach(tx => {
    const dateKey = Utilities.formatDate(tx.date, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

    // 日付範囲の追跡
    if (!minDate || tx.date < minDate) minDate = tx.date;
    if (!maxDate || tx.date > maxDate) maxDate = tx.date;

    // 既存行を検索（同一日付・同一口座でCSVデータがあるか）
    const existingRow = findExistingTransaction_(sheet, cols, tx.date, tx.content);

    if (existingRow > 0) {
      // 既存行を上書き
      sheet.getRange(existingRow, cols.CONTENT).setValue(tx.content);
      sheet.getRange(existingRow, cols.DEPOSIT).setValue(tx.deposit || '');
      sheet.getRange(existingRow, cols.WITHDRAWAL).setValue(tx.withdrawal || '');
      sheet.getRange(existingRow, cols.SOURCE).setValue(CF_CONFIG.SOURCE.CSV);
      updated++;
    } else {
      // 新しい行を挿入（日付順に正しい位置へ）
      const insertRow = findInsertRow_(sheet, cols.DATE, tx.date, headerRows);
      if (insertRow <= lastRow + added + 1) {
        sheet.insertRowAfter(insertRow - 1);
      }

      sheet.getRange(insertRow, cols.DATE).setValue(tx.date);
      sheet.getRange(insertRow, cols.CONTENT).setValue(tx.content);
      if (tx.deposit > 0) sheet.getRange(insertRow, cols.DEPOSIT).setValue(tx.deposit);
      if (tx.withdrawal > 0) sheet.getRange(insertRow, cols.WITHDRAWAL).setValue(tx.withdrawal);
      sheet.getRange(insertRow, cols.SOURCE).setValue(CF_CONFIG.SOURCE.CSV);
      added++;
    }
  });

  const dateRange = minDate && maxDate
    ? `${Utilities.formatDate(minDate, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd')} 〜 ${Utilities.formatDate(maxDate, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd')}`
    : '不明';

  return { added, updated, dateRange };
}

/**
 * 同一日付・同一内容のトランザクションが既にあるか検索
 * @return {number} 行番号（見つからない場合0）
 */
function findExistingTransaction_(sheet, cols, date, content) {
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return 0;

  const dateCol = sheet.getRange(headerRows + 1, cols.DATE, lastRow - headerRows, 1).getValues();
  const contentCol = sheet.getRange(headerRows + 1, cols.CONTENT, lastRow - headerRows, 1).getValues();
  const sourceCol = sheet.getRange(headerRows + 1, cols.SOURCE, lastRow - headerRows, 1).getValues();

  const targetDateKey = Utilities.formatDate(date, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

  for (let i = 0; i < dateCol.length; i++) {
    if (!(dateCol[i][0] instanceof Date)) continue;
    const rowDateKey = Utilities.formatDate(dateCol[i][0], CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

    if (rowDateKey === targetDateKey
        && String(contentCol[i][0]).trim() === String(content).trim()
        && sourceCol[i][0] === CF_CONFIG.SOURCE.CSV) {
      return i + headerRows + 1;
    }
  }

  return 0;
}

/**
 * 日付順で正しい挿入位置を見つける
 * @return {number} 挿入すべき行番号
 */
function findInsertRow_(sheet, dateCol, targetDate, headerRows) {
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
// CSV取込のショートカット関数
// ==============================

function importCsvCF005() { showCsvUploadDialog('CF005'); }
function importCsvCF003() { showCsvUploadDialog('CF003'); }
function importCsvSeibu() { showCsvUploadDialog('SEIBU'); }
