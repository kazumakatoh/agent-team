/**
 * 財務レポート自動化システム - メインエントリーポイント v1.0
 * MF会計 API連携 PLレポート自動生成
 *
 * ■ 定期実行トリガー設定
 *   Setup.gs の setupTriggers() を1度だけ実行してください
 *   - monthlyPLUpdate() : 毎月1日 AM6:00 に自動実行
 */

// ==============================
// 定期実行メイン処理
// ==============================

/**
 * 【毎月1日 AM6:00 自動実行】
 * 当期のPLレポートを最新データで更新する
 */
function monthlyPLUpdate() {
  try {
    Logger.log('=== 月次PLレポート更新開始 ===');
    const fiscalYear = getCurrentFiscalYear();
    Logger.log(`対象事業年度: ${getFiscalPeriodLabel(fiscalYear)} (${fiscalYear})`);

    _runPLUpdate(fiscalYear);

    Logger.log('=== 月次PLレポート更新完了 ===');
  } catch (e) {
    Logger.log(`エラー: ${e.message}\n${e.stack}`);
    _notifyError('monthlyPLUpdate', e);
  }
}

// ==============================
// 手動実行関数（メニュー・直接実行）
// ==============================

/**
 * 当期のPLレポートを手動更新する
 */
function runCurrentYearUpdate() {
  const ui         = SpreadsheetApp.getUi();
  const fiscalYear = getCurrentFiscalYear();
  const label      = getFiscalPeriodLabel(fiscalYear);

  const confirm = ui.alert(
    `${label} PLレポート更新`,
    `${label}（${fiscalYear}年3月〜${fiscalYear+1}年2月）の\nPLレポートを更新します。\n\n約2〜3分かかります。よろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  try {
    _runPLUpdate(fiscalYear);
    ui.alert(`✅ ${label} PLレポートの更新が完了しました。`);
  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n${e.message}`);
    _notifyError('runCurrentYearUpdate', e);
  }
}

/**
 * 指定年度のPLレポートを更新する（年度指定ダイアログ）
 */
function runSpecificYearUpdate() {
  const ui   = SpreadsheetApp.getUi();
  const curr = getCurrentFiscalYear();

  const result = ui.prompt(
    '年度指定',
    `更新する事業年度の開始年を入力してください。\n例: ${curr} → ${getFiscalPeriodLabel(curr)}（${curr}/3〜${curr+1}/2）`,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const fiscalYear = parseInt(result.getResponseText()) || curr;
  try {
    _runPLUpdate(fiscalYear);
    ui.alert(`✅ ${getFiscalPeriodLabel(fiscalYear)} PLレポートの更新が完了しました。`);
  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n${e.message}`);
    _notifyError('runSpecificYearUpdate', e);
  }
}

/**
 * 特定部門のPLのみ更新する（デバッグ・部分更新用）
 */
function runSingleDeptUpdate() {
  const ui   = SpreadsheetApp.getUi();
  const depts = [...CONFIG.DEPARTMENTS.map(d => d.name), '全体'];
  const result = ui.prompt(
    '部門指定',
    `更新する部門を選択してください。\n選択肢: ${depts.join(' / ')}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const deptName   = result.getResponseText().trim();
  const fiscalYear = getCurrentFiscalYear();

  if (!depts.includes(deptName)) {
    ui.alert(`⚠️ "${deptName}" は有効な部門名ではありません。`);
    return;
  }

  try {
    Logger.log(`部門別更新開始: ${deptName}`);
    const months = getFiscalMonths(fiscalYear);
    const dept   = CONFIG.DEPARTMENTS.find(d => d.name === deptName);
    const monthlyRows = {};

    months.forEach(m => {
      const items = MFApiClient.getTrialBalance(m.startDate, m.endDate, dept ? dept.id : undefined);
      monthlyRows[m.label] = PLFormatter.buildPLRows(items);
      Utilities.sleep(500);
    });

    SheetManager.writePLSheet(fiscalYear, deptName, monthlyRows);
    ui.alert(`✅ ${deptName} のPLシートを更新しました。`);
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

/**
 * MF会計の勘定科目一覧を確認用シートに出力する
 * 勘定科目名の照合に使用してください
 */
function exportAccountItems() {
  const ui = SpreadsheetApp.getUi();
  try {
    const items = MFApiClient.getAccountItems();
    const ss    = SheetManager.getSpreadsheet();

    let sheet = ss.getSheetByName('_勘定科目確認');
    if (!sheet) sheet = ss.insertSheet('_勘定科目確認');
    sheet.clearContents();

    sheet.getRange(1, 1, 1, 5).setValues([['ID', '科目名', 'カテゴリ', 'カテゴリ名', '税区分']]);
    const rows = items.map(item => [
      item.id,
      item.name,
      item.account_category,
      item.account_category_name || '',
      item.tax_name || '',
    ]);
    if (rows.length) sheet.getRange(2, 1, rows.length, 5).setValues(rows);

    ui.alert(`✅ 勘定科目一覧を "_勘定科目確認" シートに出力しました。\n${items.length}科目`);
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

/**
 * 部門一覧を確認用シートに出力する
 * Config.gs の DEPARTMENTS 設定に使用してください
 */
function exportDepartments() {
  const ui = SpreadsheetApp.getUi();
  try {
    const depts = MFApiClient.getDepartments();
    const ss    = SheetManager.getSpreadsheet();

    let sheet = ss.getSheetByName('_部門確認');
    if (!sheet) sheet = ss.insertSheet('_部門確認');
    sheet.clearContents();

    sheet.getRange(1, 1, 1, 3).setValues([['ID', '部門名', '親部門ID']]);
    const rows = depts.map(d => [d.id, d.name, d.parent_id || '']);
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    ui.alert(`✅ 部門一覧を "_部門確認" シートに出力しました。\n${depts.length}部門\n\nConfig.gs の DEPARTMENTS に各部門のIDを設定してください。`);
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

// ==============================
// メニュー設定
// ==============================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 財務レポート')
    .addItem('🔄 当期PLを更新', 'runCurrentYearUpdate')
    .addItem('📅 年度指定して更新', 'runSpecificYearUpdate')
    .addItem('🏢 部門指定して更新', 'runSingleDeptUpdate')
    .addSeparator()
    .addItem('📋 勘定科目一覧を出力（設定確認用）', 'exportAccountItems')
    .addItem('🏢 部門一覧を出力（設定確認用）', 'exportDepartments')
    .addSeparator()
    .addItem('🔑 MF会計 認証を開始', 'authorize')
    .addItem('🔑 認証コードを入力', 'exchangeAuthCode')
    .addItem('🗑️ 認証情報をリセット', 'clearAuth')
    .addSeparator()
    .addItem('⚙️ 月次トリガーを設定', 'setupTriggers')
    .addItem('⏸ トリガーを削除', 'removeTriggers')
    .addToUi();
}

// ==============================
// 内部処理
// ==============================

/**
 * PL全体更新の共通処理
 * @param {number} fiscalYear
 */
function _runPLUpdate(fiscalYear) {
  const periodLabel = getFiscalPeriodLabel(fiscalYear);
  Logger.log(`${periodLabel} 全部門データ取得開始`);

  // 全部門のデータをまとめて取得
  const allDeptData = PLFormatter.fetchAllDepartmentsPL(fiscalYear);

  // ── 部門別PLシート ──
  CONFIG.DEPARTMENTS.forEach(dept => {
    Logger.log(`${dept.name} PLシート書き込み中...`);
    SheetManager.writePLSheet(fiscalYear, dept.name, allDeptData[dept.name] || {});
  });

  // ── 全体（統合）PLシート ──
  Logger.log('全体 PLシート書き込み中...');
  SheetManager.writePLSheet(fiscalYear, '全体', allDeptData['全体'] || {});

  // ── 部門別推移表 ──
  Logger.log('部門別推移表書き込み中...');
  SheetManager.writeTrendByDeptSheet(fiscalYear, allDeptData);

  // ── 全体推移表 ──
  Logger.log('全体推移表書き込み中...');
  SheetManager.writeTrendConsolidatedSheet(fiscalYear, allDeptData);

  Logger.log(`${periodLabel} 全シート更新完了`);
}

/**
 * エラーを記録する
 */
function _notifyError(funcName, error) {
  try {
    const ss = SheetManager.getSpreadsheet();
    let logSheet = ss.getSheetByName('_エラーログ');
    if (!logSheet) {
      logSheet = ss.insertSheet('_エラーログ');
      logSheet.appendRow(['日時', '関数名', 'エラー', 'スタック']);
    }
    logSheet.appendRow([new Date(), funcName, error.message, error.stack || '']);
  } catch (e) {
    // ログ書き込み失敗は無視
  }
}
