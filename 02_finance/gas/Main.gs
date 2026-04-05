/**
 * 財務レポート自動化システム - メインエントリーポイント v1.2
 * MF会計 API連携 / CSVインポート PLレポート自動生成
 *
 * ■ 定期実行トリガー設定
 *   Setup.gs の setupTriggers() を1度だけ実行してください
 *   - dailyPLUpdate() : 毎日 AM6:00 に自動実行
 *     → 当月を含む過去月のデータのみ更新（未来月は手入力予測値を保持）
 */

// ==============================
// 定期実行メイン処理
// ==============================

/**
 * 【毎日 AM6:00 自動実行】
 * 当期の過去・当月分PLデータを最新データで更新する
 * ※ 未来月のセルは削除しない（手入力の予測値を保持）
 */
function dailyPLUpdate() {
  try {
    Logger.log('=== 日次PLレポート更新開始 ===');
    const fiscalYear = getCurrentFiscalYear();
    Logger.log(`対象事業年度: ${getFiscalPeriodLabel(fiscalYear)} (${fiscalYear})`);

    const months          = getFiscalMonths(fiscalYear);
    const monthsToUpdate  = PLFormatter.getUpdatableMonths(months);
    Logger.log(`更新対象月: ${monthsToUpdate.map(m => m.label).join(', ')}`);

    _runPLUpdate(fiscalYear, monthsToUpdate);

    Logger.log('=== 日次PLレポート更新完了 ===');
  } catch (e) {
    Logger.log(`エラー: ${e.message}\n${e.stack}`);
    _notifyError('dailyPLUpdate', e);
  }
}

// ==============================
// 手動実行関数（メニュー）
// ==============================

/**
 * 当期のPLレポートを手動更新する（過去・当月のみ）
 */
function runCurrentYearUpdate() {
  const ui         = SpreadsheetApp.getUi();
  const fiscalYear = getCurrentFiscalYear();
  const label      = getFiscalPeriodLabel(fiscalYear);
  const months     = getFiscalMonths(fiscalYear);
  const monthsToUpdate = PLFormatter.getUpdatableMonths(months);

  const confirm = ui.alert(
    `${label} PLレポート更新`,
    `${label}（${fiscalYear}年3月〜${fiscalYear+1}年2月）の\nPLレポートを更新します。\n\n` +
    `更新対象: ${monthsToUpdate.map(m => m.label).join('、')}（${monthsToUpdate.length}ヶ月）\n` +
    `未来月のセルは保持されます。\n\n約2〜3分かかります。よろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  try {
    _runPLUpdate(fiscalYear, monthsToUpdate);
    ui.alert(`✅ ${label} PLレポートの更新が完了しました。\n（${monthsToUpdate.length}ヶ月分を更新）`);
  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n${e.message}`);
    _notifyError('runCurrentYearUpdate', e);
  }
}

/**
 * 指定年度のPLレポートを全月分フル更新する（年度指定ダイアログ）
 * ※ 過去年度の確定値入力や全月更新が必要な場合に使用
 */
function runSpecificYearUpdate() {
  const ui   = SpreadsheetApp.getUi();
  const curr = getCurrentFiscalYear();

  const result = ui.prompt(
    '年度指定（全月フル更新）',
    `更新する事業年度の開始年を入力してください。\n例: ${curr} → ${getFiscalPeriodLabel(curr)}（${curr}/3〜${curr+1}/2）\n\n※ 全12ヶ月のデータを上書きします。`,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const fiscalYear = parseInt(result.getResponseText()) || curr;
  const months     = getFiscalMonths(fiscalYear);

  try {
    _runPLUpdate(fiscalYear, months);
    ui.alert(`✅ ${getFiscalPeriodLabel(fiscalYear)} PLレポートの全月更新が完了しました。`);
  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n${e.message}`);
    _notifyError('runSpecificYearUpdate', e);
  }
}

/**
 * 特定部門のPLのみ更新する（部分更新用）
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

  const deptName        = result.getResponseText().trim();
  const fiscalYear      = getCurrentFiscalYear();
  const months          = getFiscalMonths(fiscalYear);
  const monthsToUpdate  = PLFormatter.getUpdatableMonths(months);

  if (!depts.includes(deptName)) {
    ui.alert(`⚠️ "${deptName}" は有効な部門名ではありません。`);
    return;
  }

  try {
    Logger.log(`部門別更新開始: ${deptName}`);
    const dept        = CONFIG.DEPARTMENTS.find(d => d.name === deptName);
    const monthlyRows = {};

    monthsToUpdate.forEach(m => {
      const items = MFApiClient.getTrialBalance(m.startDate, m.endDate, dept ? dept.id : undefined);
      monthlyRows[m.label] = PLFormatter.buildPLRows(items);
      Utilities.sleep(500);
    });

    SheetManager.writePLSheet(fiscalYear, deptName, monthlyRows, monthsToUpdate);
    ui.alert(`✅ ${deptName} のPLシートを更新しました。\n（${monthsToUpdate.length}ヶ月分）`);
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

/**
 * 通期比較シートを作成する（第1期〜現在）
 */
function runPeriodComparison() {
  const ui        = SpreadsheetApp.getUi();
  const BASE_YEAR = 2018;
  const currYear  = getCurrentFiscalYear();
  const fiscalYears = [];
  for (let y = BASE_YEAR; y <= currYear; y++) fiscalYears.push(y);

  try {
    SheetManager.writePeriodComparisonSheet(fiscalYears);
    ui.alert(
      `✅ 通期比較シートを作成しました。\n` +
      `（${getFiscalPeriodLabel(BASE_YEAR)}〜${getFiscalPeriodLabel(currYear)}　${fiscalYears.length}期分）\n\n` +
      `シート名: 通期比較_全体`
    );
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

/**
 * 部門別CSVインポートを実行する
 */
function runCSVImport() {
  CSVImporter.importAllFromDrive();
}

// ==============================
// メニュー設定
// ==============================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 財務レポート')
    .addItem('🔄 当期PLを更新（過去・当月のみ）', 'runCurrentYearUpdate')
    .addItem('📅 年度指定して全月フル更新', 'runSpecificYearUpdate')
    .addItem('🏢 部門指定して更新', 'runSingleDeptUpdate')
    .addSeparator()
    .addItem('📥 部門別CSVをインポート（MF会計 推移試算表）', 'runCSVImport')
    .addItem('📊 通期比較シートを作成（第1期〜現在）', 'runPeriodComparison')
    .addSeparator()
    .addItem('🔑 MF会計 認証を開始', 'authorize')
    .addItem('📋 Web App URL を確認（リダイレクトURI用）', 'showWebAppUrl')
    .addItem('🗑️ 認証情報をリセット', 'clearAuth')
    .addSeparator()
    .addItem('⚙️ 日次トリガーを設定', 'setupTriggers')
    .addItem('⏸ トリガーを削除', 'removeTriggers')
    .addToUi();
}

// ==============================
// 内部処理
// ==============================

/**
 * PL全体更新の共通処理
 * @param {number} fiscalYear
 * @param {Array}  monthsToUpdate - 更新する月リスト（getFiscalMonths()の部分集合）
 */
function _runPLUpdate(fiscalYear, monthsToUpdate) {
  const periodLabel = getFiscalPeriodLabel(fiscalYear);
  Logger.log(`${periodLabel} データ取得開始（${monthsToUpdate.length}ヶ月）`);

  const allDeptData = PLFormatter.fetchAllDepartmentsPL(fiscalYear, monthsToUpdate);

  CONFIG.DEPARTMENTS.forEach(dept => {
    Logger.log(`${dept.name} PLシート書き込み中...`);
    SheetManager.writePLSheet(fiscalYear, dept.name, allDeptData[dept.name] || {}, monthsToUpdate);
  });

  Logger.log('全体 PLシート書き込み中...');
  SheetManager.writePLSheet(fiscalYear, '全体', allDeptData['全体'] || {}, monthsToUpdate);

  Logger.log('部門別推移表書き込み中...');
  SheetManager.writeTrendByDeptSheet(fiscalYear, allDeptData, monthsToUpdate);

  Logger.log('全体推移表書き込み中...');
  SheetManager.writeTrendConsolidatedSheet(fiscalYear, allDeptData, monthsToUpdate);

  Logger.log(`${periodLabel} 全シート更新完了`);
}

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
