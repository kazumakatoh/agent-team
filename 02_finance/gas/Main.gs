/**
 * 財務レポート自動化システム - メインエントリーポイント v1.1
 * MF会計 API連携 PLレポート自動生成
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
// 手動実行関数（メニュー・直接実行）
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
    _runPLUpdate(fiscalYear, months); // 全月を渡す
    ui.alert(`✅ ${getFiscalPeriodLabel(fiscalYear)} PLレポートの全月更新が完了しました。`);
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
    const dept       = CONFIG.DEPARTMENTS.find(d => d.name === deptName);
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
 * MF会計の勘定科目一覧を確認用シートに出力する
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
      item.id || (item.account_item && item.account_item.id) || '',
      item.name || (item.account_item && item.account_item.name) || '',
      item.account_category || (item.account_item && item.account_item.account_category) || '',
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
 * API生レスポンスを確認用シートに出力する（初回セットアップ時の診断用）
 * 実際のレスポンス構造を確認して Config.gs・PLFormatter.gs を調整してください
 */
function exportRawApiResponse() {
  const ui = SpreadsheetApp.getUi();
  try {
    const fiscalYear = getCurrentFiscalYear();
    const months     = getFiscalMonths(fiscalYear);
    const thisMonth  = PLFormatter.getUpdatableMonths(months).slice(-1)[0]; // 直近月

    if (!thisMonth) {
      ui.alert('⚠️ 取得対象の月がありません。');
      return;
    }

    // 全社（部門なし）で最新月のデータを取得
    const items = MFApiClient.getTrialBalance(thisMonth.startDate, thisMonth.endDate);

    const ss = SheetManager.getSpreadsheet();
    let sheet = ss.getSheetByName('_API生データ確認');
    if (!sheet) sheet = ss.insertSheet('_API生データ確認');
    sheet.clearContents();

    sheet.getRange(1, 1).setValue(
      `MF会計 試算表API 生データ（${thisMonth.label}・全社）\n` +
      `取得日時: ${new Date().toLocaleString('ja-JP')}\n` +
      `取得件数: ${items.length}件\n\n` +
      '↓ 下の表でフィールド名・値を確認して Config.gs / PLFormatter.gs を調整してください'
    );

    // JSON全体を1セルに出力（最初の3件）
    const sample = items.slice(0, 3);
    sheet.getRange(6, 1).setValue('【サンプル JSON（最初の3件）】');
    sheet.getRange(7, 1).setValue(JSON.stringify(sample, null, 2));
    sheet.getRange(7, 1).setWrap(true);
    sheet.setColumnWidth(1, 600);
    sheet.setRowHeight(7, 400);

    // フラット形式でも出力
    if (items.length > 0) {
      const keys = Object.keys(items[0]).concat(
        items[0].account_item ? Object.keys(items[0].account_item).map(k => `account_item.${k}`) : []
      );
      sheet.getRange(12, 1, 1, keys.length).setValues([keys]);
      const rows = items.slice(0, 20).map(item => keys.map(k => {
        if (k.startsWith('account_item.')) {
          return (item.account_item && item.account_item[k.replace('account_item.', '')]) || '';
        }
        return item[k] !== undefined ? String(item[k]) : '';
      }));
      if (rows.length) sheet.getRange(13, 1, rows.length, keys.length).setValues(rows);
    }

    ui.alert(
      `✅ API生データを "_API生データ確認" シートに出力しました。\n` +
      `${items.length}件取得 / ${thisMonth.label}\n\n` +
      `シートを確認してフィールド名を把握したら:\n` +
      `・PLFormatter.gs の _buildAccountMap を調整\n` +
      `・Config.gs の SEGMENT_PARAM を調整`
    );
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}\n\nまず「🔑 MF会計 認証を開始」から認証してください。`);
  }
}

/**
 * 部門一覧を確認用シートに出力する
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 財務レポート')
    .addItem('🔄 当期PLを更新（過去・当月のみ）', 'runCurrentYearUpdate')
    .addItem('📅 年度指定して全月フル更新', 'runSpecificYearUpdate')
    .addItem('🏢 部門指定して更新', 'runSingleDeptUpdate')
    .addSeparator()
    .addItem('📋 勘定科目一覧を出力（設定確認用）', 'exportAccountItems')
    .addItem('🏢 部門一覧を出力（設定確認用）', 'exportDepartments')
    .addItem('🔬 API生データを出力（初回診断用）', 'exportRawApiResponse')
    .addSeparator()
    .addItem('🔑 MF会計 認証を開始', 'authorize')
    .addItem('🔑 認証コードを入力', 'exchangeAuthCode')
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

  // 全部門のデータをまとめて取得（更新対象月のみ）
  const allDeptData = PLFormatter.fetchAllDepartmentsPL(fiscalYear, monthsToUpdate);

  // ── 部門別PLシート ──
  CONFIG.DEPARTMENTS.forEach(dept => {
    Logger.log(`${dept.name} PLシート書き込み中...`);
    SheetManager.writePLSheet(fiscalYear, dept.name, allDeptData[dept.name] || {}, monthsToUpdate);
  });

  // ── 全体（統合）PLシート ──
  Logger.log('全体 PLシート書き込み中...');
  SheetManager.writePLSheet(fiscalYear, '全体', allDeptData['全体'] || {}, monthsToUpdate);

  // ── 部門別推移表 ──
  Logger.log('部門別推移表書き込み中...');
  SheetManager.writeTrendByDeptSheet(fiscalYear, allDeptData, monthsToUpdate);

  // ── 全体推移表 ──
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
