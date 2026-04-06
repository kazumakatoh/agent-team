/**
 * 毎月更新ダッシュボード
 *
 * チェックボックスをクリックするだけで月次更新が完了します。
 * ① MF会計からCSVをダウンロード
 * ② Googleドライブの所定フォルダにアップロード
 * ③ チェックボックスをクリック → 自動実行
 */

const Dashboard = {
  SHEET_NAME: '_ダッシュボード',
  COL_CHECK:  2,   // B列: チェックボックス
  ROW_BTN_PL: 9,  // PLインポートボタン行
  ROW_BTN_BS: 13, // BS更新ボタン行
  ROW_STATUS: 17, // ステータス表示行

  /**
   * ダッシュボードシートを作成・初期化する
   * メニュー「ダッシュボードを作成」から実行
   */
  setup() {
    const ss    = SheetManager.getSpreadsheet();
    const sName = Dashboard.SHEET_NAME;
    let sheet   = ss.getSheetByName(sName);
    if (!sheet) sheet = ss.insertSheet(sName, 0); // 先頭タブに配置
    sheet.clearContents();
    sheet.clearFormats();

    const year  = getCurrentFiscalYear();
    const label = getFiscalPeriodLabel(year);

    // 列幅
    sheet.setColumnWidth(1, 20);
    sheet.setColumnWidth(2, 50);
    sheet.setColumnWidth(3, 380);
    sheet.setColumnWidth(4, 20);

    // ── タイトル ──────────────────────────────────
    sheet.getRange(1, 2, 1, 2).merge()
         .setValue('📊 財務レポート 毎月更新')
         .setFontSize(15).setFontWeight('bold')
         .setBackground('#1a1a2e').setFontColor('#ffffff')
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    // ── 手順説明 ──────────────────────────────────
    sheet.getRange(3, 2, 1, 2).merge()
         .setValue('毎月の手順').setFontWeight('bold').setFontSize(11)
         .setFontColor('#333333');
    sheet.setRowHeight(3, 24);

    var steps = [
      '① MF会計でCSVをダウンロード（推移試算表 → 部門選択 → 全期間）',
      '② Googleドライブの所定フォルダにアップロード',
      '③ 下のチェックボックスをクリック → 自動で実行されます',
    ];
    steps.forEach(function(txt, i) {
      sheet.getRange(4 + i, 3).setValue(txt).setFontSize(10);
      sheet.setRowHeight(4 + i, 22);
    });

    // ── PL インポートボタン ───────────────────────
    sheet.getRange(8, 2, 1, 2).merge()
         .setValue('📥  PLインポート ＋ 通期比較更新　（' + label + '）')
         .setBackground('#1565c0').setFontColor('#ffffff')
         .setFontWeight('bold').setFontSize(12)
         .setVerticalAlignment('middle');
    sheet.setRowHeight(8, 40);

    sheet.getRange(Dashboard.ROW_BTN_PL, 2).insertCheckboxes();
    sheet.getRange(Dashboard.ROW_BTN_PL, 3)
         .setValue('← ここをチェックすると実行開始')
         .setFontColor('#1565c0').setFontWeight('bold').setFontSize(11);
    sheet.setRowHeight(Dashboard.ROW_BTN_PL, 40);

    // ── BS 更新ボタン ──────────────────────────────
    sheet.getRange(11, 2, 1, 2).merge()
         .setValue('📊  BS（貸借対照表）通期比較を更新')
         .setBackground('#2e7d32').setFontColor('#ffffff')
         .setFontWeight('bold').setFontSize(12)
         .setVerticalAlignment('middle');
    sheet.setRowHeight(11, 40);

    sheet.getRange(Dashboard.ROW_BTN_BS, 2).insertCheckboxes();
    sheet.getRange(Dashboard.ROW_BTN_BS, 3)
         .setValue('← ここをチェックすると実行開始')
         .setFontColor('#2e7d32').setFontWeight('bold').setFontSize(11);
    sheet.setRowHeight(Dashboard.ROW_BTN_BS, 40);

    // ── 区切り線 ──────────────────────────────────
    sheet.getRange(15, 2, 1, 2).merge().setBackground('#dddddd');
    sheet.setRowHeight(15, 6);

    // ── ステータス ────────────────────────────────
    sheet.getRange(Dashboard.ROW_STATUS, 2, 1, 2).merge()
         .setValue('最終実行: （未実行）')
         .setFontColor('#888888').setFontSize(10);
    sheet.setRowHeight(Dashboard.ROW_STATUS, 22);

    sheet.getRange(18, 2, 1, 2).merge()
         .setValue('対象年度: ' + label + '（' + year + '年3月〜' + (year + 1) + '年2月）')
         .setFontSize(10).setFontColor('#555555');
    sheet.setRowHeight(18, 22);

    Logger.log('ダッシュボード作成完了: ' + sName);
    try {
      SpreadsheetApp.getUi().alert(
        '✅ ダッシュボードを作成しました。\n' +
        'シート「' + sName + '」のチェックボックスから操作できます。'
      );
    } catch(e) {}
  },

  /**
   * ステータス行を更新する（実行結果の表示）
   */
  updateStatus(msg) {
    try {
      const ss    = SheetManager.getSpreadsheet();
      const sheet = ss.getSheetByName(Dashboard.SHEET_NAME);
      if (!sheet) return;
      const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      sheet.getRange(Dashboard.ROW_STATUS, 2, 1, 2)
           .setValue('最終実行: ' + now + '　' + msg)
           .setFontColor(msg.indexOf('❌') >= 0 ? '#c62828' : '#1b5e20')
           .setFontWeight('bold');
    } catch(e) {
      Logger.log('updateStatus error: ' + e.message);
    }
  },
};

/**
 * セル編集トリガー（チェックボックスクリック検知）
 * ※ Google Apps Script が自動的に呼び出します（インストール不要）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== Dashboard.SHEET_NAME) return;
    if (e.range.getColumn() !== Dashboard.COL_CHECK) return;
    if (e.value !== 'TRUE') return;

    e.range.setValue(false); // すぐリセット（ボタンとして動作）

    const row = e.range.getRow();
    if (row === Dashboard.ROW_BTN_PL) {
      runCurrentYearUpdate();
    } else if (row === Dashboard.ROW_BTN_BS) {
      runBSPeriodComparison();
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err.message);
    try { Dashboard.updateStatus('❌ エラー: ' + err.message); } catch(e2) {}
  }
}
