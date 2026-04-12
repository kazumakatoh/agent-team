/**
 * Amazon Dashboard - 商品マスター管理モジュール
 *
 * キャッシュフロー管理シートから商品情報をインポート
 */

/**
 * キャッシュフロー管理シートから全商品をインポート
 * 既存の商品マスターにあるASINは上書きしない（新規のみ追加）
 */
function importProductsFromCfSheet() {
  Logger.log('===== CF管理シートから商品インポート開始 =====');

  const cfSheetId = getCredential('CF_SHEET_ID');
  const cfSpreadsheet = SpreadsheetApp.openById(cfSheetId);

  // 在庫管理シートを取得（gid=347226196 のシート）
  const sheets = cfSpreadsheet.getSheets();
  let cfSheet = null;

  // シート名 or gid で探す
  for (const s of sheets) {
    if (s.getSheetId() === 347226196) {
      cfSheet = s;
      break;
    }
  }

  if (!cfSheet) {
    // gidで見つからなければ最初のシートを使う
    cfSheet = sheets[0];
    Logger.log('⚠️ gid=347226196 が見つからないため、最初のシート「' + cfSheet.getName() + '」を使用');
  }

  Logger.log('CFシート: ' + cfSheet.getName());

  // データ範囲を取得
  const lastCol = cfSheet.getLastColumn();
  const lastRow = Math.min(cfSheet.getLastRow(), 10); // ヘッダー部分のみ読む

  if (lastCol < 5 || lastRow < 2) {
    Logger.log('❌ データが不十分です');
    return;
  }

  const headerData = cfSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ASIN行と商品名行を探す
  let asinRow = -1;
  let nameRow = -1;
  let priceRow = -1;

  for (let r = 0; r < headerData.length; r++) {
    // 全列をチェックしてASINラベルまたはASINパターンを探す
    for (let c = 0; c < Math.min(lastCol, 10); c++) {
      const cellVal = String(headerData[r][c]).trim();

      // 「ASIN」ラベルを検出（A〜J列のどこかに）
      if (cellVal.toLowerCase() === 'asin' && asinRow === -1) {
        asinRow = r;
      }

      // B0で始まるASINパターンを検出
      if (cellVal.match(/^B0[A-Z0-9]{8,}$/) && asinRow === -1) {
        asinRow = r;
      }
    }

    const firstCell = String(headerData[r][0]).trim().toLowerCase();
    const secondCell = String(headerData[r][1]).trim().toLowerCase();
    const thirdCell = String(headerData[r][2]).trim().toLowerCase();

    // 仕入単価行を検出
    if (firstCell.includes('仕入単価') || secondCell.includes('仕入単価') || thirdCell.includes('仕入単価')) {
      priceRow = r;
    }

    // 商品名行を検出（「商品名」ラベルがある行の次）
    if ((firstCell === '商品名' || secondCell === '商品名' || thirdCell === '商品名') && nameRow === -1) {
      nameRow = r;
    }
  }

  // ASIN行が見つからない場合、商品名行から推定
  if (asinRow === -1) {
    Logger.log('⚠️ ASIN行が見つかりません。商品名行から商品を取得します。');
    // 商品名が入っている行を探す（E列以降に商品名がある行）
    for (let r = 0; r < headerData.length; r++) {
      let hasProductNames = false;
      for (let c = 4; c < lastCol; c++) {
        const val = String(headerData[r][c]).trim();
        if (val.length > 1 && !val.match(/^\d/) && val !== '合計' && val !== '残高') {
          hasProductNames = true;
          break;
        }
      }
      if (hasProductNames) {
        nameRow = r;
        break;
      }
    }
  } else if (nameRow === -1) {
    // ASIN行の次の行が商品名行
    nameRow = asinRow + 1;
  }

  Logger.log('ASIN行: ' + (asinRow >= 0 ? asinRow + 1 : '未検出'));
  Logger.log('商品名行: ' + (nameRow >= 0 ? nameRow + 1 : '未検出'));

  // 商品リストを構築
  const products = [];
  const startCol = 4; // E列 (0-indexed = 4) から商品開始

  for (let c = startCol; c < lastCol; c++) {
    let asin = '';
    let name = '';

    if (asinRow >= 0) {
      asin = String(headerData[asinRow][c]).trim();
    }
    if (nameRow >= 0) {
      name = String(headerData[nameRow][c]).trim();
    }

    // 空でない商品のみ追加
    if (name && name !== '合計' && name !== '残高' && name.length > 1) {
      products.push({
        asin: asin || '',
        name: name,
      });
    }
  }

  Logger.log('CFシートの商品数: ' + products.length);

  // 既存の商品マスターを取得
  const existingMap = getProductMasterMap();
  const existingAsins = Object.keys(existingMap);
  const existingNames = Object.values(existingMap).map(v => v.name);

  // 新規商品のみフィルタ
  const newProducts = products.filter(p => {
    // ASINがあればASINで重複チェック
    if (p.asin && existingAsins.includes(p.asin)) return false;
    // ASINがなければ商品名で重複チェック
    if (!p.asin && existingNames.includes(p.name)) return false;
    return true;
  });

  Logger.log('新規追加対象: ' + newProducts.length + ' 件');

  if (newProducts.length === 0) {
    Logger.log('✅ 新規商品はありません。');
    return;
  }

  // 商品マスターに追加
  const rows = newProducts.map(p => [
    p.asin,                    // ASIN
    p.name,                    // 商品名
    '',                        // カテゴリ（後で手入力）
    'アクティブ',               // ステータス
    '',                        // 仕入単価
    '',                        // 仕入れ先
    'CFシートからインポート',    // 備考
  ]);

  appendRows(SHEET_NAMES.M1_PRODUCT_MASTER, rows);

  Logger.log('✅ 商品マスター: ' + newProducts.length + ' 件追加完了');

  // 追加した商品一覧を表示
  newProducts.forEach(p => {
    Logger.log('  + ' + (p.asin || '(ASIN未設定)') + ' : ' + p.name);
  });

  Logger.log('===== インポート完了 =====');
}
