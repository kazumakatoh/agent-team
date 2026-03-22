/**
 * 民泊自動化システム - エアサポ請求書パーサー
 *
 * エアサポ（support@air-sapo.com）から届く月次請求書メールを自動解析し、
 * 経費入力シートの該当月に代行手数料・清掃費・リネン費・備品費を反映する。
 *
 * ■ 前提条件
 *   - Google Apps Script で「Drive API」Advanced Service を有効化すること
 *     (GASエディタ: サービス追加 → Drive API v3)
 *
 * ■ 請求書フォーマット（OCRテキスト解析）
 *   - タイトル: "YYYY年M月 御請求書" → 対象年月を抽出
 *   - 運業代行費合計: 代行手数料・清掃費 の合計（税込）
 *   - 物件経費合計: リネン費・備品費 の合計（税込）
 *   各明細行:
 *     "代行手数料  12,000円" / "清掃費  8,000円" / "リネン費  3,000円" / "備品・消耗品  2,000円"
 *   ※ 金額は請求書上が税抜きの場合、×1.1 で税込に変換して格納
 */

// ==============================
// 公開エントリーポイント
// ==============================

/**
 * エアサポ請求書メールを検索・解析して経費入力シートに書き込む
 * @param {Date|null} since - この日付以降のメールを対象（省略時: 過去90日）
 * @return {number} 処理した請求書の件数
 */
function fetchAirSapoInvoices(since) {
  const sinceDate = since || (() => {
    const d = new Date();
    d.setDate(d.getDate() - 90);
    return d;
  })();

  const afterStr = Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  const query = `from:support@air-sapo.com subject:請求書 after:${afterStr}`;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log('エアサポ請求書メールが見つかりませんでした');
    return 0;
  }

  // 処理済みラベルを取得または作成
  const processedLabel = getOrCreateLabel_('民泊_エアサポ処理済み');
  let processedCount = 0;

  threads.forEach(thread => {
    // すでに処理済みラベルが付いているスレッドはスキップ
    const threadLabels = thread.getLabels().map(l => l.getName());
    if (threadLabels.includes('民泊_エアサポ処理済み')) {
      return;
    }

    const messages = thread.getMessages();
    messages.forEach(msg => {
      const attachments = msg.getAttachments();
      attachments.forEach(att => {
        const name = att.getName() || '';
        // PDF添付のみ対象
        if (!name.toLowerCase().endsWith('.pdf')) return;

        try {
          const text = extractPdfText_(att);
          if (!text) {
            Logger.log(`PDF OCR失敗: ${name}`);
            return;
          }

          const parsed = parseInvoiceText_(text, name);
          if (!parsed) {
            Logger.log(`請求書解析失敗: ${name}\n本文: ${text.substring(0, 200)}`);
            return;
          }

          const written = writeCostToSheet_(parsed);
          if (written) {
            Logger.log(`経費書込完了: ${parsed.yearMonth} → 代行=${parsed.agencyFee}, 清掃=${parsed.cleaning}, リネン=${parsed.linen}, 備品=${parsed.supplies}, その他=${parsed.others}`);
            processedCount++;
          }
        } catch (e) {
          Logger.log(`エアサポ請求書処理エラー [${name}]: ${e.message}`);
        }
      });
    });

    // 処理済みラベルを付与
    thread.addLabel(processedLabel);
  });

  return processedCount;
}

// ==============================
// PDF テキスト抽出（Drive API OCR）
// ==============================

/**
 * Gmail 添付ファイルを Drive API で OCR し、テキストを返す
 * @param {GmailAttachment} attachment
 * @return {string|null}
 */
function extractPdfText_(attachment) {
  let fileId = null;
  try {
    // Blob を Drive にアップロード（Drive API v3 + OCR）
    // mimeType を Google Doc に指定することで自動OCR変換される
    const blob = attachment.copyBlob().setContentType('application/pdf');
    const resource = {
      name: `airsapo_ocr_${Date.now()}.pdf`,
      mimeType: 'application/vnd.google-apps.document'
    };
    const options = {
      ocrLanguage: 'ja',
      fields: 'id'
    };

    const file = Drive.Files.create(resource, blob, options);
    fileId = file.id;

    // Google Doc として取得してテキストを読む
    const doc = DocumentApp.openById(fileId);
    const text = doc.getBody().getText();
    return text;

  } finally {
    // 一時ファイルを削除（Drive API v3）
    if (fileId) {
      try { Drive.Files.remove(fileId); } catch (e) { /* 無視 */ }
    }
  }
}

// ==============================
// 請求書テキスト解析
// ==============================

/**
 * OCR テキストから請求情報を抽出する
 * @param {string} text - OCR結果のテキスト
 * @return {Object|null} { yearMonth, agencyFee, cleaning, linen, supplies } または null
 */
function parseInvoiceText_(text, filename) {
  // Unicode正規化（CJK互換漢字・合成文字の統一）
  text = text.normalize('NFKC');
  // 全角数字・記号を半角に正規化
  text = text.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
             .replace(/[¥￥]/g, '¥')
             .replace(/，/g, ',')
             .replace(/：/g, ':')           // 全角コロン→半角
             .replace(/[\u200b\ufeff]/g, '') // ゼロ幅スペース・BOM除去
             .replace(/\u00a0/g, ' ');       // ノーブレークスペース→スペース

  // 年月抽出: テキストから試みる → 失敗時はファイル名からフォールバック
  let ymMatch = text.match(/(\d{4})年\s*(\d{1,2})月/);
  if (!ymMatch && filename) {
    ymMatch = filename.match(/(\d{4})年(\d{1,2})月/);
    if (ymMatch) Logger.log(`年月をファイル名から取得: ${filename}`);
  }
  if (!ymMatch) {
    Logger.log('年月が見つかりません');
    return null;
  }
  const year  = parseInt(ymMatch[1]);
  const month = parseInt(ymMatch[2]);
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;

  // デバッグ: OCRテキストを分割出力（金額項目の確認用）
  for (let _i = 0; _i < Math.min(text.length, 1200); _i += 400) {
    Logger.log(`[OCR ${yearMonth} p${Math.floor(_i/400)+1}]: ${text.substring(_i, _i+400)}`);
  }
  // デバッグ: キーワード存在確認
  Logger.log(`[KW] 代行サービス料金=${text.includes('代行サービス料金')} 清掃料金=${text.includes('清掃料金')} リネン=${text.includes('リネン')}`);
  Logger.log(`[KW] beds24=${text.toLowerCase().includes('beds24')} maneKEY=${text.toLowerCase().includes('manekey')}`);

  // 金額抽出ヘルパー: 数値文字列からカンマを除去して整数に変換
  const extractAmount = (pattern) => {
    const m = text.match(pattern);
    if (!m) return 0;
    return parseInt(m[1].replace(/,/g, '')) || 0;
  };

  // 代行手数料（税込）
  // パターン例: "代行サービス料金 ¥12,000" / "代行手数料 ¥12,000" / "運営代行費 12,000円"
  const agencyFee = extractAmountWithTax_(
    text,
    /(?:代行サービス料金|代行手数料|運営代行費)[^\d]*([\d,]{3,})/,
    /(?:代行サービス料金|代行手数料|運営代行費).*?([0-9,]{3,})\s*円/
  );

  // 清掃費（税込）
  // パターン例: "清掃料金 ¥8,000" / "清掃費 8,000円"
  const cleaning = extractAmountWithTax_(
    text,
    /(?:清掃料金|清掃費)[^\d]*([\d,]{3,})/,
    /(?:清掃料金|清掃費).*?([0-9,]{3,})\s*円/
  );

  // リネン費（税込）
  // 形式A: 「リネン補償金 ¥1,000」→ リネンの後に金額（税抜き × 1.1）
  // 形式B: 「¥9,240 12月リネン回収クリーニング費用」→ 月別備考でリネン前に税込金額
  const linen = (() => {
    const mA = text.match(/リネン[^\d\n]*([\d,]{3,})/);
    if (mA) {
      const raw = parseInt(mA[1].replace(/,/g, '')) || 0;
      if (raw > 0) {
        Logger.log(`  [リネン形式A] raw=${raw} → ${Math.round(raw * 1.1)}`);
        return Math.round(raw * 1.1);
      }
    }
    // 形式B: 「9,240 12月リネン...」のように、月表記の前の数値が税込金額
    const mB = text.match(/([\d,]{3,})\s+\d+月.*?リネン/);
    if (mB) {
      const raw = parseInt(mB[1].replace(/,/g, '')) || 0;
      if (raw > 0) {
        Logger.log(`  [リネン形式B] raw=${raw} 税込とみなす → ${raw}`);
        return raw;
      }
    }
    return 0;
  })();

  // 備品・消耗品費（税込）
  // パターン例: "備品 ¥2,000" / "備品・消耗品 2,000円" / "消耗品費 2,000"
  const supplies = extractAmountWithTax_(
    text,
    /(?:備品[・・]消耗品|備品費|消耗品費|備品)[^\d]*([\d,]{3,})/,
    /(?:備品[・・]消耗品|備品費|消耗品費|備品).*?([0-9,]{3,})\s*円/
  );

  // その他経費（税込）
  // beds24 / maneKEY は請求書内に複数行ある場合があるため matchAll で全行合算
  const others = extractAllAmountsWithTax_(text, [
    /beds24[^\d]*([\d,]{3,})/ig,
    /maneKEY[^\d]*([\d,]{3,})/ig
  ]);

  // 全項目が0の場合は解析失敗とみなす
  if (agencyFee === 0 && cleaning === 0 && linen === 0 && supplies === 0 && others === 0) {
    Logger.log(`全項目0円: yearMonth=${yearMonth}, テキスト先頭=${text.substring(0, 300)}`);
    return null;
  }

  return { yearMonth, agencyFee, cleaning, linen, supplies, others };
}

/**
 * テキストから金額を抽出し、必要に応じて消費税(10%)を加算して返す
 * 請求書に "税込" という記載があれば税込金額として、なければ税抜きとして×1.1
 * @param {string} text
 * @param {...RegExp} patterns - 試みる正規表現（先に一致したものを使用）
 * @return {number}
 */
function extractAmountWithTax_(text, ...patterns) {
  for (const pat of patterns) {
    const m = text.match(pat);
    if (m) {
      const raw = parseInt(m[1].replace(/,/g, '')) || 0;
      if (raw === 0) continue;

      // 金額直後に "税込" の記載があるか確認（前後40文字）
      const idx = text.indexOf(m[0]);
      const context = text.substring(Math.max(0, idx - 10), idx + m[0].length + 40);
      const isTaxIncluded = /税込/.test(context);

      const result = isTaxIncluded ? raw : Math.round(raw * 1.1);
      Logger.log(`  [金額抽出] pattern=${pat} matched="${m[0]}" raw=${raw} 税込=${isTaxIncluded} → ${result}`);
      return result;
    }
  }
  return 0;
}

/**
 * 複数パターン・複数マッチを全行合算して返す（beds24/maneKEY等の複数行対応）
 * @param {string} text
 * @param {RegExp[]} patterns - グローバルフラグ付き正規表現の配列
 * @return {number} 税込合計金額
 */
function extractAllAmountsWithTax_(text, patterns) {
  let total = 0;
  patterns.forEach(pat => {
    for (const m of text.matchAll(pat)) {
      const raw = parseInt(m[1].replace(/,/g, '')) || 0;
      if (raw === 0) continue;
      const idx = text.indexOf(m[0]);
      const context = text.substring(Math.max(0, idx - 10), idx + m[0].length + 40);
      const isTaxIncluded = /税込/.test(context);
      const val = isTaxIncluded ? raw : Math.round(raw * 1.1);
      Logger.log(`  [その他抽出] matched="${m[0]}" raw=${raw} 税込=${isTaxIncluded} → ${val}`);
      total += val;
    }
  });
  Logger.log(`  [その他合計] ${total}`);
  return total;
}

// ==============================
// 経費入力シートへの書き込み
// ==============================

/**
 * 解析結果を経費入力シートの該当月行に書き込む
 * @param {Object} parsed - { yearMonth, agencyFee, cleaning, linen, supplies }
 * @return {boolean} 書き込み成功なら true
 */
function writeCostToSheet_(parsed) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COSTS);

  if (!sheet) {
    Logger.log('経費入力シートが存在しません');
    return false;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('経費入力シートにデータがありません（migrateCostSheet を実行してください）');
    return false;
  }

  // 年月列を検索して対象行を特定
  const ymValues = sheet.getRange(2, CONFIG.COST_COLS.YEAR_MONTH, lastRow - 1, 1).getValues();
  let targetRow = -1;

  for (let i = 0; i < ymValues.length; i++) {
    const cell = ymValues[i][0];
    let cellStr;
    if (cell instanceof Date) {
      cellStr = `${cell.getFullYear()}-${String(cell.getMonth() + 1).padStart(2, '0')}`;
    } else {
      cellStr = String(cell).trim();
    }
    if (cellStr === parsed.yearMonth) {
      targetRow = i + 2; // 1-indexed, +1 for header
      break;
    }
  }

  if (targetRow === -1) {
    // 対象月の行がなければ末尾に追加
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, CONFIG.COST_COLS.YEAR_MONTH).setValue(parsed.yearMonth);
    Logger.log(`経費入力: ${parsed.yearMonth} の行が存在しないため末尾に追加`);
  }

  const C = CONFIG.COST_COLS;

  // 請求書から取得した値で上書き（加算しない）
  sheet.getRange(targetRow, C.AGENCY_FEE).setValue(parsed.agencyFee);
  sheet.getRange(targetRow, C.CLEANING).setValue(parsed.cleaning);
  sheet.getRange(targetRow, C.LINEN).setValue(parsed.linen);
  sheet.getRange(targetRow, C.SUPPLIES).setValue(parsed.supplies);
  sheet.getRange(targetRow, C.OTHER).setValue(parsed.others);

  return true;
}

// ==============================
// ユーティリティ
// ==============================

/**
 * Gmail ラベルを取得または作成する
 * @param {string} name
 * @return {GmailLabel}
 */
function getOrCreateLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

// ==============================
// デバッグ用：処理済みラベルのリセット
// ==============================

/**
 * エアサポ処理済みラベルをすべてのメールから外す
 * 再取得・再解析したい場合にGASエディタから直接実行する
 */
function resetAirSapoLabels() {
  const label = GmailApp.getUserLabelByName('民泊_エアサポ処理済み');
  if (!label) {
    Logger.log('ラベル「民泊_エアサポ処理済み」が存在しません（すでにクリーン状態）');
    return;
  }
  const threads = label.getThreads();
  threads.forEach(t => t.removeLabel(label));
  Logger.log(`処理済みラベルを ${threads.length} スレッドから削除しました`);
}

// ==============================
// デバッグ用：メール検索確認
// ==============================
/**
 * エアサポ関連メールをさまざまなクエリで検索してログに出力する
 * GASエディタで直接実行して差出人・件名を確認する
 */
function debugSearchAirSapoMail() {
  const queries = [
    'from:support@air-sapo.com',
    'from:air-sapo.com',
    'subject:請求書',
    'subject:ご請求',
    'エアサポ',
    'air-sapo'
  ];

  queries.forEach(q => {
    const threads = GmailApp.search(q, 0, 5);
    Logger.log(`【${q}】→ ${threads.length}件`);
    threads.forEach(t => {
      const msg = t.getMessages()[0];
      Logger.log(`  差出人: ${msg.getFrom()}`);
      Logger.log(`  件名: ${msg.getSubject()}`);
      Logger.log(`  日時: ${msg.getDate()}`);
    });
  });
}
