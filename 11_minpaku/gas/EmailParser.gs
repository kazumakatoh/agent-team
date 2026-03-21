/**
 * 民泊自動化システム - メール解析モジュール
 * Airbnb / Booking.com の予約確定メールを解析して予約データを抽出する
 */

/**
 * 未処理の予約確定メールを全件取得して解析する
 * @return {Array} 解析済み予約データの配列
 */
function fetchNewReservationEmails() {
  const reservations = [];
  const processedLabel = getOrCreateLabel_(CONFIG.GMAIL.PROCESSED_LABEL);
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - CONFIG.GMAIL.SEARCH_DAYS);

  // Airbnb メール検索
  const airbnbQuery = buildGmailQuery_('AIRBNB', cutoffDate);
  const airbnbThreads = GmailApp.search(airbnbQuery);
  airbnbThreads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      if (isAlreadyProcessed_(msg, processedLabel)) return;
      const data = parseAirbnbEmail(msg);
      if (data) {
        reservations.push(data);
        msg.getThread().addLabel(processedLabel);
      }
    });
  });

  // Booking.com メール検索
  const bookingQuery = buildGmailQuery_('BOOKING', cutoffDate);
  const bookingThreads = GmailApp.search(bookingQuery);
  bookingThreads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      if (isAlreadyProcessed_(msg, processedLabel)) return;
      const data = parseBookingEmail(msg);
      if (data) {
        reservations.push(data);
        msg.getThread().addLabel(processedLabel);
      }
    });
  });

  Logger.log(`新規予約メール検出数: ${reservations.length}件`);
  return reservations;
}

/**
 * Airbnb 予約確定メールを解析する
 * @param {GmailMessage} msg
 * @return {Object|null} 予約データ、解析失敗時はnull
 */
function parseAirbnbEmail(msg) {
  const subject = msg.getSubject();
  const body    = msg.getPlainBody();
  const from    = msg.getFrom();

  // Airbnbからのメールか確認
  const isBeds24 = from.includes('email.master@co-reception.com');
  const isAirbnb = CONFIG.GMAIL.FROM.AIRBNB.some(f => from.includes(f)) ||
                   from.includes('airbnb');
  if (!isAirbnb) return null;

  // Beds24経由の場合、件名が「予約:」で始まるものだけAirbnb
  if (isBeds24 && !subject.startsWith('予約:')) return null;

  // 予約確定メールか確認（キャンセル・変更は除外）
  const isCancel = CONFIG.GMAIL.SUBJECTS.CANCEL.some(s => subject.includes(s));
  const isConfirmation = !isCancel && CONFIG.GMAIL.SUBJECTS.AIRBNB.some(s =>
    subject.includes(s)
  );
  if (!isConfirmation) return null;

  try {
    const data = {
      emailId:     msg.getId(),
      platform:    'Airbnb',
      bookedDate:  msg.getDate(),
      status:      '予約'
    };

    // 予約ID（例: HM... または数字、Beds24形式も対応）
    const idMatch = body.match(/予約コード[：:\s]*([A-Z0-9]+)/i) ||
                    body.match(/Confirmation code[：:\s]*([A-Z0-9]+)/i) ||
                    body.match(/Airbnb\s+([A-Z0-9]{6,})/i) ||
                    body.match(/([A-Z]{2}[0-9]{9})/);
    data.reservationId = idMatch ? idMatch[1] : `AB_${msg.getId().substring(0,8)}`;

    // チェックイン・チェックアウト日
    // 日本語形式: 2025年12月25日
    const jpDatePattern = /(\d{4})年(\d{1,2})月(\d{1,2})日/g;
    const jpDates = [];
    let m;
    while ((m = jpDatePattern.exec(body)) !== null) {
      jpDates.push(new Date(m[1], m[2] - 1, m[3]));
    }

    // 英語形式: December 25, 2025 または 25 Dec 2025
    const enDatePattern1 = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})/gi;
    const enDatePattern2 = /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{4})/gi;
    const enDates = [];
    while ((m = enDatePattern1.exec(body)) !== null) {
      enDates.push(new Date(`${m[1]} ${m[2]}, ${m[3]}`));
    }
    while ((m = enDatePattern2.exec(body)) !== null) {
      enDates.push(new Date(`${m[2]} ${m[1]}, ${m[3]}`));
    }

    const allDates = [...jpDates, ...enDates].filter(d => !isNaN(d));
    if (allDates.length >= 2) {
      // 重複除去・ソートして最初の2つをチェックイン・チェックアウトに
      const uniqueDates = [...new Set(allDates.map(d => d.getTime()))]
        .sort((a, b) => a - b)
        .map(t => new Date(t));
      data.checkin  = uniqueDates[0];
      data.checkout = uniqueDates[1];
      data.nights   = Math.round((uniqueDates[1] - uniqueDates[0]) / (1000 * 60 * 60 * 24));
    }

    // 宿泊人数（Beds24形式: "People 2" も対応）
    const guestMatch = body.match(/(\d+)\s*(名|人|ゲスト|guests?)/i) ||
                       body.match(/^People\s+(\d+)/im);
    data.guests = guestMatch ? parseInt(guestMatch[1]) : 1;

    // ゲスト名（Beds24形式: "Name Rene Kolenkovic" も対応）
    const guestNameMatch = body.match(/ゲスト[：:\s]+([^\n\r]+)/) ||
                           body.match(/Guest[：:\s]+([^\n\r]+)/i) ||
                           body.match(/^Name\s+(.+)$/im);
    data.guestName = guestNameMatch ? guestNameMatch[1].trim() : '';

    // 売上（金額）（Beds24形式: "Total Price 88,813.00" も対応）
    const priceMatch = body.match(/合計[：:\s]*[¥￥]?\s*([\d,]+)/) ||
                       body.match(/Total Price\s+([\d,]+(?:\.\d+)?)/i) ||
                       body.match(/Total[：:\s]*[¥￥$]?\s*([\d,]+)/i) ||
                       body.match(/[¥￥]([\d,]+)/);
    data.revenue = priceMatch ? parseInt(priceMatch[1].replace(/,/g, '')) : 0;

    // 清掃費
    const cleaningMatch = body.match(/清掃[料費][：:\s]*[¥￥]?\s*([\d,]+)/) ||
                          body.match(/Cleaning fee[：:\s]*[¥￥$]?\s*([\d,]+)/i);
    data.cleaningFee = cleaningMatch ? parseInt(cleaningMatch[1].replace(/,/g, '')) : 0;

    // 手数料計算（Airbnbホスト手数料は売上の3%）
    data.commission = Math.round(data.revenue * CONFIG.COMMISSION_RATE.AIRBNB);

    Logger.log(`Airbnb予約解析成功: ${data.reservationId} (${data.checkin} - ${data.checkout})`);
    return data;

  } catch (e) {
    Logger.log(`Airbnbメール解析エラー (${msg.getId()}): ${e.message}`);
    return null;
  }
}

/**
 * Booking.com 予約確定メールを解析する
 * @param {GmailMessage} msg
 * @return {Object|null} 予約データ、解析失敗時はnull
 */
function parseBookingEmail(msg) {
  const subject = msg.getSubject();
  const body    = msg.getPlainBody();
  const from    = msg.getFrom();

  const isBeds24Booking = from.includes('email.master@co-reception.com');
  const isBooking = CONFIG.GMAIL.FROM.BOOKING.some(f => from.includes(f)) ||
                    from.includes('booking.com');
  if (!isBooking) return null;

  // Beds24経由の場合、件名が「Booking Modified:」で始まるものだけBooking.com
  if (isBeds24Booking && !subject.startsWith('Booking Modified:')) return null;

  // キャンセルメールは除外
  const isCancel = CONFIG.GMAIL.SUBJECTS.CANCEL.some(s => subject.includes(s));
  const isConfirmation = !isCancel && CONFIG.GMAIL.SUBJECTS.BOOKING.some(s =>
    subject.includes(s)
  );
  if (!isConfirmation) return null;

  try {
    const data = {
      emailId:   msg.getId(),
      platform:  'Booking.com',
      bookedDate: msg.getDate(),
      status:    '予約'
    };

    // 予約番号（Beds24形式: "予約ID: 83458884" も対応）
    const idMatch = body.match(/予約ID[：:\s]*(\d+)/i) ||
                    body.match(/予約番号[：:\s]*(\d+)/) ||
                    body.match(/Booking number[：:\s]*(\d+)/i) ||
                    body.match(/Booking Ref[：:\s]*(\d+)/i) ||
                    body.match(/Reservation number[：:\s]*(\d+)/i);
    data.reservationId = idMatch ? `BC_${idMatch[1]}` : `BC_${msg.getId().substring(0,8)}`;

    // チェックイン・チェックアウト（英語日付形式: "Fri 13 Mar 2026" も対応）
    const enDate2 = /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{4})/i;
    const checkinLine  = body.match(/チェックイン[^\n]*\n?([^\n]+)/) || body.match(/チェックイン\s+\S+\s+(\d.+)/);
    const checkoutLine = body.match(/チェックアウト[^\n]*\n?([^\n]+)/) || body.match(/チェックアウト\s+\S+\s+(\d.+)/);

    const checkinMatch = body.match(/チェックイン[：:\s]*(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/) ||
                         body.match(/チェックイン\s+\w+\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})/i) ||
                         body.match(/Check-?in[：:\s]*(\w+)\s+(\d{1,2}),?\s+(\d{4})/i);
    const checkoutMatch = body.match(/チェックアウト[：:\s]*(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/) ||
                          body.match(/チェックアウト\s+\w+\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})/i) ||
                          body.match(/Check-?out[：:\s]*(\w+)\s+(\d{1,2}),?\s+(\d{4})/i);

    if (checkinMatch && checkoutMatch) {
      if (/^\d{4}$/.test(checkinMatch[1])) {
        // 日本語形式: 年/月/日
        data.checkin  = new Date(checkinMatch[1], checkinMatch[2] - 1, checkinMatch[3]);
        data.checkout = new Date(checkoutMatch[1], checkoutMatch[2] - 1, checkoutMatch[3]);
      } else if (/^\d{1,2}$/.test(checkinMatch[1])) {
        // 英語形式: DD Mon YYYY
        data.checkin  = new Date(`${checkinMatch[2]} ${checkinMatch[1]}, ${checkinMatch[3]}`);
        data.checkout = new Date(`${checkoutMatch[2]} ${checkoutMatch[1]}, ${checkoutMatch[3]}`);
      } else {
        data.checkin  = new Date(`${checkinMatch[1]} ${checkinMatch[2]}, ${checkinMatch[3]}`);
        data.checkout = new Date(`${checkoutMatch[1]} ${checkoutMatch[2]}, ${checkoutMatch[3]}`);
      }
      data.nights = Math.round((data.checkout - data.checkin) / (1000 * 60 * 60 * 24));
    }

    // 宿泊人数（大人+子供の合計、Beds24形式対応）
    const adultMatch = body.match(/大人\s*(\d+)/);
    const childMatch = body.match(/子供\s*(\d+)/);
    if (adultMatch) {
      data.guests = parseInt(adultMatch[1]) + (childMatch ? parseInt(childMatch[1]) : 0);
    } else {
      const guestMatch = body.match(/(\d+)\s*(名|人|guests?|adults?)/i);
      data.guests = guestMatch ? parseInt(guestMatch[1]) : 1;
    }

    // ゲスト名（Beds24形式: "名前 CZARINA CATAMBING" も対応）
    const nameMatch = body.match(/^名前\s+(.+)$/im) ||
                      body.match(/ゲスト名[：:\s]+([^\n\r]+)/) ||
                      body.match(/Guest name[：:\s]+([^\n\r]+)/i) ||
                      body.match(/^Name\s+(.+)$/im);
    data.guestName = nameMatch ? nameMatch[1].trim() : '';

    // 売上（Beds24形式: "価格 69,779.00" も対応）
    const priceMatch = body.match(/^価格\s+([\d,]+(?:\.\d+)?)/im) ||
                       body.match(/合計金額[：:\s]*[¥￥]?\s*([\d,]+)/) ||
                       body.match(/Total Price\s+([\d,]+(?:\.\d+)?)/i) ||
                       body.match(/Total[：:\s]*[¥￥]?\s*([\d,]+)/i) ||
                       body.match(/[¥￥]([\d,]+)/);
    data.revenue = priceMatch ? parseInt(priceMatch[1].replace(/,/g, '')) : 0;

    // 清掃費
    const cleaningMatch = body.match(/清掃[料費][：:\s]*[¥￥]?\s*([\d,]+)/);
    data.cleaningFee = cleaningMatch ? parseInt(cleaningMatch[1].replace(/,/g, '')) : 0;

    // Booking.com手数料（売上の15%）
    data.commission = Math.round(data.revenue * CONFIG.COMMISSION_RATE.BOOKING);

    Logger.log(`Booking.com予約解析成功: ${data.reservationId}`);
    return data;

  } catch (e) {
    Logger.log(`Booking.comメール解析エラー (${msg.getId()}): ${e.message}`);
    return null;
  }
}

/**
 * キャンセルメールを解析して予約IDを返す
 * @param {GmailMessage} msg
 * @return {Object|null} { reservationId, emailId, cancelDate }
 */
function parseCancellationEmail(msg) {
  const subject = msg.getSubject();
  const body    = msg.getPlainBody();

  const isCancel = CONFIG.GMAIL.SUBJECTS.CANCEL.some(s => subject.includes(s)) ||
                   body.includes('has been cancelled') ||
                   body.includes('キャンセル');
  if (!isCancel) return null;

  // プラットフォームを件名から判定
  const platform = subject.includes('- Airbnb') ? 'Airbnb' : 'Booking.com';

  // 予約IDを抽出（AirbnbコードまたはBooking Ref）
  let reservationId = null;
  if (platform === 'Airbnb') {
    const m = body.match(/Airbnb\s+([A-Z0-9]{6,})/i) ||
              body.match(/予約コード[：:\s]*([A-Z0-9]+)/i);
    reservationId = m ? m[1] : null;
  } else {
    const m = body.match(/予約ID[：:\s]*(\d+)/i) ||
              body.match(/Booking Ref[：:\s]+([0-9]+)/i);
    reservationId = m ? `BC_${m[1]}` : null;
  }
  if (!reservationId) return null;

  return {
    reservationId,
    platform,
    emailId:    msg.getId(),
    cancelDate: msg.getDate()
  };
}

/**
 * キャンセルメールを検索して処理する
 * @param {Date|null} sinceDate - この日以降を対象。nullで全期間
 * @return {number} キャンセル更新件数
 */
function processCancellationEmails(sinceDate) {
  const processedLabel = getOrCreateLabel_(CONFIG.GMAIL.PROCESSED_LABEL);

  const fromAddresses = [
    ...CONFIG.GMAIL.FROM.AIRBNB,
    ...CONFIG.GMAIL.FROM.BOOKING
  ].filter((v, i, a) => a.indexOf(v) === i)
   .map(f => `from:${f}`).join(' OR ');

  const datePart = sinceDate
    ? ` after:${Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd')}`
    : (() => {
        const d = new Date();
        d.setDate(d.getDate() - CONFIG.GMAIL.SEARCH_DAYS);
        return ` after:${Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd')}`;
      })();

  // キャンセルメールは処理済みラベルを除外しない（すでにラベル付きでも再処理する）
  const query = `(${fromAddresses})${datePart}`;
  const threads = GmailApp.search(query, 0, 500);

  let updated = 0;
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      // 件名で「予約キャンセルになりました:」のみ対象
      if (!msg.getSubject().startsWith('予約キャンセルになりました:')) return;
      const cancel = parseCancellationEmail(msg);
      if (cancel) {
        const result = updateReservationStatus(cancel.reservationId, 'キャンセル');
        if (result) updated++;
        Logger.log(`キャンセル処理: ${cancel.reservationId} → ${result ? '成功' : '予約ID未発見'}`);
      }
    });
  });

  return updated;
}

// ==============================
// プライベート関数
// ==============================

function buildGmailQuery_(platform, cutoffDate) {
  const fromAddresses = CONFIG.GMAIL.FROM[platform].map(f => `from:${f}`).join(' OR ');
  const dateStr = Utilities.formatDate(cutoffDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  return `(${fromAddresses}) after:${dateStr} -label:${CONFIG.GMAIL.PROCESSED_LABEL}`;
}

/**
 * 指定日以降の予約確定メールを全件取得して解析する（遡り取込用）
 * @param {Date|null} sinceDate - この日以降のメールを対象。null の場合は全期間
 * @return {Array} 解析済み予約データの配列
 */
function fetchReservationEmailsSince(sinceDate) {
  const reservations = [];
  const processedLabel = getOrCreateLabel_(CONFIG.GMAIL.PROCESSED_LABEL);

  const buildQuery = (platform) => {
    const fromAddresses = CONFIG.GMAIL.FROM[platform].map(f => `from:${f}`).join(' OR ');
    const datePart = sinceDate
      ? ` after:${Utilities.formatDate(sinceDate, 'Asia/Tokyo', 'yyyy/MM/dd')}`
      : '';
    // 処理済みラベルを除外しない（過去分は重複をEmailIDで防ぐ）
    return `(${fromAddresses})${datePart}`;
  };

  [
    { platform: 'AIRBNB',  parser: parseAirbnbEmail },
    { platform: 'BOOKING', parser: parseBookingEmail }
  ].forEach(({ platform, parser }) => {
    const query   = buildQuery(platform);
    const threads = GmailApp.search(query, 0, 500); // 最大500スレッド

    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const data = parser(msg);
        if (data) {
          reservations.push(data);
          // 処理済みラベルを付与（次回の通常チェックで重複しないよう）
          thread.addLabel(processedLabel);
        }
      });
    });
  });

  Logger.log(`遡り取込: ${reservations.length}件のメールを解析`);
  return reservations;
}

function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log(`Gmailラベル作成: ${labelName}`);
  }
  return label;
}

function isAlreadyProcessed_(msg, processedLabel) {
  const labels = msg.getThread().getLabels();
  return labels.some(l => l.getName() === processedLabel.getName());
}
