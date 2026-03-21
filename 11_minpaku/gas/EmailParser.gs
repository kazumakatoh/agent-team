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
  const isAirbnb = CONFIG.GMAIL.FROM.AIRBNB.some(f => from.includes(f)) ||
                   from.includes('airbnb');
  if (!isAirbnb) return null;

  // 予約確定メールか確認
  const isConfirmation = CONFIG.GMAIL.SUBJECTS.AIRBNB.some(s =>
    subject.includes(s)
  );
  if (!isConfirmation) return null;

  try {
    const data = {
      emailId:     msg.getId(),
      platform:    'Airbnb',
      bookedDate:  msg.getDate(),
      status:      '確定'
    };

    // 予約ID（例: HM... または数字）
    const idMatch = body.match(/予約コード[：:\s]*([A-Z0-9]+)/i) ||
                    body.match(/Confirmation code[：:\s]*([A-Z0-9]+)/i) ||
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

    // 宿泊人数
    const guestMatch = body.match(/(\d+)\s*(名|人|ゲスト|guests?)/i);
    data.guests = guestMatch ? parseInt(guestMatch[1]) : 1;

    // ゲスト名
    const guestNameMatch = body.match(/ゲスト[：:\s]+([^\n\r]+)/) ||
                           body.match(/Guest[：:\s]+([^\n\r]+)/i);
    data.guestName = guestNameMatch ? guestNameMatch[1].trim() : '';

    // 売上（金額）
    const priceMatch = body.match(/合計[：:\s]*[¥￥]?\s*([\d,]+)/) ||
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

  const isBooking = CONFIG.GMAIL.FROM.BOOKING.some(f => from.includes(f)) ||
                    from.includes('booking.com');
  if (!isBooking) return null;

  const isConfirmation = CONFIG.GMAIL.SUBJECTS.BOOKING.some(s =>
    subject.includes(s)
  );
  if (!isConfirmation) return null;

  try {
    const data = {
      emailId:   msg.getId(),
      platform:  'Booking.com',
      bookedDate: msg.getDate(),
      status:    '確定'
    };

    // 予約番号
    const idMatch = body.match(/予約番号[：:\s]*(\d+)/) ||
                    body.match(/Booking number[：:\s]*(\d+)/i) ||
                    body.match(/Reservation number[：:\s]*(\d+)/i);
    data.reservationId = idMatch ? `BC_${idMatch[1]}` : `BC_${msg.getId().substring(0,8)}`;

    // チェックイン・チェックアウト
    const checkinMatch = body.match(/チェックイン[：:\s]*(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/) ||
                         body.match(/Check-?in[：:\s]*(\w+)\s+(\d{1,2}),?\s+(\d{4})/i);
    const checkoutMatch = body.match(/チェックアウト[：:\s]*(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/) ||
                          body.match(/Check-?out[：:\s]*(\w+)\s+(\d{1,2}),?\s+(\d{4})/i);

    if (checkinMatch && checkoutMatch) {
      // 日本語形式
      if (/^\d{4}$/.test(checkinMatch[1])) {
        data.checkin  = new Date(checkinMatch[1], checkinMatch[2] - 1, checkinMatch[3]);
        data.checkout = new Date(checkoutMatch[1], checkoutMatch[2] - 1, checkoutMatch[3]);
      } else {
        data.checkin  = new Date(`${checkinMatch[1]} ${checkinMatch[2]}, ${checkinMatch[3]}`);
        data.checkout = new Date(`${checkoutMatch[1]} ${checkoutMatch[2]}, ${checkoutMatch[3]}`);
      }
      data.nights = Math.round((data.checkout - data.checkin) / (1000 * 60 * 60 * 24));
    }

    // 宿泊人数
    const guestMatch = body.match(/(\d+)\s*(名|人|大人|guests?|adults?)/i);
    data.guests = guestMatch ? parseInt(guestMatch[1]) : 1;

    // ゲスト名
    const nameMatch = body.match(/ゲスト名[：:\s]+([^\n\r]+)/) ||
                      body.match(/Guest name[：:\s]+([^\n\r]+)/i) ||
                      body.match(/Name[：:\s]+([^\n\r]+)/i);
    data.guestName = nameMatch ? nameMatch[1].trim() : '';

    // 売上
    const priceMatch = body.match(/合計金額[：:\s]*[¥￥]?\s*([\d,]+)/) ||
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

// ==============================
// プライベート関数
// ==============================

function buildGmailQuery_(platform, cutoffDate) {
  const fromAddresses = CONFIG.GMAIL.FROM[platform].map(f => `from:${f}`).join(' OR ');
  const dateStr = Utilities.formatDate(cutoffDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  return `(${fromAddresses}) after:${dateStr} -label:${CONFIG.GMAIL.PROCESSED_LABEL}`;
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
