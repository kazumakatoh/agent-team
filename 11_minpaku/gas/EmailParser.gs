/**
 * 民泊自動化システム - メール解析モジュール
 * Airbnb / Booking.com の予約確定メールを解析して予約データを抽出する
 */

/**
 * メール本文からチェックイン・チェックアウト日を抽出する共通ヘルパー
 * 「チェックイン」「チェックアウト」キーワードの後に来る日付を優先使用し、
 * 見つからない場合はメール本文中の全日付から最初の2件を使用する
 * @param {string} body メール本文
 * @return {{checkin: Date|null, checkout: Date|null}}
 */
function extractDatesFromBody_(body) {
  function parseDate_(s) {
    let m;
    m = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (m) return new Date(+m[1], +m[2]-1, +m[3]);
    m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2]-1, +m[3]);
    m = s.match(/(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/i);
    if (m) { const d = new Date(`${m[2]} ${m[1]}, ${m[3]}`); if (!isNaN(d)) return d; }
    m = s.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})/i);
    if (m) { const d = new Date(`${m[1]} ${m[2]}, ${m[3]}`); if (!isNaN(d)) return d; }
    return null;
  }

  // キーワードの後（同行 + 次行）から日付を取得
  function findAfterKeyword_(keywords) {
    for (const kw of keywords) {
      const re = new RegExp(kw + '[^\\n]{0,50}(?:\\n[^\\n]{0,50})?', 'i');
      const m = body.match(re);
      if (m) {
        const d = parseDate_(m[0]);
        if (d && !isNaN(d)) return d;
      }
    }
    return null;
  }

  const checkin  = findAfterKeyword_(['チェックイン',  'Check-?in']);
  const checkout = findAfterKeyword_(['チェックアウト', 'Check-?out']);
  if (checkin && checkout) return { checkin, checkout };

  // フォールバック: 全日付を集めて昇順ソート後、最初の2つ
  const dates = [];
  let m;
  const p1 = /(\d{4})年(\d{1,2})月(\d{1,2})日/g;
  while ((m = p1.exec(body)) !== null) dates.push(new Date(+m[1], +m[2]-1, +m[3]));
  const p2 = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})/gi;
  while ((m = p2.exec(body)) !== null) { const d = new Date(`${m[1]} ${m[2]}, ${m[3]}`); if (!isNaN(d)) dates.push(d); }
  const p3 = /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{4})/gi;
  while ((m = p3.exec(body)) !== null) { const d = new Date(`${m[2]} ${m[1]}, ${m[3]}`); if (!isNaN(d)) dates.push(d); }

  const unique = [...new Set(dates.filter(d => !isNaN(d)).map(d => d.getTime()))]
    .sort((a, b) => a - b).map(t => new Date(t));

  return { checkin: checkin || unique[0] || null, checkout: checkout || unique[1] || null };
}

/**
 * メール本文から宿泊人数を抽出する共通ヘルパー
 * 以下のキーワードの後に続く数字を合算して返す:
 *   People / Adults / Children / 大人 / 子供 / 人
 * @param {string} body メール本文
 * @return {number} 宿泊人数（最低1）
 */
function extractGuestsFromBody_(body) {
  // Adults + Children の合算（英語形式）
  const adultsMatch   = body.match(/Adults?\s*[：:＝=]?\s*(\d+)/i);
  const childrenMatch = body.match(/Children?\s*[：:＝=]?\s*(\d+)/i);
  if (adultsMatch) {
    return parseInt(adultsMatch[1]) + (childrenMatch ? parseInt(childrenMatch[1]) : 0);
  }

  // People（既に合計人数）
  const peopleMatch = body.match(/People\s*[：:＝=]?\s*(\d+)/i);
  if (peopleMatch) return parseInt(peopleMatch[1]);

  // 大人 + 子供（日本語形式）
  const daijinMatch = body.match(/大人\s*[：:＝=]?\s*(\d+)/);
  const kodomMatch  = body.match(/子供\s*[：:＝=]?\s*(\d+)/);
  if (daijinMatch) {
    return parseInt(daijinMatch[1]) + (kodomMatch ? parseInt(kodomMatch[1]) : 0);
  }

  // 人・名・ゲスト・guests（汎用パターン）
  const generalMatch = body.match(/[人名]\s*[：:＝=]?\s*(\d+)/) ||
                       body.match(/(\d+)\s*(?:名|人|ゲスト|guests?)/i);
  if (generalMatch) return parseInt(generalMatch[1]);

  return 1;
}

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

  // Beds24経由の場合、件名末尾が「- Airbnb」のものだけAirbnb
  if (isBeds24 && detectPlatformFromSubject_(subject) !== 'Airbnb') return null;

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

    // 予約ID: Beds24が割り振る「Booking Ref: 8桁」を最優先、次に「予約ID: 8桁」
    // AirbnbのHMコードやConfirmation codeはプラットフォーム依存のため採用しない
    const idMatch = body.match(/Booking Ref[：:\s]*(\d{8})/i) ||
                    body.match(/予約ID[：:\s]*(\d{8})/);
    data.reservationId = idMatch ? idMatch[1] : `AB_${msg.getId().substring(0,8)}`;

    // チェックイン・チェックアウト日
    const { checkin: ci, checkout: co } = extractDatesFromBody_(body);
    data.checkin  = ci;
    data.checkout = co;
    data.nights   = (ci && co) ? Math.round((co - ci) / (1000 * 60 * 60 * 24)) : 0;

    // 宿泊人数
    data.guests = extractGuestsFromBody_(body);

    // ゲスト名（Beds24形式: "Name Rene Kolenkovic" も対応）
    const guestNameMatch = body.match(/ゲスト[：:\s]+([^\n\r]+)/) ||
                           body.match(/Guest[：:\s]+([^\n\r]+)/i) ||
                           body.match(/^Name\s+(.+)$/im);
    data.guestName = guestNameMatch ? guestNameMatch[1].trim() : '';

    // 売上：Total Price（ゲストが支払う合計金額）
    const priceMatch = body.match(/Total Price\s+([\d,]+(?:\.\d+)?)/i) ||
                       body.match(/合計[：:\s]*[¥￥]?\s*([\d,]+)/);
    data.revenue = priceMatch ? parseInt(priceMatch[1].replace(/,/g, '')) : 0;

    // 宿泊料：Base Price
    const basePriceMatch = body.match(/Base Price\s+([\d,]+(?:\.\d+)?)\s*JPY/i) ||
                           body.match(/Base Price[：:\s]*([\d,]+)/i);
    data.accommodationFee = basePriceMatch ? parseInt(basePriceMatch[1].replace(/,/g, '')) : 0;

    // 清掃費：Cleaning fee（費用ではなくゲスト負担の売上の一部）
    const cleaningMatch = body.match(/Cleaning fee\s+([\d,]+(?:\.\d+)?)\s*JPY/i) ||
                          body.match(/Cleaning fee[：:\s]*[¥￥$]?\s*([\d,]+)/i);
    data.cleaningFee = cleaningMatch ? parseInt(cleaningMatch[1].replace(/,/g, '')) : 0;

    // OTA手数料：Host Fee（Airbnbに支払う販売手数料・負の値で記載されるため絶対値）
    const hostFeeMatch = body.match(/Host Fee\s*([-\d,]+(?:\.\d+)?)\s*JPY/i) ||
                         body.match(/Host Fee[：:\s]*([-\d,]+)/i);
    data.otaFee = hostFeeMatch ? Math.abs(parseInt(hostFeeMatch[1].replace(/,/g, ''))) : 0;

    // 利用日数・総利用人数
    data.usageDays   = (data.nights || 0) + 1;
    data.totalGuests = data.usageDays * (data.guests || 1);

    // 振込手数料：Airbnbはメール記載なし → 常に0
    data.transferFee = 0;

    // 入金金額：Expected Payout Amount
    const payoutMatch = body.match(/Expected Payout Amount\s+([\d,]+(?:\.\d+)?)\s*JPY/i) ||
                        body.match(/Expected Payout Amount[：:\s]*([\d,]+)/i);
    data.payoutAmount = payoutMatch ? parseInt(payoutMatch[1].replace(/,/g, '')) : 0;

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

  // Beds24経由の場合、件名末尾が「- Booking.com」のものだけBooking.com
  if (isBeds24Booking && detectPlatformFromSubject_(subject) !== 'Booking.com') return null;

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

    // 予約番号：Beds24が割り振る「Booking Ref: 8桁」を最優先、次に「予約ID: 8桁」
    // BC_ プレフィックスは付けない（Airbnbと統一したBeds24番号を採用）
    const idMatch = body.match(/Booking Ref[：:\s]*(\d{8})/i) ||
                    body.match(/予約ID[：:\s]*(\d{8})/);
    data.reservationId = idMatch ? idMatch[1] : `BC_${msg.getId().substring(0,8)}`;

    // チェックイン・チェックアウト日
    const { checkin: ci2, checkout: co2 } = extractDatesFromBody_(body);
    data.checkin  = ci2;
    data.checkout = co2;
    data.nights   = (ci2 && co2) ? Math.round((co2 - ci2) / (1000 * 60 * 60 * 24)) : 0;

    // 宿泊人数
    data.guests = extractGuestsFromBody_(body);

    // ゲスト名（Beds24形式: "名前 CZARINA CATAMBING" も対応）
    const nameMatch = body.match(/^名前\s+(.+)$/im) ||
                      body.match(/ゲスト名[：:\s]+([^\n\r]+)/) ||
                      body.match(/Guest name[：:\s]+([^\n\r]+)/i) ||
                      body.match(/^Name\s+(.+)$/im);
    data.guestName = nameMatch ? nameMatch[1].trim() : '';

    // 売上：合計金額（Beds24形式: "価格 69,779.00" も対応）
    const priceMatch = body.match(/合計金額[：:\s]*[¥￥]?\s*([\d,]+)/) ||
                       body.match(/^価格\s+([\d,]+(?:\.\d+)?)/im) ||
                       body.match(/Total Price\s+([\d,]+(?:\.\d+)?)/i);
    data.revenue = priceMatch ? parseInt(priceMatch[1].replace(/,/g, '')) : 0;

    // 宿泊料：Standard Rate（連泊の場合は複数行を合算）
    // 例: "Standard Rate) JPY 69779" の形式
    const standardRateMatches = [...body.matchAll(/Standard Rate[^0-9\n]*([\d,]+)/gi)];
    data.accommodationFee = standardRateMatches.reduce(
      (sum, m) => sum + parseInt(m[1].replace(/,/g, '')), 0
    );

    // 清掃費：メール記載なし → 固定14,410円（費用ではなくゲスト負担の売上の一部）
    data.cleaningFee = 14410;

    // OTA手数料：Total Commission（Booking.comに支払う販売手数料）
    const commissionMatch = body.match(/Total Commission\s+([\d,]+(?:\.\d+)?)/i) ||
                            body.match(/Total Commission[：:\s]*([\d,]+)/i);
    data.otaFee = commissionMatch ? parseInt(commissionMatch[1].replace(/,/g, '')) : 0;

    // 振込手数料：Payment Charge
    const transferMatch = body.match(/Payment Charge\s+([\d,]+(?:\.\d+)?)/i) ||
                          body.match(/Payment Charge[：:\s]*([\d,]+)/i);
    data.transferFee = transferMatch ? parseInt(transferMatch[1].replace(/,/g, '')) : 0;

    // 利用日数・総利用人数
    data.usageDays   = (data.nights || 0) + 1;
    data.totalGuests = data.usageDays * (data.guests || 1);

    // 入金金額：売上 - OTA手数料 - 振込手数料
    data.payoutAmount = data.revenue - data.otaFee - data.transferFee;
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

  // プラットフォームを件名末尾から判定
  const platform = detectPlatformFromSubject_(subject) || 'Booking.com';

  // 予約IDを抽出（Beds24の「Booking Ref: 8桁」を最優先、両プラットフォーム共通）
  let reservationId = null;
  const beds24Match = body.match(/Booking Ref[：:\s]*(\d{8})/i) ||
                      body.match(/予約ID[：:\s]*(\d{8})/);
  reservationId = beds24Match ? beds24Match[1] : null;
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

/**
 * 件名末尾からプラットフォームを判定する
 * 例: "予約: ... - Booking.com" → 'Booking.com'
 *     "予約: ... - Airbnb"      → 'Airbnb'
 * @param {string} subject
 * @return {string|null} 'Airbnb' | 'Booking.com' | null
 */
function detectPlatformFromSubject_(subject) {
  if (/- Booking\.com\s*$/i.test(subject)) return 'Booking.com';
  if (/- Airbnb\s*$/i.test(subject))       return 'Airbnb';
  return null;
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
