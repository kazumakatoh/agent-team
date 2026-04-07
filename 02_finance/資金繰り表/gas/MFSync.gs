/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * 使い方:
 *   syncFromMF(2025)  → 第7期（資金繰り表_2025）にMFデータを反映
 *   syncFromMF(2026)  → 第8期（資金繰り表_2026）にMFデータを反映
 *
 * メニューから実行:
 *   スプシ上部の「MF連携」メニューから同期可能
 */

function sync2025() { syncFromMF(2025); }
function sync2026() { syncFromMF(2026); }

// ==============================
// 設定: 期数と年度の対応
// ==============================

/** 決算年(シート名の年) → 期の開始年月・終了年月 */
function fiscalDates_(year) {
  // year=2025 → 第7期: 2024/3/1 〜 2025/2/28
  return {
    startDate: (year - 1) + '-03-01',
    endDate:   year + '-02-28',
    startYear: year - 1
  };
}

/** 月キーのリスト ["2024-03", "2024-04", ... "2025-02"] */
function monthKeys_(year) {
  var fy = fiscalDates_(year);
  var keys = [];
  for (var m = 3; m <= 12; m++) {
    keys.push(fy.startYear + '-' + padZero_(m));
  }
  keys.push(year + '-01');
  keys.push(year + '-02');
  return keys;
}

function padZero_(n) { return n < 10 ? '0' + n : '' + n; }

// ==============================
// MFデータ取得
// ==============================

/**
 * PL月次推移を取得
 * 返り値: { "売上高": {"2024-03": 12345, ...}, "広告宣伝費": {...}, ... }
 */
function fetchPL_(year) {
  var fy = fiscalDates_(year);
  var officeId = getOfficeId_();
  var data = mfApiGet_('/offices/' + officeId + '/trial_pl', {
    fiscal_year: fy.startYear,
    start_date: fy.startDate,
    end_date: fy.endDate
  });

  var result = {};
  var items = data.trial_pl || [];
  items.forEach(function(item) {
    var name = item.account_item_name || '';
    if (!result[name]) result[name] = {};
    (item.monthly || []).forEach(function(m) {
      result[name][m.year_month] = m.credit_amount || m.debit_amount || 0;
    });
  });
  return result;
}

/**
 * BS月次推移を取得
 * 返り値: { "現金": {"2024-03": 12345, ...}, ... }
 */
function fetchBS_(year) {
  var fy = fiscalDates_(year);
  var officeId = getOfficeId_();
  var data = mfApiGet_('/offices/' + officeId + '/trial_bs', {
    fiscal_year: fy.startYear,
    start_date: fy.startDate,
    end_date: fy.endDate
  });

  var result = {};
  var items = data.trial_bs || [];
  items.forEach(function(item) {
    var name = item.account_item_name || '';
    if (!result[name]) result[name] = {};
    (item.monthly || []).forEach(function(m) {
      result[name][m.year_month] = m.closing_balance || 0;
    });
  });
  return result;
}

/**
 * 仕訳データを取得
 */
function fetchJournals_(year) {
  var fy = fiscalDates_(year);
  var officeId = getOfficeId_();
  var allJournals = [];
  var page = 1;

  while (true) {
    var data = mfApiGet_('/offices/' + officeId + '/journals', {
      start_date: fy.startDate,
      end_date: fy.endDate,
      page: page,
      per_page: 500
    });
    var journals = data.journals || [];
    if (journals.length === 0) break;
    allJournals = allJournals.concat(journals);
    page++;
    if (page > 100) break; // 安全弁
  }
  return allJournals;
}

// ==============================
// 仕訳データの集計
// ==============================

/**
 * 仕訳から月別の売掛金回収額を集計
 * （貸方：売掛金、借方：普通預金 の仕訳）
 */
function calcARCollection_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    (j.details || []).forEach(function(d) {
      if (d.account_item_name === '売掛金' && (d.credit_amount || 0) > 0) {
        monthly[ym] = (monthly[ym] || 0) + d.credit_amount;
      }
    });
  });
  return monthly;
}

/**
 * 仕訳から月別の現金売上を集計
 * （借方：普通預金、貸方：売上高 で売掛金を介さない仕訳）
 */
function calcCashSales_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    var details = j.details || [];
    // 単一仕訳で借方が預金、貸方が売上高のパターン
    details.forEach(function(d) {
      if (d.account_item_name === '売上高' && (d.credit_amount || 0) > 0) {
        // 同じ仕訳に売掛金がなければ現金売上
        var hasSafeReceivable = details.some(function(d2) {
          return d2.account_item_name === '売掛金';
        });
        if (!hasSafeReceivable) {
          monthly[ym] = (monthly[ym] || 0) + d.credit_amount;
        }
      }
    });
  });
  return monthly;
}

/**
 * 仕訳から月別の買掛金支払を集計
 * （借方：未払金、貸方：普通預金 の仕訳合計 − 仕入値引き）
 */
function calcAPPayment_(journals) {
  var monthlyPayment = {};
  var monthlyDiscount = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    (j.details || []).forEach(function(d) {
      // 未払金の支払（借方：未払金）
      if (d.account_item_name === '未払金' && (d.debit_amount || 0) > 0) {
        monthlyPayment[ym] = (monthlyPayment[ym] || 0) + d.debit_amount;
      }
      // 仕入値引き（貸方：仕入値引）
      if (d.account_item_name === '仕入値引' && (d.credit_amount || 0) > 0) {
        monthlyDiscount[ym] = (monthlyDiscount[ym] || 0) + d.credit_amount;
      }
    });
  });

  // 未払金支払 − 仕入値引き
  var result = {};
  var allYMs = Object.keys(monthlyPayment).concat(Object.keys(monthlyDiscount));
  allYMs.forEach(function(ym) {
    result[ym] = (monthlyPayment[ym] || 0) - (monthlyDiscount[ym] || 0);
  });
  return result;
}

/**
 * 仕訳から月別の現金仕入を集計
 * （借方：仕入、貸方：普通預金 の仕訳）
 */
function calcCashPurchase_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    var details = j.details || [];
    details.forEach(function(d) {
      if (d.account_item_name === '仕入高' && (d.debit_amount || 0) > 0) {
        // 同じ仕訳に未払金がなければ現金仕入
        var hasPayable = details.some(function(d2) {
          return d2.account_item_name === '未払金';
        });
        if (!hasPayable) {
          monthly[ym] = (monthly[ym] || 0) + d.debit_amount;
        }
      }
    });
  });
  return monthly;
}

/**
 * 仕訳から月別の借入金返済（元本）を集計
 */
function calcLoanRepayment_(journals) {
  var shortTerm = {};
  var longTerm = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    (j.details || []).forEach(function(d) {
      if (d.account_item_name === '短期借入金' && (d.debit_amount || 0) > 0) {
        shortTerm[ym] = (shortTerm[ym] || 0) + d.debit_amount;
      }
      if (d.account_item_name === '長期借入金' && (d.debit_amount || 0) > 0) {
        longTerm[ym] = (longTerm[ym] || 0) + d.debit_amount;
      }
    });
  });
  return { short: shortTerm, long: longTerm };
}

// ==============================
// メイン同期処理
// ==============================

/**
 * MFデータを取得してスプシに反映
 * @param {number} year - 決算年（シート名の年）
 */
function syncFromMF(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('資金繰り表_' + year);
  if (!sheet) {
    throw new Error('資金繰り表_' + year + ' シートが見つかりません');
  }

  var keys = monthKeys_(year);
  Logger.log('MFデータ取得中... (' + year + ')');

  // --- データ取得 ---
  var pl = fetchPL_(year);
  var bs = fetchBS_(year);
  var journals = fetchJournals_(year);

  // --- 仕訳ベースの集計 ---
  var arCollection = calcARCollection_(journals);
  var cashSales = calcCashSales_(journals);
  var apPayment = calcAPPayment_(journals);
  var cashPurchase = calcCashPurchase_(journals);
  var loanRepay = calcLoanRepayment_(journals);

  // --- 千円単位に変換するヘルパー ---
  function toK(val) { return Math.round((val || 0) / 1000); }

  function monthlyToK(data, keys) {
    return keys.map(function(k) { return toK(data[k]); });
  }

  function plSum(accountNames, keys) {
    return keys.map(function(k) {
      var total = 0;
      accountNames.forEach(function(name) {
        if (pl[name] && pl[name][k]) total += pl[name][k];
      });
      return toK(total);
    });
  }

  // --- BS残高から前月繰越金を算出 ---
  // 前月繰越金 = 現金及び預金 + 売掛金（前月末残高）
  var cashAccounts = ['現金', '普通預金', '当座預金', '定期預金'];
  var prevMonthEnd = (year - 1) + '-02'; // 前期末
  var allKeys = [prevMonthEnd].concat(keys);

  var carryForward = [];
  for (var i = 0; i < keys.length; i++) {
    var prevKey = (i === 0) ? prevMonthEnd : keys[i - 1];
    var cashBal = 0;
    cashAccounts.forEach(function(acct) {
      if (bs[acct] && bs[acct][prevKey]) cashBal += bs[acct][prevKey];
    });
    var arBal = (bs['売掛金'] && bs['売掛金'][prevKey]) || 0;
    carryForward.push(toK(cashBal + arBal));
  }

  // --- スプシに書き込み ---
  // 行番号は新レイアウトに対応
  // C〜N列 = 3月〜2月

  // 行3: 売上（今期）
  var salesData = plSum(['売上高'], keys);
  sheet.getRange('C3:N3').setValues([salesData]);

  // 行5: 前月繰越金（3月のみ手入力値を設定、4月以降は数式）
  sheet.getRange('C5').setValue(carryForward[0]);

  // 行8: 現金売上
  sheet.getRange('C8:N8').setValues([monthlyToK(cashSales, keys)]);

  // 行9: 売掛金回収
  sheet.getRange('C9:N9').setValues([monthlyToK(arCollection, keys)]);

  // 行11: 現金仕入
  sheet.getRange('C11:N11').setValues([monthlyToK(cashPurchase, keys)]);

  // 行12: 買掛金支払（未払金 − 仕入値引き）
  sheet.getRange('C12:N12').setValues([monthlyToK(apPayment, keys)]);

  // 行13: 人件費
  var personnelData = plSum(['給料手当', '賞与', '法定福利費'], keys);
  sheet.getRange('C13:N13').setValues([personnelData]);

  // 行14: 商品棚卸高（期首と期末の差）
  var inventoryChange = keys.map(function(k, i) {
    var prevKey = (i === 0) ? prevMonthEnd : keys[i - 1];
    var prevInv = (bs['商品'] && bs['商品'][prevKey]) || 0;
    var curInv = (bs['商品'] && bs['商品'][k]) || 0;
    return toK(curInv - prevInv);
  });
  sheet.getRange('C14:N14').setValues([inventoryChange]);

  // 行15: 諸経費（販売費及び一般管理費 − 人件費関連）
  // PLの販管費合計から人件費科目を除いた残りを諸経費とする
  var miscExpenses = keys.map(function(k, i) {
    var totalExpense = 0;
    // 販管費に含まれる全科目を合計（売上原価・人件費を除く）
    var excludeAccounts = ['売上高', '売上原価', '仕入高', '給料手当', '賞与', '法定福利費', '商品'];
    Object.keys(pl).forEach(function(acctName) {
      if (excludeAccounts.indexOf(acctName) === -1 && pl[acctName][k]) {
        totalExpense += pl[acctName][k];
      }
    });
    return toK(totalExpense);
  });
  sheet.getRange('C15:N15').setValues([miscExpenses]);

  // 行18: 経常外収入
  var nonOpIncomeData = plSum(['受取利息', '雑収入', '受取配当金'], keys);
  sheet.getRange('C18:N18').setValues([nonOpIncomeData]);

  // 行19: 経常外支出
  var nonOpExpenseData = plSum(['支払利息', '雑損失'], keys);
  sheet.getRange('C19:N19').setValues([nonOpExpenseData]);

  // 行26: 借入金返済（短期）
  sheet.getRange('C26:N26').setValues([monthlyToK(loanRepay.short, keys)]);

  // 行27: 借入金返済（長期）
  sheet.getRange('C27:N27').setValues([monthlyToK(loanRepay.long, keys)]);

  // ※ 行23,24（公庫・信金の借入）は手入力のまま（融資実行時のみ）

  SpreadsheetApp.flush();
  Logger.log('同期完了: 資金繰り表_' + year);
  SpreadsheetApp.getActiveSpreadsheet().toast('MFデータの同期が完了しました', '資金繰り表_' + year);
}

/**
 * 前期の売上データを取得して行4に反映
 * syncFromMF() とは別に実行（前期データの取得が必要なため）
 */
function syncPrevYearSales(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('資金繰り表_' + year);
  if (!sheet) {
    throw new Error('資金繰り表_' + year + ' シートが見つかりません');
  }

  var prevYear = year - 1;
  var keys = monthKeys_(prevYear);
  var pl = fetchPL_(prevYear);

  function toK(val) { return Math.round((val || 0) / 1000); }

  var prevSales = keys.map(function(k) {
    return toK((pl['売上高'] && pl['売上高'][k]) || 0);
  });

  sheet.getRange('C4:N4').setValues([prevSales]);
  Logger.log('前期売上の同期完了');
}
