/**
 * 民泊自動化システム - ダッシュボード生成モジュール
 * 月別KPI・年間集計・グラフを自動生成する
 */

/**
 * ダッシュボードシートを全面更新する
 * @param {number} fiscalYear - 事業年度（開始年）
 */
function updateDashboard(fiscalYear) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEETS.DASHBOARD}」が見つかりません。`);

  // 月別データを取得
  const months = [];
  for (let m = 4; m <= 12; m++) months.push({ year: fiscalYear, month: m });
  for (let m = 1; m <= 3;  m++) months.push({ year: fiscalYear + 1, month: m });

  const monthlyRows = months.map(({ year, month }) => {
    const res   = getMonthlyReservationData(year, month);
    const cost  = getMonthlyCostData(year, month);
    const daysInMonth = new Date(year, month, 0).getDate();

    // 売上 = 宿泊料 + 清掃料（ゲスト実質負担）
    const grossRevenue  = res.accommodationFee + res.cleaningFee;
    const variableCosts = res.otaFee + res.transferFee
                        + (cost.agencyFee || 0) + cost.cleaning + (cost.linen || 0) + cost.supplies;
    const fixedCosts    = cost.utilities + cost.rent + cost.other;
    const profit        = grossRevenue - variableCosts - fixedCosts;
    const profitRate    = grossRevenue > 0 ? Math.round(profit / grossRevenue * 1000) / 10 : 0;

    const adr           = KPICalculator.calcADR(grossRevenue, res.usageDays);
    const revpar        = KPICalculator.calcRevPAR(grossRevenue, daysInMonth);
    const occupancyRate = KPICalculator.calcOccupancyRate(res.usageDays, daysInMonth);

    return {
      label:            `${year}/${String(month).padStart(2, '0')}`,
      bookingCount:     res.bookingCount,
      usageDays:        res.usageDays,
      guests:           res.guests,
      revenue:          res.revenue,
      accommodationFee: res.accommodationFee,
      cleaningFee:      res.cleaningFee,
      otaFee:           res.otaFee,
      transferFee:      res.transferFee,
      payout:           res.payout,
      agencyFee:        cost.agencyFee || 0,
      cleaning:         cost.cleaning,
      linen:            cost.linen || 0,
      supplies:         cost.supplies,
      utilities:        cost.utilities,
      rent:             cost.rent,
      other:            cost.other,
      grossRevenue,
      variableCosts,
      fixedCosts,
      profit,
      profitRate,
      adr,
      revpar,
      occupancyRate,
      daysInMonth
    };
  });

  const annualKPIs = KPICalculator.calcAnnualKPIs(monthlyRows);

  sheet.clearContents();
  sheet.clearFormats();

  let currentRow = 1;

  // ──────────────────────────────────────
  // ① タイトル & 年度情報
  // ──────────────────────────────────────
  currentRow = renderTitle_(sheet, fiscalYear, annualKPIs, currentRow);
  currentRow++;

  // ──────────────────────────────────────
  // ② KPIサマリーカード（年間）
  // ──────────────────────────────────────
  currentRow = renderAnnualKPICards_(sheet, annualKPIs, currentRow);
  currentRow++;

  // ──────────────────────────────────────
  // ③ 月別集計テーブル
  // ──────────────────────────────────────
  currentRow = renderMonthlyTable_(sheet, monthlyRows, currentRow);
  currentRow += 2;

  // ──────────────────────────────────────
  // ④ グラフ（売上・稼働率・利益）
  // ──────────────────────────────────────
  renderCharts_(sheet, ss, monthlyRows, currentRow);

  // 列幅調整
  sheet.setColumnWidths(1, 16, 110);
  sheet.setColumnWidth(1, 100);   // 年月
  sheet.setColumnWidth(10, 80);   // 利益率
  sheet.setColumnWidth(13, 80);   // 稼働率

  Logger.log(`ダッシュボード更新完了 (${fiscalYear}年度)`);
}

// ==============================
// 各セクションのレンダリング関数
// ==============================

function renderTitle_(sheet, fiscalYear, kpis, startRow) {
  const title = `🏠 ${CONFIG.PROPERTY.NAME} 民泊事業レポート ${fiscalYear}年度（${fiscalYear}/04 〜 ${fiscalYear + 1}/03）`;
  sheet.getRange(startRow, 1, 1, 16).merge()
       .setValue(title)
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontSize(14)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 40);

  const updatedAt = `最終更新: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}`;
  sheet.getRange(startRow + 1, 1, 1, 16).merge()
       .setValue(updatedAt)
       .setFontColor('#5f6368')
       .setHorizontalAlignment('right')
       .setFontSize(10);

  return startRow + 2;
}

function renderAnnualKPICards_(sheet, kpis, startRow) {
  // 見出し
  sheet.getRange(startRow, 1, 1, 16).merge()
       .setValue('■ 年間KPIサマリー')
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#f1f3f4');
  startRow++;

  const grossRev   = kpis.grossRevenue || 0;
  const varCosts   = kpis.variableCosts || 0;
  const fixCosts   = kpis.fixedCosts || 0;
  const profit     = kpis.profit || 0;
  const pctOf      = (v) => grossRev > 0 ? (Math.round(v / grossRev * 1000) / 10) + '%' : '-';

  // ──────────────────────────
  // ROW A: 財務4指標（金額＋売上比率）
  // 各カードは4列幅: COL1-4 / COL5-8 / COL9-12 / COL13-16
  //   左3列: 金額、右1列: 比率バッジ
  // ──────────────────────────
  const finCards = [
    { label: '年間売上（宿泊料＋清掃料）', amount: grossRev,  ratio: '100%',         bg: '#e8f5e9', fg: '#1e8e3e' },
    { label: '流動費',         amount: varCosts,  ratio: pctOf(varCosts), bg: '#fff8e1', fg: '#f57f17' },
    { label: '固定費',         amount: fixCosts,  ratio: pctOf(fixCosts), bg: '#fce4ec', fg: '#880e4f' },
    { label: '利益',           amount: profit,    ratio: pctOf(profit),   bg: profit >= 0 ? '#e8f5e9' : '#fce8e6', fg: profit >= 0 ? '#1e8e3e' : '#d93025' }
  ];

  finCards.forEach((card, i) => {
    const col = i * 4 + 1;
    const row = startRow;
    sheet.getRange(row, col, 1, 4).merge()
         .setValue(card.label)
         .setBackground(card.bg).setFontColor(card.fg)
         .setFontSize(10).setFontWeight('bold')
         .setHorizontalAlignment('center');
    // 金額（左3列）
    sheet.getRange(row + 1, col, 1, 3).merge()
         .setValue(`¥${card.amount.toLocaleString()}`)
         .setBackground(card.bg).setFontColor(card.fg)
         .setFontSize(15).setFontWeight('bold')
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 比率バッジ（右1列）
    sheet.getRange(row + 1, col + 3, 1, 1)
         .setValue(card.ratio)
         .setBackground(card.bg).setFontColor(card.fg)
         .setFontSize(11).setFontWeight('bold')
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(row + 1, 42);
    sheet.getRange(row + 2, col, 1, 4).merge().setBackground(card.bg);
  });
  startRow += 3;

  // ──────────────────────────
  // ROW B: 稼働・運営KPI（4カード × 4列）
  // ──────────────────────────
  const opsCards = [
    { label: '稼働率（/365日）',  value: `${kpis.occupancyRate365}%`,      bg: '#e3f2fd', fg: '#1565c0' },
    { label: '稼働率（/180日）',  value: `${kpis.legalOccupancyRate}%\n(${kpis.usageDays}日/180日)`, bg: '#e3f2fd', fg: '#1565c0' },
    { label: '利用件数',          value: `${kpis.bookingCount || 0}件`,     bg: '#f3e5f5', fg: '#6a1b9a' },
    { label: 'ADR',               value: `¥${(kpis.adr || 0).toLocaleString()}`, bg: '#fff3e0', fg: '#e65100' }
  ];
  const opsCards2 = [
    { label: 'RevPAR',            value: `¥${(kpis.revpar || 0).toLocaleString()}`, bg: '#fff3e0', fg: '#e65100' },
    { label: '初期投資額',        value: `¥${(kpis.initialInvestment || 4262824).toLocaleString()}`, bg: '#eceff1', fg: '#37474f' },
    { label: 'ROI（初期投資回収率）', value: `${kpis.roi}%`, bg: kpis.roi >= 0 ? '#e8f5e9' : '#fce8e6', fg: kpis.roi >= 0 ? '#1e8e3e' : '#d93025' },
    { label: '利益率',            value: `${kpis.profitRate || 0}%`,        bg: profit >= 0 ? '#e8f5e9' : '#fce8e6', fg: profit >= 0 ? '#1e8e3e' : '#d93025' }
  ];

  [opsCards, opsCards2].forEach(cardRow => {
    cardRow.forEach((card, i) => {
      const col = i * 4 + 1;
      const row = startRow;
      sheet.getRange(row, col, 1, 4).merge()
           .setValue(card.label)
           .setBackground(card.bg).setFontColor(card.fg)
           .setFontSize(10).setFontWeight('bold')
           .setHorizontalAlignment('center');
      sheet.getRange(row + 1, col, 1, 4).merge()
           .setValue(card.value)
           .setBackground(card.bg).setFontColor(card.fg)
           .setFontSize(14).setFontWeight('bold')
           .setHorizontalAlignment('center').setVerticalAlignment('middle')
           .setWrap(true);
      sheet.setRowHeight(row + 1, 42);
      sheet.getRange(row + 2, col, 1, 4).merge().setBackground(card.bg);
    });
    startRow += 3;
  });

  return startRow + 1;
}

function renderMonthlyTable_(sheet, monthlyRows, startRow) {
  const NUM_COLS = 15;

  // 見出し
  sheet.getRange(startRow, 1, 1, NUM_COLS).merge()
       .setValue('■ 月別集計')
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#f1f3f4');
  startRow++;

  // ヘッダー（16列）
  const headers = [
    '年月', '利用件数', '稼働日数', '人数',
    '売上', '流動費', '固定費',
    '利益', '利益率(%)', 'ADR', 'RevPAR', '稼働率(%)',
    '代行手数料', '清掃費', 'リネン費'
  ];
  sheet.getRange(startRow, 1, 1, headers.length)
       .setValues([headers])
       .setBackground('#4285f4')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
  sheet.setFrozenRows(startRow);
  startRow++;

  // データ行
  const rows = monthlyRows.map(r => [
    r.label,
    r.bookingCount,
    r.usageDays,
    r.guests,
    r.grossRevenue,   // 売上 = 宿泊料+清掃料
    r.variableCosts,
    r.fixedCosts,
    r.profit,
    r.profitRate,
    r.adr,
    r.revpar,
    r.occupancyRate,
    r.agencyFee,
    r.cleaning,
    r.linen
  ]);

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

    // 金額・パーセント書式（15列構成）
    sheet.getRange(startRow, 5, rows.length, 4).setNumberFormat('¥#,##0');  // 売上〜利益（COL5-8）
    sheet.getRange(startRow, 9, rows.length, 1).setNumberFormat('0.0"%"');  // 利益率(%)
    sheet.getRange(startRow, 10, rows.length, 2).setNumberFormat('¥#,##0'); // ADR, RevPAR
    sheet.getRange(startRow, 12, rows.length, 1).setNumberFormat('0.0"%"'); // 稼働率
    sheet.getRange(startRow, 13, rows.length, 3).setNumberFormat('¥#,##0'); // 代行〜リネン

    // 交互に色付け
    rows.forEach((row, i) => {
      const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      sheet.getRange(startRow + i, 1, 1, headers.length).setBackground(bg);
      if (row[7] < 0) { // 利益がマイナス（COL8 = 利益）
        sheet.getRange(startRow + i, 8).setFontColor('#d93025').setFontWeight('bold');
      }
    });
  }

  // 合計行（金額列のみSUM）
  const totalRow = startRow + rows.length;
  const sumCols  = [2, 3, 4, 5, 6, 7, 8, 10, 11, 13, 14, 15];
  const totalData = new Array(headers.length).fill('');
  totalData[0] = '合計';
  sumCols.forEach(col => {
    totalData[col - 1] = rows.reduce((s, r) => s + (Number(r[col - 1]) || 0), 0);
  });

  sheet.getRange(totalRow, 1, 1, headers.length)
       .setValues([totalData])
       .setBackground('#fce8e6')
       .setFontWeight('bold');
  sheet.getRange(totalRow, 5, 1, 4).setNumberFormat('¥#,##0');  // 売上〜利益
  sheet.getRange(totalRow, 10, 1, 2).setNumberFormat('¥#,##0'); // ADR, RevPAR
  sheet.getRange(totalRow, 13, 1, 3).setNumberFormat('¥#,##0'); // 代行〜リネン

  return totalRow + 1;
}

function renderCharts_(sheet, ss, monthlyRows, dataStartRow) {
  // 既存グラフを削除
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  if (monthlyRows.length === 0) return;

  // グラフ用データテーブルを別の場所に配置
  const chartDataRow = dataStartRow + 2;
  const labels       = monthlyRows.map(r => r.label);
  const revenues     = monthlyRows.map(r => r.revenue);
  const profits      = monthlyRows.map(r => r.profit);
  const occupancies  = monthlyRows.map(r => r.occupancyRate);

  // チャートデータを書き込み（非表示にする）
  sheet.getRange(chartDataRow, 1).setValue('グラフデータ（自動生成）');
  sheet.getRange(chartDataRow + 1, 1, 1, labels.length + 1)
       .setValues([['項目', ...labels]]);
  sheet.getRange(chartDataRow + 2, 1, 1, revenues.length + 1)
       .setValues([['売上', ...revenues]]);
  sheet.getRange(chartDataRow + 3, 1, 1, profits.length + 1)
       .setValues([['利益', ...profits]]);
  sheet.getRange(chartDataRow + 4, 1, 1, occupancies.length + 1)
       .setValues([['稼働率(%)', ...occupancies]]);

  // 売上・利益グラフ
  const revenueChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(chartDataRow + 1, 1, 3, labels.length + 1))
    .setPosition(chartDataRow + 6, 1, 0, 0)
    .setOption('title', '月別 売上・利益')
    .setOption('width', 600)
    .setOption('height', 300)
    .setOption('colors', ['#4285f4', '#34a853'])
    .setOption('vAxis.format', '¥#,##0')
    .build();
  sheet.insertChart(revenueChart);

  // 稼働率グラフ
  const occupancyChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(chartDataRow + 1, 1, 1, labels.length + 1))
    .addRange(sheet.getRange(chartDataRow + 4, 1, 1, labels.length + 1))
    .setPosition(chartDataRow + 6, 7, 0, 0)
    .setOption('title', '月別 稼働率')
    .setOption('width', 400)
    .setOption('height', 300)
    .setOption('colors', ['#ea4335'])
    .setOption('vAxis.format', '0"%"')
    .setOption('vAxis.maxValue', 100)
    .build();
  sheet.insertChart(occupancyChart);
}
