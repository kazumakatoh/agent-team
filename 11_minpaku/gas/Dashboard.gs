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

    const totalCleaning = res.cleaningFee + cost.cleaning;
    const totalCosts    = totalCleaning + cost.supplies + cost.utilities + cost.rent + cost.other;
    const netRevenue    = res.revenue - res.commission;
    const profit        = netRevenue - totalCosts;

    const kpis = KPICalculator.calcMonthlyKPIs({
      revenue: res.revenue, commission: res.commission,
      usageDays: res.usageDays, daysInMonth, totalCosts
    });

    return {
      label:         `${year}/${String(month).padStart(2, '0')}`,
      bookingCount:  res.bookingCount,
      usageDays:     res.usageDays,
      guests:        res.guests,
      revenue:       res.revenue,
      commission:    res.commission,
      cleaningFee:   totalCleaning,
      supplies:      cost.supplies,
      utilities:     cost.utilities,
      rent:          cost.rent,
      totalCosts,
      profit,
      roi:           kpis.roi,
      adr:           kpis.adr,
      revpar:        kpis.revpar,
      occupancyRate: kpis.occupancyRate,
      daysInMonth
    };
  });

  const annualKPIs = KPICalculator.calcAnnualKPIs(
    monthlyRows.map(r => ({
      revenue: r.revenue, commission: r.commission,
      usageDays: r.usageDays, totalCosts: r.totalCosts, daysInMonth: r.daysInMonth
    }))
  );

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
  sheet.setColumnWidths(1, 20, 120);
  sheet.setColumnWidth(1, 100);

  Logger.log(`ダッシュボード更新完了 (${fiscalYear}年度)`);
}

// ==============================
// 各セクションのレンダリング関数
// ==============================

function renderTitle_(sheet, fiscalYear, kpis, startRow) {
  const title = `🏠 ${CONFIG.PROPERTY.NAME} 民泊事業レポート ${fiscalYear}年度（${fiscalYear}/04 〜 ${fiscalYear + 1}/03）`;
  sheet.getRange(startRow, 1, 1, 12).merge()
       .setValue(title)
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontSize(14)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 40);

  const updatedAt = `最終更新: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}`;
  sheet.getRange(startRow + 1, 1, 1, 12).merge()
       .setValue(updatedAt)
       .setFontColor('#5f6368')
       .setHorizontalAlignment('right')
       .setFontSize(10);

  return startRow + 2;
}

function renderAnnualKPICards_(sheet, kpis, startRow) {
  // KPIカード定義
  const cards = [
    { label: '年間売上',         value: `¥${kpis.revenue.toLocaleString()}`,          bg: '#e8f5e9', textBg: '#1e8e3e' },
    { label: '年間利益',         value: `¥${kpis.profit.toLocaleString()}`,            bg: kpis.profit >= 0 ? '#e8f5e9' : '#fce8e6', textBg: kpis.profit >= 0 ? '#1e8e3e' : '#d93025' },
    { label: '稼働日数 / 180日', value: `${kpis.usageDays}日 (${kpis.legalOccupancyRate}%)`, bg: '#e3f2fd', textBg: '#1565c0' },
    { label: '稼働率（暦日）',   value: `${kpis.occupancyRate}%`,                      bg: '#e3f2fd', textBg: '#1565c0' },
    { label: 'ADR',              value: `¥${kpis.adr.toLocaleString()}`,               bg: '#fff3e0', textBg: '#e65100' },
    { label: 'RevPAR',           value: `¥${kpis.revpar.toLocaleString()}`,            bg: '#fff3e0', textBg: '#e65100' },
    { label: 'ROI',              value: `${kpis.roi}%`,                                bg: kpis.roi >= 0 ? '#e8f5e9' : '#fce8e6', textBg: kpis.roi >= 0 ? '#1e8e3e' : '#d93025' },
    { label: '利用件数',         value: `${kpis.bookingCount || '-'}件`,               bg: '#f3e5f5', textBg: '#6a1b9a' }
  ];

  // 見出し
  sheet.getRange(startRow, 1, 1, 12).merge()
       .setValue('■ 年間KPIサマリー')
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#f1f3f4');
  startRow++;

  // カードを4列×2行で表示
  cards.forEach((card, i) => {
    const col = (i % 4) * 3 + 1;
    const row = startRow + Math.floor(i / 4) * 3;

    // ラベル
    sheet.getRange(row, col, 1, 3).merge()
         .setValue(card.label)
         .setBackground(card.bg)
         .setFontColor(card.textBg)
         .setFontSize(10)
         .setHorizontalAlignment('center')
         .setFontWeight('bold');

    // 値
    sheet.getRange(row + 1, col, 1, 3).merge()
         .setValue(card.value)
         .setBackground(card.bg)
         .setFontColor(card.textBg)
         .setFontSize(16)
         .setFontWeight('bold')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(row + 1, 40);

    // 区切り線代わりのスペース
    sheet.getRange(row + 2, col, 1, 3).merge().setBackground(card.bg);
  });

  return startRow + 6 + 1;
}

function renderMonthlyTable_(sheet, monthlyRows, startRow) {
  // 見出し
  sheet.getRange(startRow, 1, 1, 19).merge()
       .setValue('■ 月別集計')
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground('#f1f3f4');
  startRow++;

  // ヘッダー
  const headers = [
    '年月', '利用件数', '稼働日数', '人数',
    '売上', '手数料', '清掃費', '備品・消耗品', '水光熱費', '家賃',
    '総経費', '利益', 'ROI(%)', 'ADR', 'RevPAR', '稼働率(%)'
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
    r.revenue,
    r.commission,
    r.cleaningFee,
    r.supplies,
    r.utilities,
    r.rent,
    r.totalCosts,
    r.profit,
    r.roi,
    r.adr,
    r.revpar,
    r.occupancyRate
  ]);

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

    // フォーマット
    const moneyRange = sheet.getRange(startRow, 5, rows.length, 8); // 売上〜利益
    moneyRange.setNumberFormat('¥#,##0');

    sheet.getRange(startRow, 13, rows.length, 1).setNumberFormat('0.0"%"'); // ROI
    sheet.getRange(startRow, 14, rows.length, 1).setNumberFormat('¥#,##0'); // ADR
    sheet.getRange(startRow, 15, rows.length, 1).setNumberFormat('¥#,##0'); // RevPAR
    sheet.getRange(startRow, 16, rows.length, 1).setNumberFormat('0.0"%"'); // 稼働率

    // 交互に色付け
    rows.forEach((_, i) => {
      const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      sheet.getRange(startRow + i, 1, 1, headers.length).setBackground(bg);

      // 利益がマイナスの場合は赤にハイライト
      if (rows[i][11] < 0) {
        sheet.getRange(startRow + i, 12).setFontColor('#d93025').setFontWeight('bold');
      }
    });
  }

  // 合計行
  const totalRow = startRow + rows.length;
  const totals   = rows.reduce((acc, r) => {
    for (let c = 1; c < r.length - 4; c++) { // KPI列は除く
      acc[c] = (acc[c] || 0) + (Number(r[c]) || 0);
    }
    return acc;
  }, ['合計']);
  while (totals.length < headers.length) totals.push('');

  sheet.getRange(totalRow, 1, 1, headers.length)
       .setValues([totals])
       .setBackground('#fce8e6')
       .setFontWeight('bold');

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
  const revenueChart = ss.newChart()
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
  const occupancyChart = ss.newChart()
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
