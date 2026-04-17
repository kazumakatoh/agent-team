/**
 * Amazon Dashboard - M3 販促費マスター管理（レイヤー2: 最終利益計算）
 *
 * ## M3 シート構造（社長の手入力は4列のみ）
 *
 * | 年月    | Amence | その他ツール | 荷造運賃 | 納品人件費 | 備考 |
 * |---------|--------|--------------|---------|-----------|------|
 * | 2026-04 | 44,000 | 11,300       | (手入力)| (手入力)  |      |
 *
 * - Amence: 44,000円/月（業務委託費）
 * - その他ツール: 11,300円/月（Keepa 4,300円 + M19 7,000円）
 * - 荷造運賃: ヤマト運輸配送料（変動・手入力）
 * - 納品人件費: ママ友バイト実績（変動・手入力）
 * - 備考: メモ欄
 *
 * ## M3 に載せない項目（Dashboard 側で自動計算）
 *
 * - M19 成果報酬 = 月間広告費 × 6.0%
 *
 * ## 最終利益（レイヤー2）の計算式
 *
 * 最終利益 = Amazon内粗利
 *   − Amence
 *   − その他ツール
 *   − 荷造運賃
 *   − 納品人件費
 *   − (広告費 × 6.0%)
 *
 * ## 期間が部分月の場合（例: 当月1日〜今日）
 *
 * 固定費（Amence / その他ツール）は日数按分:
 *   Amence × (経過日数 / 月日数)
 *
 * 変動費（荷造運賃 / 納品人件費）は手入力値そのまま × 按分比
 * 成果報酬は広告費自体が既に按分されているのでそのまま × 6.0%
 */

const M19_PERFORMANCE_RATE = 0.06; // M19 成果報酬率（広告費対比）
const DEFAULT_AMENCE = 44000;      // Amence 業務委託費（毎月固定）
const DEFAULT_OTHER_TOOLS = 11300; // その他ツール（Keepa 4,300 + M19 7,000）

const M3_COLS = 6;
const M3_HEADERS = ['年月', 'Amence', 'その他ツール', '荷造運賃', '納品人件費', '備考'];

/**
 * M3 販促費マスターシートを作成・初期化
 * 現在月から遡って12ヶ月分、Amence/その他ツールを自動プリフィル
 */
function setupPromoCostSheet() {
  const ss = getMainSpreadsheet();
  const sheet = getOrCreateSheet(SHEET_NAMES.M3_PROMO_COST);

  // ヘッダー
  sheet.getRange(1, 1, 1, M3_COLS).setValues([M3_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 既存データの年月を収集
  const existingYMs = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existing.forEach(r => {
      const ym = formatYearMonth(r[0]);
      if (ym) existingYMs.add(ym);
    });
  }

  // 現在月から12ヶ月分遡って、未登録の月だけ追加
  const today = new Date();
  const newRows = [];
  for (let i = 0; i < 12; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    if (existingYMs.has(ym)) continue;
    newRows.push([ym, DEFAULT_AMENCE, DEFAULT_OTHER_TOOLS, 0, 0, '']);
  }

  // 昇順（古い順）で並べて挿入
  newRows.sort((a, b) => a[0].localeCompare(b[0]));

  if (newRows.length > 0) {
    const writeRow = Math.max(lastRow + 1, 2);
    sheet.getRange(writeRow, 1, newRows.length, M3_COLS).setValues(newRows);
    // 金額列のフォーマット
    sheet.getRange(writeRow, 2, newRows.length, 4).setNumberFormat('#,##0');
    Logger.log('✅ M3 販促費マスター: ' + newRows.length + ' ヶ月分を追加');
  } else {
    Logger.log('M3 販促費マスター: 12ヶ月分すべて登録済み');
  }

  // 列幅調整
  sheet.setColumnWidth(1, 90);   // 年月
  sheet.setColumnWidth(2, 90);   // Amence
  sheet.setColumnWidth(3, 110);  // その他ツール
  sheet.setColumnWidth(4, 90);   // 荷造運賃
  sheet.setColumnWidth(5, 110);  // 納品人件費
  sheet.setColumnWidth(6, 240);  // 備考

  Logger.log('✅ M3 販促費マスター準備完了');
}

/**
 * M3 販促費マスターから全データ読み込み
 * @returns {Object} { 'YYYY-MM': { amence, otherTools, shipping, labor } }
 */
function getPromoCostByMonth() {
  const sheet = getOrCreateSheet(SHEET_NAMES.M3_PROMO_COST);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const map = {};
  for (const row of data) {
    const ym = formatYearMonth(row[0]);
    if (!ym) continue;
    map[ym] = {
      amence: parseFloat(row[1]) || 0,
      otherTools: parseFloat(row[2]) || 0,
      shipping: parseFloat(row[3]) || 0,
      labor: parseFloat(row[4]) || 0,
    };
  }
  return map;
}

/**
 * 指定期間の販促費合計を計算（部分月は日数按分）
 *
 * @param {string} startDate 'YYYY-MM-DD'
 * @param {string} endDate   'YYYY-MM-DD'
 * @param {number} totalAdCost 期間内の実広告費合計（既に按分済みの値）
 * @returns {Object} {
 *   amence, otherTools, shipping, labor, m19Performance,
 *   total,
 *   breakdown: [{ ym, amence, ... }]  // 月別内訳
 * }
 */
function calcPromoCostForPeriod(startDate, endDate, totalAdCost) {
  const map = getPromoCostByMonth();

  const result = {
    amence: 0,
    otherTools: 0,
    shipping: 0,
    labor: 0,
    m19Performance: (totalAdCost || 0) * M19_PERFORMANCE_RATE,
    breakdown: [],
  };

  // 開始〜終了月まで月ごとにループし、部分月は按分
  const [sy, sm, sd] = startDate.split('-').map(Number);
  const [ey, em, ed] = endDate.split('-').map(Number);

  let y = sy, mo = sm;
  while (y < ey || (y === ey && mo <= em)) {
    const ym = y + '-' + String(mo).padStart(2, '0');
    const row = map[ym];

    // その月の按分比を計算
    const daysInMonth = new Date(y, mo, 0).getDate();
    const monthStart = (y === sy && mo === sm) ? sd : 1;
    const monthEnd = (y === ey && mo === em) ? ed : daysInMonth;
    const daysInRange = monthEnd - monthStart + 1;
    const ratio = daysInRange / daysInMonth;

    if (row) {
      const amence = row.amence * ratio;
      const otherTools = row.otherTools * ratio;
      const shipping = row.shipping * ratio;
      const labor = row.labor * ratio;

      result.amence += amence;
      result.otherTools += otherTools;
      result.shipping += shipping;
      result.labor += labor;

      result.breakdown.push({ ym, amence, otherTools, shipping, labor, ratio });
    } else {
      result.breakdown.push({ ym, amence: 0, otherTools: 0, shipping: 0, labor: 0, ratio, missing: true });
    }

    // 次月へ
    mo++;
    if (mo > 12) { mo = 1; y++; }
  }

  result.total = result.amence + result.otherTools + result.shipping + result.labor + result.m19Performance;
  return result;
}

/**
 * 日付値を 'YYYY-MM' 形式に変換
 * Date / "YYYY-MM" / "YYYY-MM-DD" などに対応
 */
function formatYearMonth(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return value.getFullYear() + '-' + String(value.getMonth() + 1).padStart(2, '0');
  }
  const s = String(value).trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + m[2].padStart(2, '0');
}

/**
 * テスト: 現在月の販促費を計算して表示
 */
function testPromoCost() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const firstOfMonth = today.substring(0, 7) + '-01';

  // 仮の広告費（実際は Dashboard から渡される）
  const testAdCost = 500000;

  const result = calcPromoCostForPeriod(firstOfMonth, today, testAdCost);
  Logger.log('期間: ' + firstOfMonth + ' 〜 ' + today);
  Logger.log('広告費（仮）: ' + testAdCost);
  Logger.log('Amence: ' + Math.round(result.amence));
  Logger.log('その他ツール: ' + Math.round(result.otherTools));
  Logger.log('荷造運賃: ' + Math.round(result.shipping));
  Logger.log('納品人件費: ' + Math.round(result.labor));
  Logger.log('M19成果報酬: ' + Math.round(result.m19Performance));
  Logger.log('販促費合計: ' + Math.round(result.total));
}
