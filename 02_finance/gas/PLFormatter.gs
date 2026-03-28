/**
 * 財務レポート自動化システム - PL集計・フォーマットロジック
 *
 * MF会計の試算表データをPL構造定義に沿って集計・整形する
 */

const PLFormatter = {

  /**
   * MF試算表データからPL行データを生成する
   *
   * @param {Array} trialBalanceItems - MF APIの試算表アイテム配列
   * @return {Array} PL行データ配列（各行は {label, amount, category, ...} の形式）
   */
  buildPLRows(trialBalanceItems) {
    // MF勘定科目名 → 金額のマップを作成
    const accountMap = PLFormatter._buildAccountMap(trialBalanceItems);

    // 各カテゴリの小計を計算
    const categoryTotals = {
      revenue:     0,
      cogs:        0,
      sga:         0,
      nonOpIncome: 0,
      nonOpExpense: 0,
      extraIncome: 0,
      extraExpense: 0,
      tax:         0,
    };

    // PL行データを生成
    const rows = [];

    CONFIG.PL_STRUCTURE.forEach(item => {
      const row = {
        label:       item.label,
        category:    item.category,
        indent:      item.indent || 0,
        isBold:      item.isBold || false,
        isBorderTop: item.isBorderTop || false,
        amount:      null,
        isHeader:    item.category === 'header',
        isSubtotal:  item.category === 'subtotal',
      };

      switch (item.category) {

        case 'header':
          // ヘッダー行（金額なし）
          row.amount = null;
          break;

        case 'subtotal': {
          // 小計行
          const total = categoryTotals[item.calcFrom] || 0;
          row.amount  = total;
          break;
        }

        // ── 利益計算行 ──
        case 'grossProfit':
          row.amount = categoryTotals.revenue - categoryTotals.cogs;
          break;

        case 'operatingProfit':
          row.amount = (categoryTotals.revenue - categoryTotals.cogs) - categoryTotals.sga;
          break;

        case 'ordinaryProfit':
          row.amount =
            (categoryTotals.revenue - categoryTotals.cogs - categoryTotals.sga) +
            categoryTotals.nonOpIncome -
            categoryTotals.nonOpExpense;
          break;

        case 'pretaxProfit':
          row.amount =
            (categoryTotals.revenue - categoryTotals.cogs - categoryTotals.sga) +
            categoryTotals.nonOpIncome -
            categoryTotals.nonOpExpense +
            categoryTotals.extraIncome -
            categoryTotals.extraExpense;
          break;

        case 'netProfit':
          row.amount =
            (categoryTotals.revenue - categoryTotals.cogs - categoryTotals.sga) +
            categoryTotals.nonOpIncome -
            categoryTotals.nonOpExpense +
            categoryTotals.extraIncome -
            categoryTotals.extraExpense -
            categoryTotals.tax;
          break;

        default: {
          // 通常の勘定科目行（accountNamesでMFデータを照合）
          if (item.accountNames) {
            let amount = 0;
            item.accountNames.forEach(name => {
              if (accountMap[name] !== undefined) {
                amount += accountMap[name];
              }
            });
            // sign: -1 の項目はマイナス（値引・返品・期末棚卸など）
            const signedAmount = amount * (item.sign || 1);
            row.amount = signedAmount;

            // カテゴリ集計（絶対値で加算してsignをここで適用）
            if (categoryTotals[item.category] !== undefined) {
              categoryTotals[item.category] += signedAmount;
            }
          }
          break;
        }
      }

      rows.push(row);
    });

    return rows;
  },

  /**
   * 今日以前に開始している月（過去・当月）のみ返す
   * 未来月はAPIから取得しても意味がなく、シートの手入力予測値を守るためスキップする
   *
   * @param {Array} months - getFiscalMonths() の結果
   * @return {Array} 更新対象月リスト
   */
  getUpdatableMonths(months) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return months.filter(m => {
      const monthStart = new Date(m.startDate);
      return today >= monthStart;
    });
  },

  /**
   * 全部門の月別PLデータを取得・集計する
   *
   * @param {number} fiscalYear    - 事業年度開始年
   * @param {Array}  monthsToFetch - 取得する月リスト（省略時は全12ヶ月）
   * @return {Object} { deptName: { monthLabel: [PLrows], ... }, '全体': { ... } }
   */
  fetchAllDepartmentsPL(fiscalYear, monthsToFetch) {
    const months = monthsToFetch || getFiscalMonths(fiscalYear);
    const result = {};

    // 部門別PL取得
    CONFIG.DEPARTMENTS.forEach(dept => {
      Logger.log(`部門別PL取得中: ${dept.name}`);
      result[dept.name] = {};

      months.forEach(m => {
        try {
          const items = MFApiClient.getTrialBalance(m.startDate, m.endDate, dept.id || undefined);
          result[dept.name][m.label] = PLFormatter.buildPLRows(items);
          Logger.log(`  ${m.label}: ${items.length}件取得`);
        } catch (e) {
          Logger.log(`  ${m.label}: エラー - ${e.message}`);
          result[dept.name][m.label] = PLFormatter.buildPLRows([]);
        }
        Utilities.sleep(500); // API負荷軽減
      });
    });

    // 全体（部門指定なし）PL取得
    Logger.log('全体PL取得中');
    result['全体'] = {};
    months.forEach(m => {
      try {
        const items = MFApiClient.getTrialBalance(m.startDate, m.endDate);
        result['全体'][m.label] = PLFormatter.buildPLRows(items);
      } catch (e) {
        Logger.log(`  全体 ${m.label}: エラー - ${e.message}`);
        result['全体'][m.label] = PLFormatter.buildPLRows([]);
      }
      Utilities.sleep(500);
    });

    return result;
  },

  /**
   * 月別PLRowsから「月別推移用の行列データ」を生成する
   * 行: PL項目, 列: 月 + 合計
   *
   * @param {Object} monthlyRows - { monthLabel: [PLrows], ... }
   * @param {Array}  months      - getFiscalMonths()の結果
   * @return {Object} { headers: [...], rows: [[...]] }
   */
  buildMonthlyMatrix(monthlyRows, months) {
    // 列ヘッダー
    const headers = ['勘定科目', ...months.map(m => m.label), '決算整理', '合計'];

    // PL構造に沿ってデータ行を生成
    const rows = CONFIG.PL_STRUCTURE.map(item => {
      const labelCell = '　'.repeat(item.indent || 0) + item.label;
      const monthValues = months.map(m => {
        const monthRows = monthlyRows[m.label] || [];
        const row = monthRows.find(r => r.label === item.label);
        return (row && row.amount !== null) ? row.amount : 0;
      });

      // 合計（決算整理は0で仮置き）
      const total = monthValues.reduce((s, v) => s + v, 0);

      return {
        label:       item.label,
        labelCell,
        category:    item.category,
        isBold:      item.isBold || false,
        isBorderTop: item.isBorderTop || false,
        isHeader:    item.category === 'header',
        indent:      item.indent || 0,
        values:      [...monthValues, 0, total], // 決算整理=0, 合計
      };
    });

    return { headers, rows };
  },

  // ==============================
  // 内部ユーティリティ
  // ==============================

  /**
   * MF試算表アイテムから「勘定科目名 → 金額」マップを作成
   * PL項目（account_category が income / expense 系）を対象とする
   */
  _buildAccountMap(trialBalanceItems) {
    const map = {};

    trialBalanceItems.forEach(item => {
      const acct = item.account_item;
      if (!acct) return;

      const name = acct.name;
      // MF API では credit_amount がプラス（収益）、debit_amount がプラス（費用）
      // closing_balance は残高（PL科目では期間合計）
      // net_amount = credit_amount - debit_amount
      const netAmount = (item.credit_amount || 0) - (item.debit_amount || 0);

      // 収益科目（売上高など）: credit_amount > debit_amount → プラス
      // 費用科目（仕入高など）: debit_amount > credit_amount → マイナス → 絶対値を使う
      const category = acct.account_category || '';
      let amount;

      if (['sales', 'non_operating_income', 'extraordinary_income'].includes(category)) {
        // 収益科目はプラス
        amount = item.credit_amount || 0;
      } else if (['cost_of_sales', 'selling_general_and_administrative', 'non_operating_expense', 'extraordinary_expense', 'income_taxes'].includes(category)) {
        // 費用科目はプラス（PLFormatterのsignで符号を決める）
        amount = item.debit_amount || 0;
      } else {
        // その他（closing_balanceで代用）
        amount = Math.abs(item.closing_balance || 0);
      }

      // 同名科目が複数ある場合は加算
      map[name] = (map[name] || 0) + amount;
    });

    Logger.log(`勘定科目マップ作成: ${Object.keys(map).length}科目`);
    return map;
  },
};
