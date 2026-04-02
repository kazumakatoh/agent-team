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
   *
   * ★ MF会計 APIのレスポンス構造について（ドキュメントで要確認）
   *   各アイテムの想定フィールド:
   *   - item.name または item.account_item.name : 勘定科目名
   *   - item.amount または item.total_amount     : 期間合計金額
   *   - item.account_category                    : 科目カテゴリ
   *   - item.financial_statement_type             : 'pl' | 'bs'
   *
   * MF会計 API v3 試算表レスポンス構造（確認済み）:
   *   columns: ["opening_balance","debit_amount","credit_amount","closing_balance","ratio"]
   *   rows（ネスト構造）:
   *     { name:"売上高合計", type:"financial_statement_item", values:[...], rows:[
   *       { name:"売上（国内）", type:"account", values:[0,0,15524254,15524254,110.1], rows:null },
   *       ...
   *     ]}
   *   → type:"account" の葉ノードのみ収集し、values[3](closing_balance)を金額として使用
   */
  _buildAccountMap(trialBalanceRows) {
    const map = {};

    // ネスト構造を再帰的に辿り、type:"account" の葉ノードを収集
    function collectAccounts(rows) {
      if (!rows) return;
      rows.forEach(row => {
        if (row.type === 'account' && row.name) {
          // values[0] = opening_balance（期首残高 = 前月までの累計）
          // values[3] = closing_balance（期末残高 = 当月末までの累計）
          // 月次金額 = closing_balance - opening_balance（当月の発生額のみ）
          const closing = (row.values && row.values[3]) || 0;
          const opening = (row.values && row.values[0]) || 0;
          const amount = Math.abs(closing - opening);
          map[row.name] = (map[row.name] || 0) + amount;
        }
        if (row.rows) collectAccounts(row.rows);
      });
    }

    collectAccounts(trialBalanceRows);
    Logger.log(`勘定科目マップ: ${Object.keys(map).length}科目 → ${JSON.stringify(Object.keys(map).slice(0, 5))} ...`);
    return map;
  },
};
