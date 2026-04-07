"""
MoneyForward クラウド会計API連携 - 資金繰り表自動生成スクリプト

日本政策金融公庫フォーマットに準拠した月別資金繰り表を
MoneyForward APIから取得したデータで自動生成する。

使い方:
  python mf_cashflow.py --period 8        # 第8期の資金繰り表を生成
  python mf_cashflow.py --period 8 --auth # 初回認証フロー
"""

import argparse
import csv
import json
import sys
import urllib.request
import urllib.parse
import urllib.error
from datetime import date
from pathlib import Path

BASE_URL = "https://accounting.moneyforward.com/api/v3"
CONFIG_PATH = Path(__file__).parent / "config.json"
OUTPUT_DIR = Path(__file__).parent

# 決算月=2月 → 第N期は (2017+N)年3月〜(2018+N)年2月
FOUNDING_YEAR = 2018  # 第1期開始年


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ──────────────────────────────────────────────
# OAuth2 認証
# ──────────────────────────────────────────────

def get_auth_url(config: dict) -> str:
    mf = config["moneyforward"]
    params = urllib.parse.urlencode({
        "client_id": mf["client_id"],
        "redirect_uri": mf["redirect_uri"],
        "response_type": "code",
        "scope": "office:read account:read journal:read report:read",
    })
    return f"https://accounting.moneyforward.com/oauth/authorize?{params}"


def exchange_code(config: dict, code: str) -> dict:
    """認可コードをアクセストークンに交換"""
    mf = config["moneyforward"]
    data = urllib.parse.urlencode({
        "client_id": mf["client_id"],
        "client_secret": mf["client_secret"],
        "redirect_uri": mf["redirect_uri"],
        "grant_type": "authorization_code",
        "code": code,
    }).encode()
    req = urllib.request.Request(
        "https://accounting.moneyforward.com/oauth/token",
        data=data,
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def refresh_access_token(config: dict) -> str:
    """リフレッシュトークンでアクセストークンを更新"""
    mf = config["moneyforward"]
    data = urllib.parse.urlencode({
        "client_id": mf["client_id"],
        "client_secret": mf["client_secret"],
        "grant_type": "refresh_token",
        "refresh_token": mf["refresh_token"],
    }).encode()
    req = urllib.request.Request(
        "https://accounting.moneyforward.com/oauth/token",
        data=data,
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        token_data = json.loads(resp.read())
    mf["access_token"] = token_data["access_token"]
    if "refresh_token" in token_data:
        mf["refresh_token"] = token_data["refresh_token"]
    save_config(config)
    return mf["access_token"]


def auth_flow(config: dict):
    """対話的にOAuth2認証を行う"""
    mf = config["moneyforward"]
    if not mf["client_id"] or not mf["client_secret"]:
        print("エラー: config.json に client_id / client_secret を設定してください。")
        print("MoneyForward クラウド会計 > API連携 > アプリ管理 から取得できます。")
        sys.exit(1)

    print("以下のURLをブラウザで開いて認証してください：")
    print(get_auth_url(config))
    print()
    code = input("認可コードを入力: ").strip()
    token_data = exchange_code(config, code)
    mf["access_token"] = token_data["access_token"]
    mf["refresh_token"] = token_data.get("refresh_token", "")
    save_config(config)
    print("認証成功。トークンを config.json に保存しました。")

    # office_id を取得
    if not mf["office_id"]:
        offices = api_get(config, "/offices")
        if offices:
            mf["office_id"] = offices[0]["id"]
            save_config(config)
            print(f"事業所ID: {mf['office_id']} を設定しました。")


# ──────────────────────────────────────────────
# API呼び出し
# ──────────────────────────────────────────────

def api_get(config: dict, endpoint: str, params: dict | None = None) -> dict:
    """MF APIへGETリクエスト"""
    mf = config["moneyforward"]
    url = BASE_URL + endpoint
    if params:
        url += "?" + urllib.parse.urlencode(params)

    headers = {"Authorization": f"Bearer {mf['access_token']}"}
    req = urllib.request.Request(url, headers=headers)

    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        if e.code == 401:
            # トークン期限切れ → リフレッシュして再試行
            print("トークン期限切れ。リフレッシュ中...")
            refresh_access_token(config)
            headers["Authorization"] = f"Bearer {config['moneyforward']['access_token']}"
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as resp:
                return json.loads(resp.read())
        raise


# ──────────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────────

def period_to_dates(period: int) -> tuple[str, str]:
    """期数から開始日・終了日を算出（決算月=2月）"""
    start_year = FOUNDING_YEAR + (period - 1)  # 第1期=2018年3月〜
    start = f"{start_year}-03-01"
    end = f"{start_year + 1}-02-28"
    return start, end


def fiscal_months(period: int) -> list[tuple[int, int]]:
    """期数から月のリストを返す [(year, month), ...]"""
    start_year = FOUNDING_YEAR + (period - 1)
    months = []
    for m in range(3, 13):
        months.append((start_year, m))
    for m in range(1, 3):
        months.append((start_year + 1, m))
    return months


def fetch_monthly_pl(config: dict, period: int) -> dict:
    """月次PL（損益計算書）の推移データを取得"""
    start, end = period_to_dates(period)
    office_id = config["moneyforward"]["office_id"]
    data = api_get(config, f"/offices/{office_id}/trial_pl", {
        "fiscal_year": start[:4],
        "start_date": start,
        "end_date": end,
    })
    return data


def fetch_monthly_bs(config: dict, period: int) -> dict:
    """月次BS（貸借対照表）の推移データを取得"""
    start, end = period_to_dates(period)
    office_id = config["moneyforward"]["office_id"]
    data = api_get(config, f"/offices/{office_id}/trial_bs", {
        "fiscal_year": start[:4],
        "start_date": start,
        "end_date": end,
    })
    return data


def fetch_journals(config: dict, period: int) -> list:
    """仕訳データを取得（借入・返済の特定用）"""
    start, end = period_to_dates(period)
    office_id = config["moneyforward"]["office_id"]
    all_journals = []
    page = 1
    while True:
        data = api_get(config, f"/offices/{office_id}/journals", {
            "start_date": start,
            "end_date": end,
            "page": page,
            "per_page": 500,
        })
        journals = data.get("journals", [])
        if not journals:
            break
        all_journals.extend(journals)
        page += 1
    return all_journals


# ──────────────────────────────────────────────
# データ集計
# ──────────────────────────────────────────────

def extract_monthly_amounts(trial_data: dict, account_names: list[str]) -> dict[str, int]:
    """
    試算表データから指定勘定科目の月別金額を抽出
    返り値: {"2024-03": 12345, "2024-04": 67890, ...}（単位：円）
    """
    monthly = {}
    items = trial_data.get("trial_pl_items", []) or trial_data.get("trial_bs_items", [])

    for item in items:
        if item.get("account_item_name") in account_names:
            for monthly_data in item.get("monthly", []):
                ym = monthly_data["year_month"]  # "2024-03" format
                amount = monthly_data.get("closing_balance", 0) or monthly_data.get("credit", 0) or 0
                monthly[ym] = monthly.get(ym, 0) + amount
    return monthly


def extract_bs_monthly_balance(trial_data: dict, account_names: list[str]) -> dict[str, int]:
    """BS残高の月次推移を抽出"""
    monthly = {}
    items = trial_data.get("trial_bs_items", [])

    for item in items:
        if item.get("account_item_name") in account_names:
            for monthly_data in item.get("monthly", []):
                ym = monthly_data["year_month"]
                balance = monthly_data.get("closing_balance", 0) or 0
                monthly[ym] = monthly.get(ym, 0) + balance
    return monthly


def classify_loan_journals(journals: list, config: dict) -> dict:
    """
    仕訳データから借入金の入金・返済を月別に分類
    返り値: {
        "loan_proceeds": {"2024-04": 20000000, ...},
        "loan_repayment_short": {...},
        "loan_repayment_long": {...}
    }
    """
    mapping = config["account_mapping"]
    result = {
        "loan_proceeds": {},
        "loan_repayment_short": {},
        "loan_repayment_long": {},
    }

    loan_accounts = set(
        mapping["loan_proceeds"]["accounts"]
        + mapping["loan_repayment_short"]["accounts"]
        + mapping["loan_repayment_long"]["accounts"]
    )

    for j in journals:
        journal_date = j.get("date", "")
        ym = journal_date[:7]  # "2024-03"
        for detail in j.get("details", []):
            acct = detail.get("account_item_name", "")
            if acct not in loan_accounts:
                continue
            debit = detail.get("debit_amount", 0) or 0
            credit = detail.get("credit_amount", 0) or 0

            # 借入金が貸方 → 借入実行（入金）
            if credit > 0 and acct in mapping["loan_proceeds"]["accounts"]:
                result["loan_proceeds"][ym] = result["loan_proceeds"].get(ym, 0) + credit

            # 借入金が借方 → 返済
            if debit > 0:
                if acct in mapping["loan_repayment_short"]["accounts"]:
                    result["loan_repayment_short"][ym] = result["loan_repayment_short"].get(ym, 0) + debit
                elif acct in mapping["loan_repayment_long"]["accounts"]:
                    result["loan_repayment_long"][ym] = result["loan_repayment_long"].get(ym, 0) + debit

    return result


# ──────────────────────────────────────────────
# 資金繰り表の生成
# ──────────────────────────────────────────────

def to_thousands(amount_yen: int) -> int:
    """円 → 千円（四捨五入）"""
    return round(amount_yen / 1000)


def build_cashflow(config: dict, period: int) -> list[list]:
    """MFデータから資金繰り表を構築"""
    mapping = config["account_mapping"]
    months = fiscal_months(period)
    month_keys = [f"{y}-{m:02d}" for y, m in months]
    month_labels = [f"{m}月" for _, m in months]

    print(f"第{period}期のデータを取得中...")

    # データ取得
    pl_data = fetch_monthly_pl(config, period)
    bs_data = fetch_monthly_bs(config, period)
    journals = fetch_journals(config, period)

    # PL項目（月次発生額）
    sales = extract_monthly_amounts(pl_data, mapping["sales"]["accounts"])
    personnel = extract_monthly_amounts(pl_data, mapping["personnel_expenses"]["accounts"])
    expenses = extract_monthly_amounts(pl_data, mapping["operating_expenses"]["accounts"])
    non_op_income = extract_monthly_amounts(pl_data, mapping["non_operating_income"]["accounts"])
    non_op_expense = extract_monthly_amounts(pl_data, mapping["non_operating_expenses"]["accounts"])

    # BS項目（月末残高）
    cash_balance = extract_bs_monthly_balance(bs_data, mapping["cash_and_deposits"]["accounts"])
    ar_balance = extract_bs_monthly_balance(bs_data, mapping["accounts_receivable"]["accounts"])
    ap_balance = extract_bs_monthly_balance(bs_data, mapping["accounts_payable"]["accounts"])
    inv_balance = extract_bs_monthly_balance(bs_data, mapping["inventory"]["accounts"])

    # 借入・返済（仕訳ベース）
    loan_data = classify_loan_journals(journals, config)

    # 前期のBS期末残高（前月繰越金の初期値用）
    prev_period_end = f"{FOUNDING_YEAR + period - 2 + 1}-02"
    prev_cash = cash_balance.get(prev_period_end, 0)

    # ---- 行データの構築 ----
    def get_val(data: dict, key: str) -> int:
        return data.get(key, 0)

    rows = {}

    # 今期売上
    rows["今期売上"] = [to_thousands(get_val(sales, k)) for k in month_keys]

    # 前月繰越金（現金・預金の月末残高から算出）
    # 3月の前月繰越金 = 前期末（2月末）の現金残高
    carry_forward = [0] * 12
    carry_forward[0] = to_thousands(prev_cash)
    # 以降は翌月繰越金から自動計算（後で設定）

    rows["前月繰越金（A）"] = carry_forward

    # 経常収入：売掛金の回収額 = 前月AR残高 + 当月売上 - 当月AR残高
    # 簡易的に当月の現金入金 = 売上（MF上の売掛金変動から算出）
    ar_collections = []
    cash_sales_row = []
    for i, k in enumerate(month_keys):
        prev_k = month_keys[i - 1] if i > 0 else prev_period_end
        prev_ar = get_val(ar_balance, prev_k)
        cur_ar = get_val(ar_balance, k)
        cur_sales = get_val(sales, k)
        # 回収額 = 前月AR + 当月売上 - 当月AR
        collection = prev_ar + cur_sales - cur_ar
        ar_collections.append(to_thousands(max(collection, 0)))
        cash_sales_row.append(0)  # Amazon物販は全て売掛金回収

    rows["収入_現金売上"] = cash_sales_row
    rows["収入_売掛金回収"] = ar_collections
    rows["収入_その他"] = [0] * 12

    # 経常収入計（B）
    rows["経常収入計（B）"] = [
        rows["収入_現金売上"][i] + rows["収入_売掛金回収"][i] + rows["収入_その他"][i]
        for i in range(12)
    ]

    # 経常支出
    # 買掛金支払 = 前月AP + 当月仕入 - 当月AP
    ap_payments = []
    for i, k in enumerate(month_keys):
        prev_k = month_keys[i - 1] if i > 0 else prev_period_end
        prev_ap = get_val(ap_balance, prev_k)
        cur_ap = get_val(ap_balance, k)
        # 仕入はPLの売上原価から推定（簡易）
        payment = prev_ap - cur_ap  # 支払分だけAPが減少
        if payment < 0:
            payment = 0  # 新規仕入でAP増加の場合は0
        ap_payments.append(to_thousands(payment))

    rows["支出_現金仕入"] = [0] * 12
    rows["支出_買掛金支払"] = ap_payments
    rows["支出_人件費"] = [to_thousands(get_val(personnel, k)) for k in month_keys]
    rows["支出_その他"] = [0] * 12

    # 商品棚卸高（在庫の増減）
    inventory_change = []
    for i, k in enumerate(month_keys):
        prev_k = month_keys[i - 1] if i > 0 else prev_period_end
        prev_inv = get_val(inv_balance, prev_k)
        cur_inv = get_val(inv_balance, k)
        change = cur_inv - prev_inv  # 在庫増=資金流出
        inventory_change.append(to_thousands(change))
    rows["支出_商品棚卸高"] = inventory_change

    rows["支出_諸経費"] = [to_thousands(get_val(expenses, k)) for k in month_keys]

    # 経常支出計（C）
    rows["経常支出計（C）"] = [
        rows["支出_現金仕入"][i]
        + rows["支出_買掛金支払"][i]
        + rows["支出_人件費"][i]
        + rows["支出_その他"][i]
        + rows["支出_商品棚卸高"][i]
        + rows["支出_諸経費"][i]
        for i in range(12)
    ]

    # 差引過不足（D）
    rows["差引過不足（D）=（B）-（C）"] = [
        rows["経常収入計（B）"][i] - rows["経常支出計（C）"][i]
        for i in range(12)
    ]

    # 経常外収支
    rows["経常外収入"] = [to_thousands(get_val(non_op_income, k)) for k in month_keys]
    rows["経常外支出"] = [to_thousands(get_val(non_op_expense, k)) for k in month_keys]
    rows["経常外収支計（E）"] = [
        rows["経常外収入"][i] - rows["経常外支出"][i]
        for i in range(12)
    ]

    # 財務収支
    rows["財務収入_借入"] = [
        to_thousands(get_val(loan_data["loan_proceeds"], k)) for k in month_keys
    ]
    rows["財務支出_借入金返済（短期）"] = [
        to_thousands(get_val(loan_data["loan_repayment_short"], k)) for k in month_keys
    ]
    rows["財務支出_借入金返済（長期）"] = [
        to_thousands(get_val(loan_data["loan_repayment_long"], k)) for k in month_keys
    ]
    rows["財務収支計（F）"] = [
        rows["財務収入_借入"][i]
        - rows["財務支出_借入金返済（短期）"][i]
        - rows["財務支出_借入金返済（長期）"][i]
        for i in range(12)
    ]

    # 翌月繰越金（G）= (A) + (D) + (E) + (F)
    next_carry = [0] * 12
    for i in range(12):
        next_carry[i] = (
            rows["前月繰越金（A）"][i]
            + rows["差引過不足（D）=（B）-（C）"][i]
            + rows["経常外収支計（E）"][i]
            + rows["財務収支計（F）"][i]
        )
        # 次月の前月繰越金を設定
        if i < 11:
            rows["前月繰越金（A）"][i + 1] = next_carry[i]

    rows["翌月繰越金（G）=（A）+（D）+（E）+（F）"] = next_carry

    # ---- CSV行の組み立て ----
    def make_row(label: str, values: list[int]) -> list:
        total = sum(values)
        avg = round(total / 12)
        return [label] + values + [total, avg]

    def make_row_no_total(label: str, values: list[int]) -> list:
        return [label] + values + ["", ""]

    header = ["単位：千円"] + [""] * 14
    title_row = [f"資金繰り表（第{period}期）"] + month_labels + ["合計", "月平均"]

    csv_rows = [
        header,
        title_row,
        make_row("今期売上", rows["今期売上"]),
        make_row_no_total("前月繰越金（A）", rows["前月繰越金（A）"]),
        make_row("収入_現金売上", rows["収入_現金売上"]),
        make_row("収入_売掛金回収", rows["収入_売掛金回収"]),
        make_row("収入_その他", rows["収入_その他"]),
        make_row("経常収入計（B）", rows["経常収入計（B）"]),
        make_row("支出_現金仕入", rows["支出_現金仕入"]),
        make_row("支出_買掛金支払", rows["支出_買掛金支払"]),
        make_row("支出_人件費", rows["支出_人件費"]),
        make_row("支出_その他", rows["支出_その他"]),
        make_row("支出_商品棚卸高", rows["支出_商品棚卸高"]),
        make_row("支出_諸経費", rows["支出_諸経費"]),
        make_row("経常支出計（C）", rows["経常支出計（C）"]),
        make_row("差引過不足（D）=（B）-（C）", rows["差引過不足（D）=（B）-（C）"]),
        make_row("経常外収入", rows["経常外収入"]),
        make_row("経常外支出", rows["経常外支出"]),
        make_row("経常外収支計（E）", rows["経常外収支計（E）"]),
        make_row("財務収入_借入", rows["財務収入_借入"]),
        make_row("財務支出_借入金返済（短期）", rows["財務支出_借入金返済（短期）"]),
        make_row("財務支出_借入金返済（長期）", rows["財務支出_借入金返済（長期）"]),
        make_row("財務収支計（F）", rows["財務収支計（F）"]),
        make_row_no_total("翌月繰越金（G）=（A）+（D）+（E）+（F）", rows["翌月繰越金（G）=（A）+（D）+（E）+（F）"]),
    ]

    return csv_rows


def write_csv(rows: list[list], period: int):
    """CSVファイルに書き出し"""
    start_year = FOUNDING_YEAR + period
    filename = OUTPUT_DIR / f"資金繰り表_{start_year}.csv"
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)
    print(f"出力完了: {filename}")


# ──────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="MoneyForward連携 資金繰り表生成")
    parser.add_argument("--period", type=int, required=True, help="期数（例: 8 = 第8期）")
    parser.add_argument("--auth", action="store_true", help="OAuth2認証フローを実行")
    args = parser.parse_args()

    config = load_config()

    if args.auth:
        auth_flow(config)
        return

    mf = config["moneyforward"]
    if not mf["access_token"]:
        print("エラー: 未認証です。--auth オプションで認証してください。")
        print(f"  python {sys.argv[0]} --period {args.period} --auth")
        sys.exit(1)

    if not mf["office_id"]:
        print("エラー: office_id が未設定です。--auth で再認証してください。")
        sys.exit(1)

    csv_rows = build_cashflow(config, args.period)
    write_csv(csv_rows, args.period)


if __name__ == "__main__":
    main()
