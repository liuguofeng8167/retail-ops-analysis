import argparse
import json
import os
import sys

import numpy as np
import pandas as pd

try:
    import xlsxwriter
except ImportError:
    raise SystemExit("缺少 xlsxwriter，请先在命令行运行：pip install xlsxwriter")

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

CANDIDATES = {
    "invoice": ["Invoice", "InvoiceNo", "invoice", "Order ID", "order_id"],
    "stock": ["StockCode", "stockcode", "Stock Code", "item_id"],
    "desc": ["Description", "description", "Product Name"],
    "qty": ["Quantity", "quantity", "Qty"],
    "date": ["InvoiceDate", "invoicedate", "Invoice Date", "Order Date", "date"],
    "price": ["Price", "price", "UnitPrice", "Unit Price"],
    "cust": ["Customer ID", "CustomerID", "Customer Id", "customer_id"],
    "country": ["Country", "country", "Region"],
}
REQUIRED = ["invoice", "qty", "date", "price"]


def find_columns(df):
    """按候选名找列，找不到就报错并列出当前列名，方便改。"""
    mapping = {}
    lower = {str(c).strip().lower(): c for c in df.columns}
    for key, names in CANDIDATES.items():
        for name in names:
            if name in df.columns:
                mapping[key] = name
                break
            if name.lower() in lower:
                mapping[key] = lower[name.lower()]
                break
    missing = [k for k in REQUIRED + ["cust"] if k not in mapping]
    if missing:
        raise SystemExit(
            f"缺少必要列：{missing}\n当前列名：{list(df.columns)}\n"
            f"请在 CANDIDATES 里补上对应列名"
        )
    return mapping


def load(path):
    if path.lower().endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(path, sheet_name=None)

        frames = []
        for name, df in sheets.items():
            if df.empty:
                continue
            df = df.copy()
            df["__sheet"] = name
            frames.append(df)
        return pd.concat(frames, ignore_index=True)
    return pd.read_csv(path, encoding="utf-8", low_memory=False)


def clean(raw):
    df = raw.copy()
    col = find_columns(df)
    df = df.rename(columns={v: k for k, v in col.items()})
    df = df.dropna(subset=["invoice", "date", "price"])

    df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
    df = df.dropna(subset=["date"])
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df = df.dropna(subset=["qty", "price"])


    df["amount"] = df["qty"] * df["price"]


    inv = df["invoice"].astype(str).str.strip()
    df["is_return"] = inv.str.upper().str.startswith("C") | (df["qty"] < 0)


    df = df[df["price"] > 0]

    df["customer_id"] = (
        df["cust"].astype(str).str.replace(r"\.0$", "", regex=True).replace("nan", np.nan)
        if "cust" in df.columns else np.nan
    )
    df["ym"] = df["date"].dt.to_period("M").astype(str)
    df["country"] = df["country"].fillna("Unknown") if "country" in df else "Unknown"
    return df


def monthly(df):
    sales = df[~df["is_return"]]
    rets = df[df["is_return"]]
    g = sales.groupby("ym").agg(
        gmv=("amount", "sum"),
        订单数=("invoice", "nunique"),
        客户数=("customer_id", "nunique"),
    )
    r = rets.groupby("ym")["amount"].sum().abs().rename("退货金额")
    out = g.join(r).fillna({"退货金额": 0})
    out["净销售额"] = out["gmv"] - out["退货金额"]
    out["客单价"] = (out["gmv"] / out["订单数"]).round(2)
    out["退货率_金额"] = (out["退货金额"] / out["gmv"]).round(4)
    out["GMV环比"] = out["gmv"].pct_change().round(4)
    return out.reset_index()


def by_country(df, top=10):
    sales = df[~df["is_return"]]
    out = sales.groupby("country").agg(
        净销售额=("amount", "sum"), 订单数=("invoice", "nunique")
    ).sort_values("净销售额", ascending=False).head(top)
    total = sales["amount"].sum()
    out["占比"] = (out["净销售额"] / total).round(4)
    out["客单价"] = (out["净销售额"] / out["订单数"]).round(2)
    return out.reset_index()


def rfm(df):
    """只对有客户 ID 的销售行做 RFM；散客(无 ID)会单独统计，不能混进复购率。"""
    sales = df[(~df["is_return"]) & df["customer_id"].notna()]
    if sales.empty:
        return pd.DataFrame(), {}
    snap = sales["date"].max() + pd.Timedelta(days=1)
    rfm_df = sales.groupby("customer_id").agg(
        R=("date", lambda s: (snap - s.max()).days),
        F=("invoice", "nunique"),
        M=("amount", "sum"),
    )
    rfm_df["R分"] = pd.qcut(rfm_df["R"], 5, labels=[5, 4, 3, 2, 1]).astype(int)
    rfm_df["F分"] = pd.qcut(rfm_df["F"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
    rfm_df["M分"] = pd.qcut(rfm_df["M"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)

    def label(row):
        r, f, m = row["R分"], row["F分"], row["M分"]
        hi = lambda x: x >= 4
        if hi(r) and hi(f) and hi(m):
            return "重要价值客户"
        if hi(r) and hi(f) and not hi(m):
            return "重要保持客户"
        if hi(r) and not hi(f) and hi(m):
            return "重要发展客户"
        if hi(r) and not hi(f) and not hi(m):
            return "新客户"
        if not hi(r) and hi(f) and hi(m):
            return "重要挽留客户"
        if not hi(r) and hi(f) and not hi(m):
            return "一般挽留客户"
        if not hi(r) and not hi(f) and hi(m):
            return "沉睡大客户"
        return "流失客户"

    rfm_df["分层"] = rfm_df.apply(label, axis=1)
    summary = rfm_df.groupby("分层").agg(
        客户数=("M", "size"), 总消费=("M", "sum"), 平均消费=("M", "mean"),
        平均订单数=("F", "mean"), 平均最近购买天数=("R", "mean"),
    ).round(2).sort_values("总消费", ascending=False)
    summary["客户占比"] = (summary["客户数"] / len(rfm_df)).round(4)
    summary["消费占比"] = (summary["总消费"] / rfm_df["M"].sum()).round(4)
    return summary.reset_index(), rfm_df


def retention(df, months=6):
    """首购月为 cohort，之后每月是否再购买。"""
    sales = df[df["is_return"] == False]
    sales = sales[sales["customer_id"].notna()]
    if sales.empty:
        return pd.DataFrame()
    first = sales.groupby("customer_id")["date"].min().dt.to_period("M")
    s = sales.copy()
    s["cohort"] = s["customer_id"].map(first)
    s["period"] = s["date"].dt.to_period("M")
    s["月序"] = (s["period"] - s["cohort"]).apply(lambda x: x.n)
    s = s[s["月序"] < months]
    base = s[s["月序"] == 0].groupby("cohort")["customer_id"].nunique()
    mat = s.groupby(["cohort", "月序"])["customer_id"].nunique().unstack(fill_value=0)
    mat = mat.div(base, axis=0)
    mat.index = mat.index.astype(str)
    mat.columns = [f"M{c}" for c in mat.columns]
    return mat.round(4)


def anomalies(m):
    """找“GMV 在涨、但退货率同时在涨”的月份——这是经营分析最该指出的信号。"""
    out = []
    for i in range(1, len(m)):
        cur, prev = m.iloc[i], m.iloc[i - 1]
        if cur["GMV环比"] > 0 and cur["退货率_金额"] > prev["退货率_金额"]:
            out.append({
                "月份": cur["ym"],
                "GMV环比": round(float(cur["GMV环比"]), 4),
                "退货率变化": round(float(cur["退货率_金额"] - prev["退货率_金额"]), 4),
            })
    return out


def write_excel(path, tables, charts=None):
    """输出一个可直接给业务看的 Excel，图用 xlsxwriter 原生图表。"""
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        for name, df in tables.items():
            df.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            for i, col in enumerate(df.columns):
                width = max(10, min(22, df[col].astype(str).str.len().max() + 4))
                sheet.set_column(i, i, width)

        if charts:
            for sheet_name, spec in charts.items():
                df = tables[sheet_name]
                sheet = writer.sheets[sheet_name]
                chart = book.add_chart({"type": spec["type"]})
                for series in spec["series"]:
                    n = len(df)
                    chart.add_series({
                        "name": series["name"],
                        "categories": [sheet_name, 1, 0, n, 0],
                        "values": [sheet_name, 1, series["col"], n, series["col"]],
                        "y2_axis": series.get("y2", False),
                    })
                chart.set_title({"name": spec["title"]})
                chart.set_x_axis({"name": spec.get("x", "")})
                if spec.get("y2"):
                    chart.set_y2_axis({"name": spec["y2"]})
                chart.set_size({"width": 780, "height": 380})
                sheet.insert_chart(spec.get("anchor", "A12"), chart)


def demo_data(months=12, seed=42):
    rng = np.random.default_rng(seed)
    countries = ["United Kingdom"] * 8 + ["Germany", "France", "EIRE", "Netherlands", "Spain"]
    rows = []
    customers = [f"C{10000 + i}" for i in range(600)] + [None] * 120
    codes = [f"S{i:04d}" for i in range(60)]
    base = pd.Timestamp("2010-12-01")
    for m in range(months):
        start = base + pd.DateOffset(months=m)
        n = rng.integers(700, 1400)
        for _ in range(n):
            order = f"{500000 + len(rows)}"
            day = start + pd.Timedelta(days=int(rng.integers(0, 28)))
            cust = rng.choice(customers)
            nlines = rng.integers(1, 4)
            for _ in range(nlines):
                qty = int(rng.integers(1, 12))
                price = round(float(rng.uniform(1.5, 40)), 2)
                rows.append([order, rng.choice(codes), "demo item", qty, day, price,
                             cust, rng.choice(countries)])
                if rng.random() < 0.06:
                    rows.append(["C" + order, rng.choice(codes), "demo item", -1, day, price,
                                 cust, countries[0]])
    return pd.DataFrame(rows, columns=["Invoice", "StockCode", "Description", "Quantity",
                                       "InvoiceDate", "Price", "Customer ID", "Country"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="交易明细 csv/xlsx 路径")
    ap.add_argument("--out", default="outputs", help="输出目录")
    ap.add_argument("--demo", action="store_true", help="用合成数据跑通流程")
    args = ap.parse_args()

    if args.demo:
        raw = demo_data()
        print("[demo] 使用合成数据，仅用于验证流程，结论不可写进简历")
    elif args.input:
        raw = load(args.input)
    else:
        raise SystemExit("请用 --input 指定数据文件，或用 --demo 先试跑")

    df = clean(raw)
    print(f"原始行数 {len(raw):,} → 清洗后 {len(df):,}；"
          f"退货行 {int(df['is_return'].sum()):,}")

    sales = df[~df["is_return"]]
    m = monthly(df)
    c = by_country(df)
    seg, rfm_raw = rfm(df)
    ret = retention(df)

    orders = sales["invoice"].nunique()
    custs = sales["customer_id"].nunique()
    repeat = 0.0
    if not rfm_raw.empty:
        repeat = (rfm_raw["F"] >= 2).mean()
    guest_share = sales["customer_id"].isna().mean()

    kpi = {
        "统计区间": f"{df['date'].min():%Y-%m-%d} ~ {df['date'].max():%Y-%m-%d}",
        "GMV": round(float(sales["amount"].sum()), 2),
        "退货金额": round(float(df.loc[df["is_return"], "amount"].sum()), 2),
        "净销售额": round(float(sales["amount"].sum() + df.loc[df["is_return"], "amount"].sum()), 2),
        "订单数": int(orders),
        "购买客户数": int(custs),
        "客单价": round(float(sales["amount"].sum() / orders), 2),
        "复购率": round(float(repeat), 4),
        "散客订单占比": round(float(guest_share), 4),
        "异常月份": anomalies(m),
    }

    os.makedirs(args.out, exist_ok=True)
    m.to_csv(os.path.join(args.out, "01_月度指标.csv"), index=False, encoding="utf-8-sig")
    c.to_csv(os.path.join(args.out, "02_国家净收入.csv"), index=False, encoding="utf-8-sig")
    if not seg.empty:
        seg.to_csv(os.path.join(args.out, "03_RFM分层.csv"), index=False, encoding="utf-8-sig")
    if not ret.empty:
        ret.to_csv(os.path.join(args.out, "04_客户留存矩阵.csv"), encoding="utf-8-sig")

    tables = {"月度指标": m, "国家净收入": c}
    charts = {
        "月度指标": {
            "type": "column", "title": "月度 GMV 与退货率", "x": "月份",
            "series": [{"name": "GMV", "col": 1},
                       {"name": "退货率", "col": m.columns.get_loc("退货率_金额"), "y2": True}],
            "y2": "退货率", "anchor": "A14",
        },
        "国家净收入": {
            "type": "bar", "title": "净收入 Top 国家", "x": "国家",
            "series": [{"name": "净销售额", "col": 1}], "anchor": "A14",
        },
    }
    if not seg.empty:
        tables["RFM分层"] = seg
        charts["RFM分层"] = {
            "type": "column", "title": "RFM 分层：客户数与消费占比", "x": "分层",
            "series": [{"name": "客户数", "col": 1},
                       {"name": "消费占比", "col": seg.columns.get_loc("消费占比"), "y2": True}],
            "y2": "消费占比", "anchor": "A12",
        }
    write_excel(os.path.join(args.out, "A_零售经营分析.xlsx"), tables, charts)

    with open(os.path.join(args.out, "关键指标.json"), "w", encoding="utf-8") as f:
        json.dump(kpi, f, ensure_ascii=False, indent=2)

    print(json.dumps(kpi, ensure_ascii=False, indent=2))
    print(f"\n结果已写入 {os.path.abspath(args.out)}")
    if args.demo:
        print("提示：以上数字来自合成数据，换成真实数据重跑后再用。")


if __name__ == "__main__":
    main()
