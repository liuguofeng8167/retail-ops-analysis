"""
B 项目｜用户行为漏斗与留存（阿里天池 UserBehavior 格式）
=========================================================
作用：把行为日志做成 漏斗 / 留存矩阵 / 用户分层 / 品类转化 四张表 + 带图 Excel。

用法：
    python funnel_analysis.py --input data/UserBehavior.csv --out outputs
    python funnel_analysis.py --input data/UserBehavior.csv --out outputs --sample 2000000
    python funnel_analysis.py --demo                 # 合成数据跑通流程

数据说明：天池 UserBehavior 无表头，列顺序为
    user_id, item_id, category_id, behavior_type(pv/cart/fav/buy), timestamp(秒)
"""
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

COLS = ["user_id", "item_id", "category_id", "behavior_type", "timestamp"]
BEHAVIORS = ["pv", "cart", "fav", "buy"]
CN = {"pv": "浏览", "cart": "加购", "fav": "收藏", "buy": "购买"}


def load(path, has_header=False, nrows=None, frac=None, seed=42):
    """完整读；文件太大时用 frac 按块随机抽样，保证覆盖整个时间区间。"""
    opts = dict(
        header=0 if has_header else None,
        names=None if has_header else COLS,
        dtype={"user_id": "int64", "item_id": "int64",
               "category_id": "int64", "timestamp": "int64"},
        low_memory=False,
    )
    if frac and frac < 1:

        parts = []
        for i, chunk in enumerate(pd.read_csv(path, chunksize=5_000_000, **opts)):
            parts.append(chunk.sample(frac=frac, random_state=seed + i))
            print(f"  已读 {5_000_000 * (i + 1):,} 行 …", end="\r")
        print(" " * 40, end="\r")
        df = pd.concat(parts, ignore_index=True)
    else:
        df = pd.read_csv(path, nrows=nrows, **opts)
    missing = [c for c in COLS if c not in df.columns]
    if missing:
        raise SystemExit(f"缺少列 {missing}；当前列 {list(df.columns)}")
    return df


def clean(df):
    df = df.dropna(subset=["user_id", "behavior_type", "timestamp"]).copy()
    df["behavior_type"] = df["behavior_type"].astype(str).str.strip().str.lower()
    df = df[df["behavior_type"].isin(BEHAVIORS)]

    df["ts"] = pd.to_datetime(df["timestamp"], unit="s", errors="coerce") + pd.Timedelta(hours=8)
    df = df.dropna(subset=["ts"])


    if len(df) > 1000:
        lo, hi = df["ts"].quantile([0.0005, 0.9995])
        bad = (df["ts"] < lo) | (df["ts"] > hi)
        if bad.any():
            print(f"  剔除时间戳异常 {int(bad.sum()):,} 行（保留 {lo:%Y-%m-%d} ~ {hi:%Y-%m-%d}）")
            df = df[~bad]

    df["day"] = df["ts"].dt.date
    df["ym"] = df["ts"].dt.to_period("M").astype(str)
    return df


def funnel(df):
    """两层口径：全量口径（有多少人做过该行为）+ 会话口径（同一天内是否按顺序走完）。"""
    total_users = df["user_id"].nunique()
    rows = []
    for b in BEHAVIORS:
        n = df.loc[df["behavior_type"] == b, "user_id"].nunique()
        rows.append({"行为": CN[b], "人数": n})
    full = pd.DataFrame(rows)
    base = full.loc[0, "人数"]
    full["转化率_占浏览"] = (full["人数"] / base).round(4)
    full["环比上一步"] = (full["人数"] / full["人数"].shift(1)).round(4)


    piv = (df.groupby(["user_id", "day", "behavior_type"])["ts"].min()
             .unstack("behavior_type"))
    for b in BEHAVIORS:
        if b not in piv.columns:
            piv[b] = pd.NaT
    sess = piv[piv["pv"].notna()]
    n_sess = len(sess)
    n_cart = int((sess["cart"].notna()).sum())
    n_cart_after = int(((sess["cart"].notna()) & (sess["cart"] > sess["pv"])).sum())
    n_buy = int(((sess["buy"].notna()) & (sess["cart"].notna()) &
                 (sess["cart"] > sess["pv"]) & (sess["buy"] > sess["cart"])).sum())
    session_view = pd.DataFrame([
        {"会话阶段": "有浏览的访问", "会话数": n_sess, "占浏览会话": 1.0},
        {"会话阶段": "浏览→加购", "会话数": n_cart_after,
         "占浏览会话": round(n_cart_after / n_sess, 4) if n_sess else 0},
        {"会话阶段": "浏览→加购→购买", "会话数": n_buy,
         "占浏览会话": round(n_buy / n_sess, 4) if n_sess else 0},
    ])
    return full, session_view


def retention(df, months=6, mode="buy", grain="M"):
    """mode=buy 看复购留存；mode=active 看回访留存。
    grain="M" 按月（跨度长的数据用），grain="D" 按天（只有几天的数据必须用这个，
    否则所有用户会落在同一个 cohort，矩阵只有一列，没有意义）。
    """
    if mode == "buy":
        d = df[df["behavior_type"] == "buy"]
        name = "复购留存" if grain == "M" else "复购留存(按天)"
    else:
        d = df
        name = "活跃留存" if grain == "M" else "活跃留存(按天)"
    if d.empty:
        return pd.DataFrame(), name
    d = d.copy()
    first = d.groupby("user_id")["ts"].min().dt.to_period(grain)
    d["cohort"] = d["user_id"].map(first)
    d["period"] = d["ts"].dt.to_period(grain)
    d["序"] = (d["period"] - d["cohort"]).apply(lambda x: x.n)
    d = d[d["序"] < months]
    base = d[d["序"] == 0].groupby("cohort")["user_id"].nunique()
    mat = d.groupby(["cohort", "序"])["user_id"].nunique().unstack(fill_value=0)
    mat = mat.div(base, axis=0).round(4)
    mat.index = mat.index.astype(str)
    prefix = "M" if grain == "M" else "D"
    mat.columns = [f"{prefix}{c}" for c in mat.columns]
    mat = mat.reset_index().rename(columns={"index": "cohort"})
    return mat, name


def segments(df):
    """没有金额字段，用购买次数代替 M，按 购买频次 × 最近活跃 分层。"""
    snap = df["ts"].max() + pd.Timedelta(days=1)
    g = df.groupby("user_id").agg(
        首次活跃=("ts", "min"), 最近活跃=("ts", "max"),
        行为次数=("ts", "size"),
    )
    buys = (df[df["behavior_type"] == "buy"].groupby("user_id")
            .agg(购买次数=("ts", "size"), 购买品类数=("category_id", "nunique")))
    g = g.join(buys).fillna({"购买次数": 0, "购买品类数": 0})
    g["R天"] = (snap - g["最近活跃"]).dt.days

    def label(r):
        f = r["购买次数"]
        if f >= 3:
            return "高频复购"
        if f == 2:
            return "中频复购"
        if f == 1:
            return "一次性购买"
        return "未购买"

    g["分层"] = g.apply(label, axis=1)
    out = g.groupby("分层").agg(
        用户数=("R天", "size"), 人均行为次数=("行为次数", "mean"),
        人均购买次数=("购买次数", "mean"), 平均最近活跃天数=("R天", "mean"),
    ).round(2)
    out["用户占比"] = (out["用户数"] / len(g)).round(4)
    return out.reset_index().sort_values("用户数", ascending=False)


def by_category(df, top=15):
    buy = df[df["behavior_type"] == "buy"]
    pv = df[df["behavior_type"] == "pv"]
    out = buy.groupby("category_id").agg(购买次数=("ts", "size"),
                                         购买人数=("user_id", "nunique"))
    view = pv.groupby("category_id")["user_id"].nunique().rename("浏览人数")
    out = out.join(view).fillna(0)
    out["浏览→购买转化"] = (out["购买人数"] / out["浏览人数"].replace(0, np.nan)).round(4)
    return out.sort_values("购买次数", ascending=False).head(top).reset_index()


def write_excel(path, tables, charts):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        for name, df in tables.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            sheet = writer.sheets[name[:31]]
            for i, col in enumerate(df.columns):
                width = max(9, min(20, df[col].astype(str).str.len().max() + 4))
                sheet.set_column(i, i, width)
        for sheet_name, spec in charts.items():
            df = tables[sheet_name]
            sheet = writer.sheets[sheet_name[:31]]
            chart = book.add_chart({"type": spec["type"]})
            n = len(df)
            for series in spec["series"]:
                chart.add_series({
                    "name": series["name"],
                    "categories": [sheet_name[:31], 1, 0, n, 0],
                    "values": [sheet_name[:31], 1, series["col"], n, series["col"]],
                })
            chart.set_title({"name": spec["title"]})
            chart.set_size({"width": 760, "height": 360})
            sheet.insert_chart(spec.get("anchor", "A10"), chart)


def demo_data(users=2500, days=90, seed=7):
    rng = np.random.default_rng(seed)
    start = pd.Timestamp("2017-11-01")
    rows = []
    cats = [f"C{i}" for i in range(30)]
    for u in range(users):
        first = start + pd.Timedelta(days=int(rng.integers(0, days - 5)))
        visits = int(rng.integers(1, 12))
        for v in range(visits):
            d = first + pd.Timedelta(days=int(rng.integers(0, 40)))
            ts0 = int(d.timestamp()) + int(rng.integers(0, 86000))
            n = int(rng.integers(2, 12))
            bought = rng.random() < 0.28
            for k in range(n):
                rows.append([1000 + u, 20000 + int(rng.integers(0, 500)), rng.choice(cats),
                             "pv", ts0 + k * 30])
            if rng.random() < 0.4:
                rows.append([1000 + u, 20000 + int(rng.integers(0, 500)), rng.choice(cats),
                             "cart", ts0 + n * 30 + 10])
            if rng.random() < 0.2:
                rows.append([1000 + u, 20000 + int(rng.integers(0, 500)), rng.choice(cats),
                             "fav", ts0 + n * 30 + 20])
            if bought:
                rows.append([1000 + u, 20000 + int(rng.integers(0, 500)), rng.choice(cats),
                             "buy", ts0 + n * 30 + 30])
    return pd.DataFrame(rows, columns=COLS)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="行为日志 csv 路径")
    ap.add_argument("--out", default="outputs")
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--sample", type=int, default=None, help="只读前 N 行，大数据集先用这个")
    ap.add_argument("--header", action="store_true", help="文件带表头时加上这个参数")
    ap.add_argument("--frac", type=float, default=None,
                    help="按比例随机抽样，如 0.08；大数据集用这个而不是 --sample")
    ap.add_argument("--grain", default="M", choices=["M", "D"],
                    help="留存粒度：M 按月，D 按天（只有几天的数据必须用 D）")
    args = ap.parse_args()

    if args.demo:
        raw = demo_data()
        print("[demo] 合成数据，仅验证流程，结论不可写进简历")
    elif args.input:
        raw = load(args.input, has_header=args.header, nrows=args.sample, frac=args.frac)
    else:
        raise SystemExit("请用 --input 指定数据，或用 --demo 试跑")

    df = clean(raw)
    print(f"行为记录 {len(df):,} 条，用户 {df['user_id'].nunique():,} 人，"
          f"区间 {df['ts'].min():%Y-%m-%d} ~ {df['ts'].max():%Y-%m-%d}")

    full, sess = funnel(df)
    ret_buy, name_buy = retention(df, mode="buy", grain=args.grain)
    ret_act, name_act = retention(df, mode="active", grain=args.grain)
    seg = segments(df)
    cat = by_category(df)

    buy_users = df.loc[df["behavior_type"] == "buy", "user_id"].nunique()

    steps = sess.iloc[1:].copy()
    steps["环节流失率"] = 1 - steps["占浏览会话"] / sess.iloc[0]["占浏览会话"]
    worst = steps.sort_values("占浏览会话").iloc[0]["会话阶段"] if len(steps) else ""
    kpi = {
        "行为记录数": int(len(df)),
        "用户数": int(df["user_id"].nunique()),
        "购买用户数": int(buy_users),
        "购买用户占比": round(buy_users / df["user_id"].nunique(), 4),
        "浏览→加购(会话口径)": float(sess.iloc[1]["占浏览会话"]),
        "浏览→加购→购买(会话口径)": float(sess.iloc[2]["占浏览会话"]),
        "流失最重的环节": worst,
        "留存粒度": "按月" if args.grain == "M" else "按天",
    }

    os.makedirs(args.out, exist_ok=True)
    full.to_csv(os.path.join(args.out, "01_漏斗_全量口径.csv"), index=False, encoding="utf-8-sig")
    sess.to_csv(os.path.join(args.out, "02_漏斗_会话口径.csv"), index=False, encoding="utf-8-sig")
    if not ret_buy.empty:
        ret_buy.to_csv(os.path.join(args.out, f"03_{name_buy}矩阵.csv"), index=False, encoding="utf-8-sig")
    if not ret_act.empty:
        ret_act.to_csv(os.path.join(args.out, f"04_{name_act}矩阵.csv"), index=False, encoding="utf-8-sig")
    seg.to_csv(os.path.join(args.out, "05_用户分层.csv"), index=False, encoding="utf-8-sig")
    cat.to_csv(os.path.join(args.out, "06_品类转化.csv"), index=False, encoding="utf-8-sig")

    tables = {"漏斗_会话口径": sess, "用户分层": seg, "品类转化": cat}
    charts = {
        "漏斗_会话口径": {
            "type": "column", "title": "会话漏斗：浏览→加购→购买",
            "series": [{"name": "会话数", "col": 1}], "anchor": "A9",
        },
        "用户分层": {
            "type": "column", "title": "用户分层：人数与占比",
            "series": [{"name": "用户数", "col": 1}], "anchor": "A9",
        },
        "品类转化": {
            "type": "bar", "title": "品类购买次数 Top",
            "series": [{"name": "购买次数", "col": 1}], "anchor": "A28",
        },
    }
    write_excel(os.path.join(args.out, "B_行为漏斗与留存.xlsx"), tables, charts)

    with open(os.path.join(args.out, "关键指标.json"), "w", encoding="utf-8") as f:
        json.dump(kpi, f, ensure_ascii=False, indent=2)

    print(json.dumps(kpi, ensure_ascii=False, indent=2))
    print(f"\n结果已写入 {os.path.abspath(args.out)}")


if __name__ == "__main__":
    main()
