"""
分析ロジックモジュール
WOS、ABC/XYZ、消化率、KPI、前年同月比の計算を担当する。
"""
import pandas as pd
from datetime import timedelta
from utils import get_config, get_season, log


def calc_wos(df_merged, cols):
    """WOS（在庫週数）を直近N日の平均売上で計算"""
    cfg = get_config()
    period = cfg["analysis"]["wos_period_days"]
    col_date = cols["date"]
    col_key  = cols["s_key"]
    col_reg  = cols["s_reg"]
    col_stk  = cols["stock_val"]

    if col_date and df_merged[col_date].notna().any():
        max_date = df_merged[col_date].max()
        cutoff   = max_date - timedelta(days=period)
        df_recent = df_merged[df_merged[col_date] >= cutoff]
        sales_recent = df_recent.groupby(col_key)[col_reg].count().reset_index(name="sales_recent")
        weeks = period / 7.0
        sales_recent["weekly_avg"] = sales_recent["sales_recent"] / weeks
        df_merged = pd.merge(df_merged, sales_recent, on=col_key, how="left")
        df_merged["weekly_avg"]   = df_merged["weekly_avg"].fillna(0.01)
        df_merged["WOS(在庫週数)"] = (df_merged[col_stk] / df_merged["weekly_avg"]).round(1)
        return df_merged, max_date

    df_merged["weekly_avg"]   = 0.01
    df_merged["WOS(在庫週数)"] = "N/A"
    return df_merged, None


def calc_abc_xyz(df_merged, cols, max_date):
    """ABC/XYZ分析（中分類=商品種別2を含む）"""
    cfg = get_config()
    a_th = cfg["analysis"]["abc_a_threshold"]
    b_th = cfg["analysis"]["abc_b_threshold"]
    x_th = cfg["analysis"]["xyz_x_threshold"]
    y_th = cfg["analysis"]["xyz_y_threshold"]

    col_key   = cols["s_key"]
    col_name  = cols["m_name"]
    col_cat_l = cols["m_cat_l"]
    col_cat_m = cols["m_cat_m"]
    col_price = cols["s_price"]
    col_cost  = cols["m_cost"]
    col_reg   = cols["s_reg"]
    col_stk   = cols["stock_val"]
    col_date  = cols["date"]

    agg_dict = {
        "販売数":   (col_reg, "count"),
        "売上金額":  (col_price, "sum") if col_price else (col_reg, lambda x: 0),
        "原価合計":  (col_cost, "sum")  if col_cost  else (col_reg, lambda x: 0),
        "平均在庫数": (col_stk, "mean"),
    }
    abc_df = df_merged.groupby(
        [col_key, col_name, col_cat_l, col_cat_m]
    ).agg(**agg_dict).reset_index()

    abc_df["売価"]   = (abc_df["売上金額"] / abc_df["販売数"]).fillna(0).round(0)
    if col_cost:
        abc_df["粗利額"] = abc_df["売上金額"] - abc_df["原価合計"]
        abc_df["原価率"] = (abc_df["原価合計"] / abc_df["売上金額"].replace(0, 1)).fillna(0)
    else:
        abc_df["粗利額"] = 0
        abc_df["原価率"] = 0
    abc_df["回転率"] = (abc_df["販売数"] / abc_df["平均在庫数"].replace(0, 0.01)).round(2)

    abc_df = abc_df.sort_values("売上金額", ascending=False)
    total  = abc_df["売上金額"].sum()
    abc_df["売上構成"]  = (abc_df["売上金額"] / total).fillna(0)
    abc_df["累計売上構成"] = abc_df["売上構成"].cumsum()
    abc_df["ABCランク"] = abc_df["累計売上構成"].apply(
        lambda r: "A" if r <= a_th else ("B" if r <= b_th else "C")
    )

    # XYZ
    if col_date and max_date is not None:
        pivot = df_merged.pivot_table(
            index=col_key,
            columns=pd.Grouper(key=col_date, freq="D"),
            values=col_reg, aggfunc="count", fill_value=0
        )
        xyz = pd.DataFrame()
        xyz["mean"] = pivot.mean(axis=1)
        xyz["std"]  = pivot.std(axis=1)
        xyz["需要変動係数"] = (xyz["std"] / xyz["mean"].replace(0, 0.01)).fillna(0).round(2)
        xyz["XYZランク"]  = xyz["需要変動係数"].apply(
            lambda cv: "X" if cv <= x_th else ("Y" if cv <= y_th else "Z")
        )
        abc_df = pd.merge(abc_df, xyz[["需要変動係数", "XYZランク"]], on=col_key, how="left")
    else:
        abc_df["需要変動係数"] = 0
        abc_df["XYZランク"] = "-"

    abc_df["ABC×XYZ"] = abc_df["ABCランク"] + abc_df["XYZランク"]

    output = abc_df[[
        col_key, col_name, col_cat_l, col_cat_m,
        "売価", "販売数", "売上金額", "原価率", "粗利額",
        "平均在庫数", "回転率", "売上構成", "累計売上構成",
        "ABCランク", "需要変動係数", "XYZランク", "ABC×XYZ"
    ]].rename(columns={
        col_key: "商品コード", col_name: "商品名",
        col_cat_l: "大分類", col_cat_m: "中分類"
    })
    return output, abc_df


def calc_sell_through(df_merged, recv_by_item, cols):
    """消化率（販売数 / 受入数量）を品番レベルで計算"""
    col_key   = cols["s_key"]
    col_reg   = cols["s_reg"]
    col_cat_m = cols["m_cat_m"]
    col_stk   = cols["stock_val"]
    col_date  = cols["date"]

    base = df_merged.groupby(col_key).agg(
        販売数=(col_reg, "count"),
        売上金額=(cols["s_price"], "sum") if cols["s_price"] else (col_reg, lambda x: 0),
        現在庫=(col_stk, "mean"),
    ).reset_index()

    cat_map = df_merged[[col_key, col_cat_m]].drop_duplicates(subset=[col_key])
    base = pd.merge(base, cat_map, on=col_key, how="left")

    if col_date:
        month_mode = (df_merged.dropna(subset=[col_date])
                      .groupby(col_key)[col_date]
                      .apply(lambda s: s.dt.month.mode().iloc[0]
                             if len(s) > 0 and len(s.dt.month.mode()) > 0 else 0))
        base["シーズン"] = month_mode.map(lambda m: get_season(int(m)) if m else "-")
    else:
        base["シーズン"] = "-"

    if recv_by_item is not None:
        recv_col = recv_by_item.columns[-1]
        base = pd.merge(base, recv_by_item.rename(columns={recv_by_item.columns[0]: col_key}),
                        on=col_key, how="left")
        base[recv_col] = base[recv_col].fillna(0)
        base["消化率"] = (base["販売数"] / base[recv_col].replace(0, 1)).clip(0, 1).round(3)
        base = base.rename(columns={recv_col: "受入数量"})
    else:
        base["受入数量"] = 0
        base["消化率"]  = 0

    return base


def calc_kpi(df_merged, cols):
    """総合KPIを計算"""
    cfg = get_config()
    col_reg     = cols["s_reg"]
    col_price   = cols["s_price"]
    col_inbound = cols["s_inbound"]
    col_member  = cols["s_member"]
    inbound_kw  = "|".join(cfg["column_mapping"]["inbound_keywords"])

    total_sales   = len(df_merged)
    total_revenue = df_merged[col_price].sum() if col_price else 0
    unique_tx     = df_merged[col_reg].nunique()
    avg_spend     = total_revenue / unique_tx if unique_tx > 0 else 0
    items_per_tx  = total_sales   / unique_tx if unique_tx > 0 else 0
    member_ratio  = df_merged[col_member].notna().mean() if col_member else 0
    inbound_ratio = (df_merged[col_inbound].astype(str)
                     .str.contains(inbound_kw, case=False, na=False)
                     .mean() if col_inbound else 0)
    return {
        "total_sales": total_sales, "total_revenue": total_revenue,
        "unique_tx": unique_tx, "avg_spend": avg_spend,
        "items_per_tx": items_per_tx, "member_ratio": member_ratio,
        "inbound_ratio": inbound_ratio,
    }


def calc_yoy(df_current, df_prev, df_master, cols, brand=None):
    """前年同月比を月別×中分類で計算"""
    col_date  = cols["date"]
    col_cat_m = cols["m_cat_m"]
    col_reg   = cols["s_reg"]
    col_price = cols["s_price"]
    col_m_key = cols["m_key"]

    if not col_date:
        return None

    # --- 今年 ---
    df_cur = df_current.copy()
    df_cur["年月"] = df_cur[col_date].dt.to_period("M").astype(str)
    cur_agg = df_cur.groupby(["年月", col_cat_m]).agg(
        今年販売数=(col_reg, "count"),
        今年売上金額=(col_price, "sum") if col_price else (col_reg, lambda x: 0),
    ).reset_index().rename(columns={col_cat_m: "中分類"})

    # --- 前年 ---
    cfg = get_config()
    si = cfg["column_mapping"]["sales"]
    prev_cols = df_prev.columns.tolist()
    
    prev_reg_col   = prev_cols[si["reg_idx"]]
    prev_key_col   = prev_cols[si["key_idx"]]
    prev_price_col = prev_cols[si["price_idx"]] if len(prev_cols) > si["price_idx"] else None
    prev_date_col  = next((c for c in prev_cols if any(x in c for x in ["日","日付","年月日","Date"])), None)

    if not prev_date_col or not prev_key_col:
        log("  警告: 前年データに日付列または品番列が見つかりません。前年比スキップ。")
        return None

    df_py = df_prev.copy()
    df_py[prev_date_col] = pd.to_datetime(df_py[prev_date_col], errors="coerce")
    if prev_price_col:
        df_py[prev_price_col] = (df_py[prev_price_col].astype(str)
                                 .str.replace("¥","").str.replace(",","").str.strip())
        df_py[prev_price_col] = pd.to_numeric(df_py[prev_price_col], errors="coerce").fillna(0)

    # 年月を「今年に投影した月」にする
    df_py["年月"] = df_py[prev_date_col].apply(
        lambda d: f"{d.year+1}-{d.month:02d}" if pd.notna(d) else None
    )

    # 前年の品番とマスタを結合して中分類を取得
    df_py[prev_key_col] = df_py[prev_key_col].astype(str)
    
    col_brand = cols["m_brand"]
    target_brand = brand or cfg["analysis"]["target_brand"]
    df_master_map = df_master[[col_m_key, col_cat_m, col_brand]].copy()
    df_master_map[col_m_key] = df_master_map[col_m_key].astype(str)

    df_py = pd.merge(df_py, df_master_map, left_on=prev_key_col, right_on=col_m_key, how="left")
    df_py = df_py[df_py[col_brand] == target_brand]
    df_py[col_cat_m] = df_py[col_cat_m].fillna("その他・不明")

    prev_agg = df_py.groupby(["年月", col_cat_m]).agg(
        前年販売数=(prev_reg_col, "count"),
        前年売上金額=(prev_price_col, "sum") if prev_price_col else (prev_reg_col, lambda x: 0),
    ).reset_index().rename(columns={col_cat_m: "中分類"})

    yoy = pd.merge(cur_agg, prev_agg, on=["年月", "中分類"], how="outer").fillna(0)
    yoy = yoy.sort_values(["年月", "中分類"]).reset_index(drop=True)
    yoy["販売数前年比"]   = yoy.apply(
        lambda r: r["今年販売数"] / r["前年販売数"] if r["前年販売数"] > 0 else None, axis=1)
    yoy["売上金額前年比"] = yoy.apply(
        lambda r: r["今年売上金額"] / r["前年売上金額"] if r["前年売上金額"] > 0 else None, axis=1)
    return yoy
