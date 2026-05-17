"""
NotebookLM向けCSVエクスポートモジュール
分析結果をCSVファイルとして自動出力し、NotebookLMへの投入を容易にする。
"""
import os
import pandas as pd
from utils import get_config, log


def export_csvs(df_merged, abc_df_raw, sell_through_df, yoy_df, kpi, cols, brand_output_dir=None):
    """NotebookLM向けCSVを サブフォルダ内の notebook_data/ に出力する"""
    cfg = get_config()
    base_dir = brand_output_dir or cfg["paths"]["output_dir"]
    output_dir = os.path.join(base_dir, "notebook_data")
    os.makedirs(output_dir, exist_ok=True)

    _export_monthly_summary(df_merged, yoy_df, cols, output_dir)
    _export_store_performance(df_merged, cols, output_dir)
    _export_abc_action_list(abc_df_raw, sell_through_df, df_merged, cols, output_dir)

    log(f"  NotebookLM用CSV出力先: {output_dir}")


def _export_monthly_summary(df_merged, yoy_df, cols, output_dir):
    """月別×カテゴリー別の売上サマリー（前年比付き）"""
    col_date = cols["date"]
    col_cat_m = cols["m_cat_m"]
    col_reg = cols["s_reg"]
    col_price = cols["s_price"]

    if not col_date or not df_merged[col_date].notna().any():
        return

    df = df_merged.copy()
    df["年月"] = df[col_date].dt.to_period("M").astype(str)

    agg_dict = {"売上数": (col_reg, "count")}
    if col_price:
        agg_dict["売上金額"] = (col_price, "sum")

    monthly = df.groupby(["年月", col_cat_m]).agg(**agg_dict).reset_index()
    monthly = monthly.rename(columns={col_cat_m: "中分類"})

    # 前年比データがあればマージ
    if yoy_df is not None and not yoy_df.empty:
        yoy_cols = ["年月", "中分類"]
        merge_cols = [c for c in ["前年販売数", "前年売上金額", "販売数前年比", "売上金額前年比"]
                      if c in yoy_df.columns]
        if merge_cols:
            monthly = pd.merge(monthly, yoy_df[yoy_cols + merge_cols],
                               on=yoy_cols, how="left")

    monthly = monthly.sort_values(["年月", "中分類"]).reset_index(drop=True)

    path = os.path.join(output_dir, "monthly_summary.csv")
    monthly.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"  出力: monthly_summary.csv ({len(monthly)}行)")


def _export_store_performance(df_merged, cols, output_dir):
    """店舗別KPI"""
    cfg = get_config()
    col_store = cols["s_store"]
    col_reg = cols["s_reg"]
    col_price = cols["s_price"]
    col_inbound = cols["s_inbound"]
    inbound_kw = "|".join(cfg["column_mapping"]["inbound_keywords"])

    agg_dict = {"売上数": (col_reg, "count")}
    if col_price:
        agg_dict["売上金額"] = (col_price, "sum")
    agg_dict["客数"] = (col_reg, "nunique")

    store = df_merged.groupby(col_store).agg(**agg_dict).reset_index()
    store = store.rename(columns={col_store: "店舗名"})

    if col_price:
        store["客単価"] = (store["売上金額"] / store["客数"].replace(0, 1)).round(0)

    if col_inbound:
        ib_rate = df_merged.groupby(col_store).apply(
            lambda g: g[col_inbound].astype(str).str.contains(
                inbound_kw, case=False, na=False
            ).mean()
        ).reset_index()
        ib_rate.columns = ["店舗名", "インバウンド比率"]
        store = pd.merge(store, ib_rate, on="店舗名", how="left")

    store = store.sort_values("売上金額" if "売上金額" in store.columns else "売上数",
                              ascending=False).reset_index(drop=True)

    path = os.path.join(output_dir, "store_performance.csv")
    store.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"  出力: store_performance.csv ({len(store)}行)")


def _export_abc_action_list(abc_df_raw, sell_through_df, df_merged, cols, output_dir):
    """ABCランク＋WOS＋消化率＋推奨アクション"""
    col_key = cols["s_key"]

    if abc_df_raw is None or abc_df_raw.empty:
        return

    df = abc_df_raw.copy()

    # WOSをマージ
    if "WOS(在庫週数)" in df_merged.columns:
        wos_map = (df_merged[[col_key, "WOS(在庫週数)"]]
                   .drop_duplicates(subset=[col_key])
                   .set_index(col_key)["WOS(在庫週数)"])
        df["WOS"] = df[col_key].map(wos_map)

    # 消化率をマージ
    if sell_through_df is not None and "消化率" in sell_through_df.columns:
        st_map = sell_through_df.set_index(col_key)["消化率"]
        df["消化率"] = df[col_key].map(st_map).fillna(0)

    # 推奨アクションの自動判定
    def recommend_action(row):
        rank = row.get("ABC×XYZ", "")
        wos = row.get("WOS", None)
        st = row.get("消化率", 0)

        if rank in ["AX", "AY"]:
            if wos is not None and wos < 2:
                return "⚠️ 欠品リスク！緊急補充"
            return "✅ 安定供給を維持"
        elif rank in ["AZ"]:
            return "📊 需要変動あり・機動的発注"
        elif rank in ["BX", "BY"]:
            return "📦 適正在庫維持"
        elif rank in ["BZ"]:
            if st < 0.3:
                return "⬇️ 消化率低・販促検討"
            return "📊 動向注視"
        elif rank in ["CX", "CY"]:
            return "📉 低在庫維持"
        elif rank in ["CZ"]:
            return "🔴 取扱終了を検討"
        return "-"

    df["推奨アクション"] = df.apply(recommend_action, axis=1)

    # 出力列の選定
    col_cat_m = cols["m_cat_m"]
    col_name = cols["m_name"]
    output_cols = [col_key, col_name, col_cat_m]
    for c in ["ABCランク", "XYZランク", "ABC×XYZ", "WOS", "消化率", "推奨アクション"]:
        if c in df.columns:
            output_cols.append(c)

    result = df[output_cols].rename(columns={
        col_key: "商品コード", col_name: "商品名", col_cat_m: "中分類"
    })

    result = result.sort_values(
        ["ABC×XYZ", "WOS"] if "WOS" in result.columns else ["ABC×XYZ"],
        ascending=True
    ).reset_index(drop=True)

    path = os.path.join(output_dir, "abc_action_list.csv")
    result.to_csv(path, index=False, encoding="utf-8-sig")
    log(f"  出力: abc_action_list.csv ({len(result)}行)")
