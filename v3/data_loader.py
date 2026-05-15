"""
データ読み込み・前処理モジュール
CSVファイルの読み込み、結合、型変換を担当する。
"""
import pandas as pd
from utils import (
    get_config, find_latest_file, find_file_by_keyword,
    load_csv, validate_csv, log, log_step_start, log_step_end
)


def load_and_merge_data():
    """
    売上・在庫・マスタCSVを読み込み、結合して返す。
    カラムマッピングはconfig.jsonから取得する。
    """
    log_step_start("データ読み込み・前処理")
    cfg = get_config()
    ptn = cfg["csv_patterns"]
    cm  = cfg["column_mapping"]

    # --- CSV読み込み ---
    df_sales  = load_csv(find_latest_file(ptn["sales"]))
    df_stock  = load_csv(find_latest_file(ptn["stock"]))
    df_master = load_csv(find_latest_file(ptn["master"]))

    # カラム名クリーンアップ
    df_sales.columns  = [c.strip() for c in df_sales.columns]
    df_stock.columns  = [c.strip() for c in df_stock.columns]
    df_master.columns = [c.strip() for c in df_master.columns]

    log(f"  件数: 売上 {len(df_sales)} / 在庫 {len(df_stock)} / マスタ {len(df_master)}")

    # --- CSVバリデーション ---
    validate_csv(df_sales,  "sales",  "売上CSV")
    validate_csv(df_stock,  "stock",  "在庫CSV")
    validate_csv(df_master, "master", "マスタCSV")

    # --- カラムマッピング（config.jsonのインデックスから取得） ---
    sc = df_sales.columns.tolist()
    mc = df_master.columns.tolist()
    ic = df_stock.columns.tolist()

    si = cm["sales"]
    mi = cm["master"]

    col_s_reg     = sc[si["reg_idx"]]
    col_s_store   = sc[si["store_idx"]]
    col_s_key     = sc[si["key_idx"]]
    col_s_size    = sc[si["size_idx"]]    if len(sc) > si["size_idx"]    else None
    col_s_price   = sc[si["price_idx"]]   if len(sc) > si["price_idx"]   else None
    col_s_inbound = sc[si["inbound_idx"]] if len(sc) > si["inbound_idx"] else None
    col_s_age     = sc[si["age_idx"]]     if len(sc) > si["age_idx"]     else None
    col_s_member  = sc[si["member_idx"]]  if len(sc) > si["member_idx"]  else None

    col_m_key   = mc[mi["key_idx"]]
    col_m_name  = mc[mi["name_idx"]]
    col_m_brand = mc[mi["brand_idx"]]
    col_m_cost  = mc[mi["cost_idx"]]
    col_m_cat_l = mc[mi["cat_l_idx"]]
    col_m_cat_m = mc[mi["cat_m_idx"]]
    col_m_cat_s = mc[mi["cat_s_idx"]]
    col_m_color = mc[mi["color_idx"]]     if len(mc) > mi["color_idx"]   else None

    col_i_key = ic[cm["stock"]["key_idx"]]

    # 日付カラム検索
    col_date = None
    for c in sc:
        if any(x in c for x in cm["date_keywords"]):
            col_date = c
            break

    # 型変換
    df_sales[col_s_key]  = df_sales[col_s_key].astype(str)
    df_stock[col_i_key]  = df_stock[col_i_key].astype(str)
    df_master[col_m_key] = df_master[col_m_key].astype(str)

    # 価格・原価クリーンアップ
    for df, col in [(df_sales, col_s_price), (df_master, col_m_cost)]:
        if col:
            df[col] = (df[col].astype(str)
                       .str.replace("¥", "").str.replace(",", "").str.strip())
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 在庫数カラム特定
    stock_val_col = ic[-1]
    for c in ic:
        if any(x in c for x in cm["stock_val_keywords"]):
            stock_val_col = c
            break

    # 受入数量カラム（消化率用）
    stock_recv_col = None
    for c in ic:
        if any(x in c for x in cm["stock_recv_keywords"]):
            stock_recv_col = c
            break

    # 売上 + マスタ 結合
    df_merged = pd.merge(df_sales, df_master,
                         left_on=col_s_key, right_on=col_m_key,
                         how="left", suffixes=("", "_master"))

    # 在庫集計（品番レベル）
    df_stock_sum = df_stock.groupby(col_i_key)[stock_val_col].sum().reset_index()
    df_merged = pd.merge(df_merged, df_stock_sum,
                         left_on=col_s_key, right_on=col_i_key, how="left")
    df_merged[stock_val_col] = df_merged[stock_val_col].fillna(0)

    # 受入数量も品番レベルで集計
    recv_by_item = None
    if stock_recv_col:
        recv_by_item = df_stock.groupby(col_i_key)[stock_recv_col].sum().reset_index()

    # 日付変換
    if col_date:
        df_merged[col_date] = pd.to_datetime(df_merged[col_date], errors="coerce")

    # カラム辞書
    cols = {
        "s_reg": col_s_reg, "s_store": col_s_store, "s_key": col_s_key,
        "s_size": col_s_size, "s_price": col_s_price, "s_inbound": col_s_inbound,
        "s_age": col_s_age, "s_member": col_s_member, "date": col_date,
        "m_key": col_m_key, "m_name": col_m_name, "m_brand": col_m_brand,
        "m_cost": col_m_cost, "m_cat_l": col_m_cat_l, "m_cat_m": col_m_cat_m,
        "m_cat_s": col_m_cat_s, "m_color": col_m_color,
        "stock_val": stock_val_col, "stock_recv": stock_recv_col,
    }

    log_step_end("データ読み込み・前処理")
    return df_merged, df_stock, df_master, df_sales, recv_by_item, cols
