"""
データベース管理モジュール
DB保存（正規化）・前年データ保存・SQL集計クエリを担当する。
"""
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
from utils import get_config, get_db_path, log


def save_to_db(df_sales, df_stock, df_master, cols):
    """
    正規化DB設計:
      sales_raw    : 売上トランザクション（マスタ未結合）
      stock_history: 在庫スナップショット（重複防止付き）
      master_data  : 商品マスタ（最新を上書き）
    """
    log("  DB保存開始...")
    conn = sqlite3.connect(get_db_path())

    # master_data: 最新を上書き
    df_master.to_sql("master_data", conn, if_exists="replace", index=False)

    # sales_raw: 日付範囲で重複防止
    col_date = cols["date"]
    if col_date and df_sales[col_date].notna().any():
        df_save = df_sales.copy()
        df_save[col_date] = df_save[col_date].astype(str)
        min_d = df_save[col_date].min()
        max_d = df_save[col_date].max()
        try:
            conn.execute(f'DELETE FROM sales_raw WHERE "{col_date}" BETWEEN ? AND ?', (min_d, max_d))
            conn.commit()
        except sqlite3.OperationalError:
            pass
        df_save.to_sql("sales_raw", conn, if_exists="append", index=False)
    else:
        df_sales.to_sql("sales_raw", conn, if_exists="append", index=False)

    # stock_history: 同日データ重複防止
    today_str = datetime.now().strftime("%Y-%m-%d")
    try:
        conn.execute("DELETE FROM stock_history WHERE 実行日 = ?", (today_str,))
        conn.commit()
    except sqlite3.OperationalError:
        pass
    df_stock_save = df_stock.copy()
    df_stock_save["実行日"] = today_str
    df_stock_save.to_sql("stock_history", conn, if_exists="append", index=False)

    conn.close()
    log("  DB保存完了。")


def save_prev_year_to_db(df_prev):
    """前年売上データをsales_raw_prevとしてDB保存"""
    conn = sqlite3.connect(get_db_path())
    df_save = df_prev.copy()
    date_cols = [c for c in df_prev.columns if any(x in c for x in ["日","日付","年月日","Date"])]
    if date_cols:
        df_save[date_cols[0]] = df_save[date_cols[0]].astype(str)
    df_save.to_sql("sales_raw_prev", conn, if_exists="replace", index=False)
    conn.close()
    log(f"  前年データDB保存完了（{len(df_prev)}件）。")


def get_stock_trend(brand, cols, weeks=8):
    """
    過去N週間分の週次在庫スナップショット履歴（ブランド・中分類別）をDBからロードし、集計して返す。
    """
    conn = sqlite3.connect(get_db_path())
    
    # 1. テーブルの存在チェック
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='stock_history'")
    if not cursor.fetchone():
        conn.close()
        log("  警告: stock_historyテーブルがDBに存在しません。空の在庫トレンドを返します。")
        return pd.DataFrame()

    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='master_data'")
    if not cursor.fetchone():
        conn.close()
        log("  警告: master_dataテーブルがDBに存在しません。空の在庫トレンドを返します。")
        return pd.DataFrame()

    # 2. 過去N週間（N*7日）のデータを取得
    cutoff_date = (datetime.now() - timedelta(days=weeks * 7)).strftime("%Y-%m-%d")
    query = "SELECT * FROM stock_history WHERE 実行日 >= ?"
    df_stock_hist = pd.read_sql_query(query, conn, params=(cutoff_date,))
    df_master = pd.read_sql_query("SELECT * FROM master_data", conn)
    conn.close()
    
    if df_stock_hist.empty or df_master.empty:
        log("  情報: 在庫履歴またはマスタが空です。")
        return pd.DataFrame()

    # 3. カラム名のstrip処理
    df_stock_hist.columns = [c.strip() for c in df_stock_hist.columns]
    df_master.columns = [c.strip() for c in df_master.columns]

    # 4. 在庫キーと在庫数カラムの特定
    cfg = get_config()
    stock_key_idx = cfg["column_mapping"]["stock"]["key_idx"]
    stock_cols = df_stock_hist.columns.tolist()
    if len(stock_cols) <= stock_key_idx:
        return pd.DataFrame()
    col_i_key = stock_cols[stock_key_idx]

    stock_val_col = None
    for c in stock_cols:
        if any(x in c for x in cfg["column_mapping"]["stock_val_keywords"]):
            stock_val_col = c
            break
    if not stock_val_col:
        stock_val_col = stock_cols[-1]

    # 5. 結合と絞り込み
    df_stock_hist[col_i_key] = df_stock_hist[col_i_key].astype(str)
    df_master[cols["m_key"]] = df_master[cols["m_key"]].astype(str)
    
    df_merged = pd.merge(df_stock_hist, df_master, left_on=col_i_key, right_on=cols["m_key"], how="inner", suffixes=("", "_master"))
    
    col_brand = cols["m_brand"]
    if col_brand not in df_merged.columns:
        if f"{col_brand}_master" in df_merged.columns:
            col_brand = f"{col_brand}_master"
        elif f"{col_brand}_x" in df_merged.columns:
            col_brand = f"{col_brand}_x"
        elif f"{col_brand}_y" in df_merged.columns:
            col_brand = f"{col_brand}_y"

    df_merged = df_merged[df_merged[col_brand] == brand].copy()
    
    if df_merged.empty:
        return pd.DataFrame()

    # 6. 金額計算
    col_cost = cols["m_cost"]
    if col_cost not in df_merged.columns:
        if f"{col_cost}_master" in df_merged.columns:
            col_cost = f"{col_cost}_master"
        elif f"{col_cost}_x" in df_merged.columns:
            col_cost = f"{col_cost}_x"
        elif f"{col_cost}_y" in df_merged.columns:
            col_cost = f"{col_cost}_y"

    df_merged[col_cost] = pd.to_numeric(df_merged[col_cost].astype(str).str.replace("¥", "").str.replace(",", "").str.strip(), errors="coerce").fillna(0)
    df_merged[stock_val_col] = pd.to_numeric(df_merged[stock_val_col], errors="coerce").fillna(0)
    df_merged["在庫金額(原価)"] = df_merged[stock_val_col] * df_merged[col_cost]
    
    # 7. 集計
    col_cat_m = cols["m_cat_m"]
    if col_cat_m not in df_merged.columns:
        if f"{col_cat_m}_master" in df_merged.columns:
            col_cat_m = f"{col_cat_m}_master"
        elif f"{col_cat_m}_x" in df_merged.columns:
            col_cat_m = f"{col_cat_m}_x"
        elif f"{col_cat_m}_y" in df_merged.columns:
            col_cat_m = f"{col_cat_m}_y"

    trend = df_merged.groupby(["実行日", col_cat_m]).agg(
        在庫数量合計=(stock_val_col, "sum"),
        在庫金額合計=("在庫金額(原価)", "sum")
    ).reset_index()
    
    trend = trend.rename(columns={col_cat_m: "中分類"})
    log(f"  在庫トレンド取得完了（{len(trend)}行の履歴データ）。")
    return trend


