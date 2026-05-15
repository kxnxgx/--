"""
データベース管理モジュール
DB保存（正規化）・前年データ保存・SQL集計クエリを担当する。
"""
import sqlite3
import pandas as pd
from datetime import datetime
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


