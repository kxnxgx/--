"""
FJALLRAVEN MD分析ツール V3
メインエントリーポイント

使い方:
  python main.py
  または「実行ボタン.bat」をダブルクリック
"""
import os
import sys
import pandas as pd

# 自身のディレクトリをパスに追加（モジュール解決用）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils import (
    load_config, get_config, init_log, log, log_step_start, log_step_end,
    backup_db, find_file_by_keyword, load_csv, get_excel_path
)
from data_loader import load_and_merge_data
from analyzer import calc_wos, calc_abc_xyz, calc_sell_through, calc_kpi, calc_yoy
from excel_writer import (
    write_summary_sheet, write_category_sheet, write_trend_sheet,
    write_abc_xyz_sheet, write_detail_sheet, write_store_comparison_sheet,
    write_sell_through_sheet, write_kanken_sheet, write_inbound_sheet,
    write_yoy_sheet
)
from db_manager import save_to_db, save_prev_year_to_db
from csv_exporter import export_csvs


def main():
    try:
        # --- 初期化 ---
        cfg = load_config()
        os.makedirs(cfg["paths"]["output_dir"], exist_ok=True)
        init_log()
        log("=" * 50)
        log("FJALLRAVEN MD分析ツール V3 開始")
        log("=" * 50)

        try:
            import openpyxl
        except ImportError:
            log("エラー: openpyxlが見つかりません。pip install openpyxl を実行してください。")
            return

        # --- Step C: DBバックアップ ---
        log_step_start("DBバックアップ")
        backup_db()
        log_step_end("DBバックアップ")

        # --- Step 1: データ読み込み ---
        log_step_start("データ読み込み・前処理")
        df_merged, df_stock, df_master, df_sales, recv_by_item, cols = load_and_merge_data()
        log(f"  件数: 売上 {len(df_sales)} / 在庫 {len(df_stock)} / マスタ {len(df_master)}")
        
        # 対象ブランドに絞り込む（config.jsonで設定可能）
        target_brand = cfg["analysis"]["target_brand"]
        col_brand = cols["m_brand"]
        df_merged = df_merged[df_merged[col_brand] == target_brand].copy()
        log(f"  対象ブランド: {target_brand}（絞り込み後: {len(df_merged)}件）")

        # --- Step 2: 前年データ読み込み ---
        log_step_start("前年データ読み込み")
        prev_kw = cfg["csv_patterns"]["prev_year_keyword"]
        prev_file = find_file_by_keyword(prev_kw)
        df_prev = None
        if prev_file:
            df_prev = load_csv(prev_file)
            df_prev.columns = [c.strip() for c in df_prev.columns]
            log(f"  前年データ: {len(df_prev)}件")
        else:
            log("  前年データなし（前年比シートはスキップ）")
        log_step_end("前年データ読み込み")

        # --- Step 3: 分析計算 ---
        log_step_start("分析計算")
        df_merged, max_date = calc_wos(df_merged, cols)
        abc_xyz_output, abc_df_raw = calc_abc_xyz(df_merged, cols, max_date)
        sell_through_df = calc_sell_through(df_merged, recv_by_item, cols)
        kpi = calc_kpi(df_merged, cols)
        yoy_df = calc_yoy(df_merged, df_prev, df_master, cols) if df_prev is not None else None
        log_step_end("分析計算")

        # --- Step 4: DB保存 ---
        log_step_start("DB保存")
        save_to_db(df_sales, df_stock, df_master, cols)
        if df_prev is not None:
            save_prev_year_to_db(df_prev)
        log_step_end("DB保存")

        # --- Step 5: Excel出力 ---
        log_step_start("Excel出力")
        output_path = get_excel_path()
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                write_summary_sheet(writer, kpi, cols, df_merged)
                write_category_sheet(writer, df_merged, cols)
                write_trend_sheet(writer, df_merged, cols)
                write_abc_xyz_sheet(writer, abc_xyz_output)
                write_detail_sheet(writer, df_merged, abc_df_raw, sell_through_df, cols)
                write_store_comparison_sheet(writer, df_merged, df_stock, cols)
                write_sell_through_sheet(writer, sell_through_df)
                write_kanken_sheet(writer, df_merged, cols)
                write_inbound_sheet(writer, df_merged, cols)
                write_yoy_sheet(writer, yoy_df)
        except PermissionError:
            log("【エラー】Excelファイルが開いています。閉じてから再実行してください。")
            return
        log_step_end("Excel出力")

        # --- Step 6: CSV出力 (NotebookLM用) ---
        log_step_start("CSV出力")
        export_csvs(df_merged, abc_df_raw, sell_through_df, yoy_df, kpi, cols)
        log_step_end("CSV出力")

        log("=" * 50)
        log(f"完了: {os.path.basename(output_path)}")
        log("=" * 50)

    except Exception as e:
        import traceback
        log(f"致命的なエラー: {e}")
        log(traceback.format_exc())


if __name__ == "__main__":
    main()
