"""
FJALLRAVEN MD分析ツール V4.0
メインエントリーポイント

使い方:
  python main.py
  または「実行ボタン.bat」をダブルクリック
"""
import os
import sys
import pandas as pd
from datetime import datetime

# 自身のディレクトリをパスに追加（モジュール解決用）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils import (
    load_config, get_config, init_log, log, log_step_start, log_step_end,
    backup_db, find_file_by_keyword, load_csv, get_excel_path
)
from data_loader import load_and_merge_data
from analyzer import calc_wos, calc_abc_xyz, calc_sell_through, calc_kpi, calc_yoy, calc_opportunity_loss
from db_manager import save_to_db, save_prev_year_to_db
from csv_exporter import export_csvs
from dashboard_service import generate_dashboard_json


def main():
    try:
        # --- 初期化 ---
        cfg = load_config()
        init_log()
        log("=" * 50)
        log("FJALLRAVEN MD分析ツール V4.0 開始")
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
        df_merged_raw, df_stock_raw, df_master_raw, df_sales_raw, recv_by_item_raw, cols = load_and_merge_data()
        log(f"  総件数: 売上 {len(df_sales_raw)} / 在庫 {len(df_stock_raw)} / マスタ {len(df_master_raw)}")
        log_step_end("データ読み込み・前処理")

        # 対象ブランドの取得 (config.jsonの配列から取得。なければ単一文字列)
        target_brands = cfg["analysis"].get("target_brands", [cfg["analysis"].get("target_brand", "FRV")])
        log(f"分析対象ブランド一覧: {target_brands}")

        # 各ブランドについてループ処理
        for brand in target_brands:
            log("-" * 50)
            log(f"ブランド【 {brand} 】の処理を開始します")
            log("-" * 50)

            # ブランドごとの出力ディレクトリ作成
            brand_output_dir = os.path.join(cfg["paths"]["output_dir"], brand)
            os.makedirs(brand_output_dir, exist_ok=True)

            # 対象ブランドにデータを絞り込む
            col_brand = cols["m_brand"]
            df_merged = df_merged_raw[df_merged_raw[col_brand] == brand].copy()
            df_sales = df_sales_raw[df_sales_raw[col_brand] == brand].copy() if col_brand in df_sales_raw.columns else df_sales_raw.copy()
            df_stock = df_stock_raw[df_stock_raw[col_brand] == brand].copy() if col_brand in df_stock_raw.columns else df_stock_raw.copy()
            df_master = df_master_raw[df_master_raw[col_brand] == brand].copy() if col_brand in df_master_raw.columns else df_master_raw.copy()

            log(f"  絞り込み後件数 ({brand}): 売上 {len(df_sales)} / 在庫 {len(df_stock)} / マスタ {len(df_master)}")

            if len(df_merged) == 0:
                log(f"  警告: ブランド {brand} の売上データがありません。スキップします。")
                continue

            # --- Step 2: 前年データ読み込み ---
            log_step_start(f"前年データ読み込み ({brand})")
            prev_kw = cfg["csv_patterns"]["prev_year_keyword"]
            prev_file = find_file_by_keyword(prev_kw)
            df_prev = None
            if prev_file:
                df_prev = load_csv(prev_file)
                df_prev.columns = [c.strip() for c in df_prev.columns]
                log(f"  前年データ: {len(df_prev)}件")
            else:
                log("  前年データなし（前年比シートはスキップ）")
            log_step_end(f"前年データ読み込み ({brand})")

            # --- Step 3: 分析計算 ---
            log_step_start(f"分析計算 ({brand})")
            df_merged, max_date = calc_wos(df_merged, cols)
            abc_xyz_output, abc_df_raw = calc_abc_xyz(df_merged, cols, max_date)
            sell_through_df = calc_sell_through(df_merged, recv_by_item_raw, cols)
            kpi = calc_kpi(df_merged, cols)
            yoy_df = calc_yoy(df_merged, df_prev, df_master_raw, cols, brand) if df_prev is not None else None
            opp_loss_df = calc_opportunity_loss(df_merged, abc_df_raw, cols)
            log_step_end(f"分析計算 ({brand})")

            # --- Step 4: DB保存 ---
            log_step_start(f"DB保存 ({brand})")
            save_to_db(df_sales, df_stock, df_master, cols)
            if df_prev is not None:
                save_prev_year_to_db(df_prev)
            log_step_end(f"DB保存 ({brand})")

            # --- Step 5: Excel出力 ---
            log_step_start(f"Excel出力 ({brand})")
            output_path = os.path.join(
                brand_output_dir,
                f"MD分析レポートv3_{brand}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )
            try:
                from excel_writer import (
                    write_summary_sheet, write_category_sheet, write_trend_sheet,
                    write_abc_xyz_sheet, write_detail_sheet, write_store_comparison_sheet,
                    write_sell_through_sheet, write_kanken_sheet, write_inbound_sheet,
                    write_yoy_sheet, write_stock_health_sheet
                )
                from db_manager import get_stock_trend
                trend_df = get_stock_trend(brand, cols)

                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    write_summary_sheet(writer, kpi, cols, df_merged)
                    write_category_sheet(writer, df_merged, cols)
                    write_trend_sheet(writer, df_merged, cols)
                    write_stock_health_sheet(writer, trend_df, df_merged, cols)
                    write_abc_xyz_sheet(writer, abc_xyz_output)
                    write_detail_sheet(writer, df_merged, abc_df_raw, sell_through_df, cols)
                    write_store_comparison_sheet(writer, df_merged, df_stock, cols)
                    write_sell_through_sheet(writer, sell_through_df)
                    write_kanken_sheet(writer, df_merged, cols)
                    write_inbound_sheet(writer, df_merged, cols)
                    write_yoy_sheet(writer, yoy_df)
            except PermissionError:
                log(f"【エラー】Excelファイルが開いています: {output_path}。閉じてから再実行してください。")
                continue
            log_step_end(f"Excel出力 ({brand})")

            # --- Step 6: CSV出力 (NotebookLM用) ---
            log_step_start(f"CSV出力 ({brand})")
            export_csvs(df_merged, abc_df_raw, sell_through_df, yoy_df, kpi, cols, brand_output_dir)
            log_step_end(f"CSV出力 ({brand})")

            # --- Step 7: ダッシュボード用JSON出力 (V4) ---
            log_step_start(f"JSON出力 ({brand})")
            generate_dashboard_json(brand, df_merged, df_stock, abc_df_raw, kpi, opp_loss_df, cols)
            log_step_end(f"JSON出力 ({brand})")

            log(f"ブランド【 {brand} 】の処理完了: {os.path.basename(output_path)}")

        log("=" * 50)
        log("全ブランドの処理が完了しました")
        log("=" * 50)

    except Exception as e:
        import traceback
        log(f"致命的なエラー: {e}")
        log(traceback.format_exc())


if __name__ == "__main__":
    main()
