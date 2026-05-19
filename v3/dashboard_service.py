import json
import os
import pandas as pd
from datetime import datetime
from utils import get_config, log

def generate_dashboard_json(brand, df_merged, df_stock, abc_df, kpi_dict, opportunity_loss_df, cols):
    """
    ダッシュボード用JSONデータを生成し保存する
    md_dashboard: MD本部用データ
    stores: 各店舗用データ
    """
    cfg = get_config()
    output_dir = cfg["paths"]["output_dir"]
    json_path = os.path.join(output_dir, f"dashboard_data_{brand}.json")
    
    col_store = cols.get("s_store")
    col_key = cols.get("s_key")
    col_color = cols.get("m_color")
    col_stk = cols.get("stock_val")
    
    # 直近1週間（最新の売上日から遡って7日間）のフィルタリングと期間文字列生成
    col_date = cols.get("date")
    period_str = "期間データなし"
    df_latest_1w = df_merged
    
    if col_date and col_date in df_merged.columns:
        valid_dates = df_merged[col_date].dropna()
        if not valid_dates.empty:
            max_date = valid_dates.max()
            start_date = max_date - pd.Timedelta(days=6)
            # 最新の7日間に絞り込み
            df_latest_1w = df_merged[(df_merged[col_date] >= start_date) & (df_merged[col_date] <= max_date)]
            period_str = f"📅 {start_date.strftime('%Y/%m/%d')} 〜 {max_date.strftime('%Y/%m/%d')}"

    # 1. MD本部用データの構築
    md_dashboard = {
        "summary": {
            "total_sales": int(kpi_dict.get("total_sales", 0)),
            "total_revenue": int(kpi_dict.get("total_revenue", 0)),
            "member_ratio": round(float(kpi_dict.get("member_ratio", 0)) * 100, 1),
            "inbound_ratio": round(float(kpi_dict.get("inbound_ratio", 0)) * 100, 1)
        },
        "opportunity_loss": opportunity_loss_df.to_dict(orient="records") if opportunity_loss_df is not None else []
    }
    
    # 2. 店舗別データの構築
    stores_data = {}
    if col_store and col_store in df_merged.columns:
        stores = df_merged[col_store].dropna().unique()
        closed_stores = cfg.get("stores", {}).get("closed", [])
        
        for store in stores:
            store_str = str(store)
            if "POP-UP" in store_str or any(c in store_str for c in closed_stores):
                continue
                
            # 直近1週間の売上データから店舗別データを切り出し
            df_store = df_latest_1w[df_latest_1w[col_store] == store]
            if len(df_store) == 0:
                continue
                
            # 店舗ごとの簡易集計 (直近1週間の実績に基づく)
            total_sales = len(df_store)
            total_rev = int(df_store[cols["s_price"]].sum()) if cols["s_price"] else 0
            unique_tx = df_store[cols["s_reg"]].nunique() if cols["s_reg"] else 1
            
            # Kanken 23510 の正しい在庫ステータス計算
            kanken_status = []
            
            # 在庫CSV(df_stock)内の正しいカラム名を定義
            col_stock_store = df_stock.columns[1] if len(df_stock.columns) > 1 else col_store
            col_stock_key = df_stock.columns[2] if len(df_stock.columns) > 2 else col_key
            
            if col_stock_store and col_stock_key and col_color and col_stk \
               and col_stock_store in df_stock.columns \
               and col_stock_key in df_stock.columns \
               and col_color in df_stock.columns \
               and col_stk in df_stock.columns:
               
                mask = (
                    (df_stock[col_stock_store] == store) &
                    (df_stock[col_stock_key].astype(str).str.startswith("23510"))
                )
                kanken_stock_df = df_stock[mask].copy()
                
                if not kanken_stock_df.empty:
                    color_stock = (
                        kanken_stock_df
                        .groupby(col_color, as_index=False)[col_stk]
                        .sum()
                    )
                    for _, row in color_stock.iterrows():
                        kanken_status.append({
                            "color_name": str(row[col_color]),
                            "stock_num": int(row[col_stk])
                        })
            
            store_kpi = {
                "sales": total_rev,
                "customers": unique_tx,
                "avg_spend": round(total_rev / unique_tx) if unique_tx > 0 else 0,
                "items_per_customer": round(total_sales / unique_tx, 1) if unique_tx > 0 else 0,
                # モック用のダミー（前週比などの実装は今後の拡張）
                "wos_danger_count": 0,
                "wos_warning_count": 0,
                "kanken_status": kanken_status
            }
            stores_data[store] = {
                "store_name": store,
                "store_kpi_dashboard": store_kpi
            }

    # 全体をまとめる
    dashboard_data = {
        "generated_at": datetime.now().isoformat(),
        "brand": brand,
        "period": period_str,
        "md_dashboard": md_dashboard,
        "stores": stores_data
    }
    
    # 保存
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(dashboard_data, f, ensure_ascii=False, indent=2)
        log(f"  ダッシュボードデータ出力: {os.path.basename(json_path)}")
    except Exception as e:
        log(f"  [エラー] ダッシュボードデータの出力に失敗しました: {e}")
