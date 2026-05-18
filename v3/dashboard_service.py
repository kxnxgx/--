import json
import os
import pandas as pd
from datetime import datetime
from utils import get_config, log

def generate_dashboard_json(brand, df_merged, abc_df, kpi_dict, opportunity_loss_df, cols):
    """
    ダッシュボード用JSONデータを生成し保存する
    md_dashboard: MD本部用データ
    stores: 各店舗用データ
    """
    cfg = get_config()
    output_dir = cfg["paths"]["output_dir"]
    json_path = os.path.join(output_dir, f"dashboard_data_{brand}.json")
    
    col_store = cols.get("s_store")
    
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
        for store in stores:
            df_store = df_merged[df_merged[col_store] == store]
            if len(df_store) == 0:
                continue
                
            # 店舗ごとの簡易集計 (実運用では店舗別KPI計算などを呼び出す)
            total_sales = len(df_store)
            total_rev = int(df_store[cols["s_price"]].sum()) if cols["s_price"] else 0
            unique_tx = df_store[cols["s_reg"]].nunique() if cols["s_reg"] else 1
            
            store_kpi = {
                "sales": total_rev,
                "customers": unique_tx,
                "avg_spend": round(total_rev / unique_tx) if unique_tx > 0 else 0,
                "items_per_customer": round(total_sales / unique_tx, 1) if unique_tx > 0 else 0,
                # モック用のダミー（前週比などの実装は今後の拡張）
                "wos_danger_count": 0,
                "wos_warning_count": 0,
                "kanken_status": [] 
            }
            stores_data[store] = {
                "store_name": store,
                "store_kpi_dashboard": store_kpi
            }
            
    # 全体をまとめる
    dashboard_data = {
        "generated_at": datetime.now().isoformat(),
        "brand": brand,
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
