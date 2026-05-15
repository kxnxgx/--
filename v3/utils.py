"""
ユーティリティモジュール
- 設定ファイル読み込み
- ログ出力（処理時間計測対応）
- ファイル検索・CSV読み込み
- Excel列幅調整
- DBバックアップ
"""
import os
import sys
import json
import glob
import shutil
import time
import pandas as pd
from datetime import datetime, timedelta

# =====================================
# 設定管理
# =====================================
_config = None
_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def load_config(path=None):
    """config.jsonを読み込んでグローバルに保持する"""
    global _config
    p = path or _config_path
    with open(p, "r", encoding="utf-8") as f:
        _config = json.load(f)
    return _config

def get_config():
    """現在の設定を返す（未読み込みなら自動読み込み）"""
    if _config is None:
        load_config()
    return _config

def get_output_dir():
    return get_config()["paths"]["output_dir"]

def get_base_dir():
    return get_config()["paths"]["base_dir"]

def get_db_path():
    cfg = get_config()
    return os.path.join(cfg["paths"]["output_dir"], cfg["paths"]["db_filename"])

def get_excel_path():
    return os.path.join(
        get_output_dir(),
        f"MD分析レポートv3_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )

def get_log_path():
    return os.path.join(get_output_dir(), "process_log.txt")


# =====================================
# ログ出力（処理時間計測対応）
# =====================================
_step_start = None

def log(msg):
    """タイムスタンプ付きログ出力"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    with open(get_log_path(), "a", encoding="utf-8") as f:
        f.write(line + "\n")
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode())

def log_step_start(step_name):
    """処理ステップの開始を記録"""
    global _step_start
    _step_start = time.time()
    log(f"[START] {step_name}")

def log_step_end(step_name):
    """処理ステップの終了と所要時間を記録"""
    global _step_start
    elapsed = time.time() - _step_start if _step_start else 0
    log(f"[DONE]  {step_name} ({elapsed:.1f}秒)")
    _step_start = None

def init_log():
    """ログファイルを初期化"""
    p = get_log_path()
    if os.path.exists(p):
        os.remove(p)


# =====================================
# ファイル検索
# =====================================
def find_latest_file(pattern):
    """パターンに一致する最新ファイルを返す"""
    files = glob.glob(os.path.join(get_base_dir(), pattern))
    if not files:
        raise FileNotFoundError(f"ファイルが見つかりません: {pattern}")
    latest = max(files, key=os.path.getmtime)
    log(f"  自動選択: {os.path.basename(latest)}")
    return latest

def find_file_by_keyword(keyword):
    """キーワードを含むCSVファイルを検索（前年データ用）"""
    files = glob.glob(os.path.join(get_base_dir(), f"*{keyword}*.csv"))
    if not files:
        return None
    result = max(files, key=os.path.getmtime)
    log(f"  前年データ検出: {os.path.basename(result)}")
    return result


# =====================================
# CSV読み込み＋バリデーション
# =====================================
def load_csv(path):
    """CSVを自動エンコーディング判定で読み込む"""
    log(f"  読み込み中: {os.path.basename(path)}")
    for enc in ["cp932", "utf-8-sig", "utf-8"]:
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False)
        except Exception:
            continue
    raise ValueError(f"読み込み失敗: {path}")

def validate_csv(df, csv_type, filename):
    """
    CSVのカラム数が設定値以上あるかチェック。
    csv_type: "sales" / "master" / "stock"
    """
    cfg = get_config()
    min_cols_map = {
        "sales":  cfg["csv_validation"]["min_sales_columns"],
        "master": cfg["csv_validation"]["min_master_columns"],
        "stock":  cfg["csv_validation"]["min_stock_columns"],
    }
    min_cols = min_cols_map.get(csv_type, 0)
    actual = len(df.columns)
    if actual < min_cols:
        msg = (f"【警告】{filename} のカラム数が想定より少ないです "
               f"（期待: {min_cols}以上, 実際: {actual}）。"
               f"CSVの出力形式が変更された可能性があります。")
        log(msg)
        return False
    return True


# =====================================
# Excel列幅自動調整
# =====================================
def auto_fit_columns(ws, max_width=40):
    """ワークシートの列幅をセル内容に応じて自動調整"""
    for col_cells in ws.columns:
        length = 0
        for cell in col_cells:
            try:
                val_len = len(str(cell.value)) if cell.value else 0
                length = max(length, val_len)
            except Exception:
                pass
        col_letter = col_cells[0].column_letter
        ws.column_dimensions[col_letter].width = min(length + 2, max_width)


# =====================================
# ヘルパー関数
# =====================================
def get_season(month):
    """月からシーズンを返す（SS: 3-8月、AW: 9-2月）"""
    return "SS" if 3 <= month <= 8 else "AW"

def get_price_band(price):
    """価格帯を返す"""
    if price <= 5000:   return "① ~5,000円"
    if price <= 15000:  return "② 5,001~15,000円"
    if price <= 30000:  return "③ 15,001~30,000円"
    return "④ 30,001円~"


# =====================================
# DBバックアップ
# =====================================
def backup_db():
    """DB実行前バックアップ（7日分保持）"""
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return

    cfg = get_config()
    backup_dir  = cfg["paths"]["backup_dir"]
    keep_days   = cfg["paths"]["backup_keep_days"]
    os.makedirs(backup_dir, exist_ok=True)

    # バックアップ作成
    today_str = datetime.now().strftime("%Y%m%d")
    backup_name = f"fjallraven_md_v3_{today_str}.bak"
    backup_path = os.path.join(backup_dir, backup_name)
    shutil.copy2(db_path, backup_path)
    log(f"  DBバックアップ作成: {backup_name}")

    # 古いバックアップ削除
    cutoff = datetime.now() - timedelta(days=keep_days)
    for f in glob.glob(os.path.join(backup_dir, "*.bak")):
        if datetime.fromtimestamp(os.path.getmtime(f)) < cutoff:
            os.remove(f)
            log(f"  古いバックアップ削除: {os.path.basename(f)}")
