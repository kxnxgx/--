# V4 実装ロードマップ

> 作成日: 2026-05-17  
> 方針: 短期（すぐできる）→ 中期（1〜2ヶ月）→ 長期（3ヶ月〜）の段階的拡張

---

## フェーズ概要

| フェーズ | 内容 | 優先度 | 難易度 | 対応課題 |
|----------|------|--------|--------|----------|
| **V4.0-A** | 複数ブランド対応 + バッチUX修正 | ★★★ | 低 | 課題C |
| **V4.0-B** | DB履歴データ活用 + 在庫ヘルス推移シート | ★★★ | 低〜中 | 課題A |
| **V4.1** | 需要予測 + 発注アラート + 機会損失額 | ★★☆ | 中〜高 | 課題B |
| **V4.2** | 値下げシミュレーション + AI連携改善 | ★☆☆ | 高 | 課題D |

---

## V4.0-A: 複数ブランド対応（最優先・半日で完了可）

### 変更ファイル: `v3/config.json`
```json
// 変更前
"target_brand": "FRV"

// 変更後  
"target_brands": ["FRV"],
"brand_display_names": { "FRV": "Fjällräven" }
```

### 変更ファイル: `v3/main.py`
```python
# brands をループして各ブランドのレポートを生成
for brand in config["target_brands"]:
    logger.info(f"[{brand}] 処理開始")
    result = run_single_brand(brand, config, args)
    logger.info(f"[{brand}] 完了")
```

### 変更ファイル: `v3/実行ボタン.bat`
```batch
@echo off
chcp 65001 > nul
python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] エラーが発生しました。上記のメッセージを確認してください。
    pause
) else (
    echo.
    echo [完了] レポートが生成されました。
    pause
)
```
> `pause` を追加するだけでエラー時に画面が閉じなくなる（5分で対応可能）

---

## V4.0-B: DB履歴データ活用（最優先・1〜2日）

### 追加ファイル: `v3/db_manager.py` に関数追加

```python
def get_stock_trend(weeks: int = 8) -> pd.DataFrame:
    """
    過去N週分の在庫スナップショット履歴を取得
    stock_history テーブルから実行日×商品コードの在庫数を集計
    """
    conn = sqlite3.connect(get_db_path())
    query = """
        SELECT
            実行日,
            中分類,
            AVG(WOS) AS avg_wos,
            SUM(在庫金額) AS total_stock_value
        FROM stock_history
        WHERE 実行日 >= DATE('now', ?)
        GROUP BY 実行日, 中分類
        ORDER BY 実行日
    """
    df = pd.read_sql_query(query, conn, params=(f'-{weeks * 7} days',))
    conn.close()
    return df

def get_inventory_turnover_monthly() -> pd.DataFrame:
    """
    月次在庫回転率を算出
    IT = 月次販売数合計 / その月の平均在庫数
    """
    # 実装はstock_historyとsales_rawを結合して算出
    pass
```

### 追加シート: `v3/excel_writer.py` に `write_stock_health_trend()` を追加

**出力内容:**
1. WOS推移 折れ線グラフ（中分類別・直近8週・赤/黄ライン入り）
2. 在庫金額残高 棒グラフ（週次）
3. 在庫回転率（IT）月次推移グラフ
4. 生データテーブル（実行日×中分類×WOS値）

---

## V4.1: 需要予測 + 発注アラート + 機会損失（1〜2ヶ月）

### 新規ファイル: `v3/forecaster.py`

**アルゴリズム: 乗法型季節指数法**
- ベーストレンド = 直近4週の週次販売数の移動平均
- 季節指数 = DB累積3ヶ月以上のデータから月別算出
- 予測値 = ベーストレンド × 季節指数
- 安全在庫 = Z値（95%SL=1.65）× 販売数標準偏差 × √リードタイム週数
- 推奨発注数 = 予測販売数 × リードタイム週数 + 安全在庫 - 現在庫

**注意事項:**
- DB蓄積が3ヶ月未満の場合はフォールバック: 季節指数=1.0（フラット）で計算
- Excelシートには「⚠ 直近実績ベース・精度保証なし」の注釈を必ず入れる

### 新規関数: `v3/analyzer.py` に `calc_opportunity_loss()` を追加

**ロジック:**
- 品番単位で「最も売れているサイズ比率」を算出
- 在庫0サイズの推定損失 = 同品番平均サイズ比率 × 欠品日数 × 売価
- 出力: 推定機会損失額の上位20SKU

---

## V4.2: 値下げシミュレーション + AI連携改善（3ヶ月〜）

### 新規関数: `v3/analyzer.py` に `calc_markdown_simulation()` を追加

**ロジック:**
- 価格弾力性（デフォルト -1.5）を `config.json` で調整可能にする
- 値下げ率 10% / 20% / 30% の3パターンで粗利影響をシミュレーション
- 「推奨判定」列: 値下げ後粗利 > 値下げ前粗利 なら「◎値下げ有効」

### AI連携改善（課題D）

**方針: Google Apps Script 経由の半自動化**
1. Python実行後に `notebook_data/` のCSVをGoogle Driveへ自動アップロード
2. GASでGemini APIを呼び出し、定型プロンプトで分析を実行
3. 結果をGoogle Sheetsに書き出し、SlackやLINE WORKSへ通知

---

## .gitignore 追加ルール（V4で新たに生成されるファイル用）

```gitignore
# V4 新規生成ファイル
v4/reports/
v4/logs/
v4/tmp/
v4/notebook_data/
*.db
*.db.bak
backups/
```

---

## 作業ブランチ運用方針

| ブランチ名 | 用途 |
|------------|------|
| `main` | 安定稼働版（V3現行） |
| `feature/v4-dashboard-spec` | 本ブランチ: フロント仕様・モック整理 |
| `feature/v4-multi-brand` | V4.0-A: 複数ブランド対応 |
| `feature/v4-stock-trend` | V4.0-B: DB履歴活用・在庫ヘルスシート |
| `feature/v4-forecaster` | V4.1: 需要予測・発注アラート |
| `feature/v4-markdown-sim` | V4.2: 値下げシミュレーション |
