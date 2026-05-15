import sqlite3
import pandas as pd

conn = sqlite3.connect('fjallraven_md_v3.db')
df = pd.read_sql("SELECT DISTINCT [3rd Item No.], [商品名] FROM sales_raw WHERE [商品名] LIKE '%Kanken%'", conn)
print(df.head(20))
conn.close()
