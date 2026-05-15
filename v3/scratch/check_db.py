import sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = sqlite3.connect('fjallraven_md_v3.db')
cur = conn.cursor()
tables = cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
for t in tables:
    name = t[0]
    count = cur.execute(f'SELECT COUNT(*) FROM [{name}]').fetchone()[0]
    cols = cur.execute(f'PRAGMA table_info([{name}])').fetchall()
    col_names = [c[1] for c in cols]
    print(f'{name}: {count} rows, {len(col_names)} cols')
    if len(col_names) > 10:
        print(f'  cols: {col_names[:10]}...')
    else:
        print(f'  cols: {col_names}')
conn.close()
