import sqlite3
import pandas as pd

# アップロードに用いるdbmap
db_map = {
  'table_name': 'customers_id',
  'column': {
    'id': {'type': 'integer', 'index': 1},
    'id_A': {'type': 'text', 'index': 2},
    'id_E': {'type': 'text', 'index': 3},
    'name': {'type': 'text', 'index': 4},
  }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
  conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_all.db')
  conn.row_factory = sqlite3.Row
  cur = conn.cursor()

  # 関数db_initを呼び出し
  db_init(db_map, cur)
  col_name = []
  val = []
  table_name = db_map['table_name']

  for k, v in db_map['column'].items():
    col_name.append(k)
    val.append(v['index'])
  
  col_names = ','.join(col_name)

  for r in db.iterrows():
    values = []
    for v in val:
      cell_val = r[1][v-1]
      if type(cell_val) is not str and type(cell_val) is not int and cell_val is not None:
        values.append(str(cell_val))
      else:
        values.append(cell_val)

    place_holder = ','.join('?'*len(values))
    sql = f'INSERT INTO {table_name} ({col_names}) VALUES({place_holder})'
    cur.execute(sql, tuple(values))

  conn.commit()
  cur.close()
  conn.close()


# 既存のテーブルの値を削除
def db_init(db_map, db):
  param = []
  table_name = db_map['table_name']

  for k, v in db_map['column'].items():
    param.append(f"{k} {v['type']}")

  params = ','.join(param)

  db.execute(f'CREATE TABLE IF NOT EXISTS {table_name}({params})')
  db.execute(f'DELETE FROM {table_name}')


# データベースへ接続
conn=sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_all.db')

# id_A, id_Eテーブルを自然結合
data = pd.read_sql_query("select * from id_E natural left outer join id_A", conn) 
data = data.reindex(columns=['id_A', 'id_E', 'name'])
data.insert(loc = 0, column = 'id', value = None)

# idを追加
for index, row in data.iterrows():
  data.iloc[index, 0] = index

# データベースとの接続を切断
conn.close()

# 関数db_insertを呼び出し
db_insert(data, db_map)