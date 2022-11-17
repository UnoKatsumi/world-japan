import sqlite3
import pandas as pd

#salon_all.dbに接続
db_name = "S:\個人作業用\アルバイト\ワールドジャパン\sqlite3\salon_all.db"
conn = sqlite3.connect(db_name)

#SQLiteを操作するためのカーソルを作成
cur = conn.cursor()

#テーブル名を入力し、テーブルを選択
table_name = input()
cur.execute('select * from ' + table_name)

#descを取得
descr = cur.description

#カラムリストの配列を作る
column_list = []

#カラム名をカラムリストに格納
for desc in descr:
  column_list.append(desc[0])

for column_name in column_list:
  sql_query = '''
              SELECT 
                count(*) 
              FROM
                contract
              where ''' + column_name + ''' <> 'nan' AND ''' + column_name + ''' <> 'NaT' AND ''' + column_name + ''' <> '';
          '''
  rt = pd.read_sql(sql_query, conn)
  print(rt["count(*)"][0])

conn.close