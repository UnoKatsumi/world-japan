import pandas as pd
import sqlite3

# アップロードに用いるdbmap
db_map = {
    'table_name': 'id_A',
    'column': {
        'id_A': {'type': 'text', 'index': 1},
        'name': {'type': 'text', 'index': 2},
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

    flag = 0

    for r in db.iterrows():
        values = []
        if r[1][0] is None:
            flag += 1
        else:
            flag = 0
            for v in val:
                cell_val = r[1][v-1]
                if type(cell_val) is not str and type(cell_val) is not int and cell_val is not None:
                    values.append(str(cell_val))
                else:
                    values.append(cell_val)

            place_holder = ','.join('?'*len(values))
            sql = f'INSERT INTO {table_name} ({col_names}) VALUES({place_holder})'
            cur.execute(sql, tuple(values))
        
        if flag == 2:
            break

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
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_all.db')

# データベースの内容を読み込み
df_bd = pd.read_sql_query('select * from questionnaire_body_4', conn)
df_bd = df_bd.loc[:, ['No', '氏名']]
df_bs = pd.read_sql_query('select * from questionnaire_bust_4', conn)
df_bs = df_bs.loc[:, ['No', '氏名']]
df_fa = pd.read_sql_query('select * from questionnaire_facial_4', conn)
df_fa = df_fa.loc[:, ['No', '氏名']]
df_hr = pd.read_sql_query('select * from questionnaire_hairremoval_4', conn)
df_hr = df_hr.loc[:, ['No', '氏名']]

# 「氏名」カラムの値に含まれる空白を削除
def name_replace(df):
    for index, row in df.iterrows():
        df.iloc[index, 1] = row[1].replace(' ', '').replace('　', '')

# 関数name_replaceを呼び出し
name_replace(df_bd)
name_replace(df_bs)
name_replace(df_fa)
name_replace(df_hr)

# データフレームを結合
df_A = pd.concat([df_bd, df_bs, df_fa, df_hr])

# 不要な値を含む行を削除
df_A = df_A.drop_duplicates(subset='No', keep='last')

# データベースとの接続を切断
conn.close()

# 関数db_insertを呼び出し
db_insert(df_A, db_map)

