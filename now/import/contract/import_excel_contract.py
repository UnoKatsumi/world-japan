import pandas as pd
import sqlite3

# アップロードに用いるdbmap
db_map = {
    'table_name': 'contract',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        'フリガナ': {'type': 'text', 'index': 3},
        '契約日': {'type': 'text', 'index': 4},
        '生年月日_和暦': {'type': 'text', 'index': 5},
        '生年月日_西暦': {'type': 'text', 'index': 6},
        '年齢': {'type': 'int', 'index': 7},
        '顧客コード': {'type': 'int', 'index': 8},
        '郵便番号': {'type': 'text', 'index': 9},
        '住所': {'type': 'text', 'index': 10},
        'TEL_自宅': {'type': 'text', 'index': 11},
        'TEL_携帯': {'type': 'text', 'index': 12},
        '職業': {'type': 'text', 'index': 13},
        '会社・学校名称': {'type': 'text', 'index': 14},
        '所在地_郵便番号': {'type': 'text', 'index': 15},
        '所在地_住所': {'type': 'text', 'index': 16},
        '入会期間': {'type': 'text', 'index': 17},
        '入会金': {'type': 'int', 'index': 18},
        '役務提供機関': {'type': 'text', 'index': 19},
        'コース名_1行目': {'type': 'text', 'index': 20},
        'コース名_2行目': {'type': 'text', 'index': 21},
        'コース名_3行目': {'type': 'text', 'index': 22},
        'コース名_4行目': {'type': 'text', 'index': 23},
        'コース名_5行目': {'type': 'text', 'index': 24},
        'コース名_6行目': {'type': 'text', 'index': 25},
        '時間_1行目': {'type': 'int', 'index': 26},
        '時間_2行目': {'type': 'int', 'index': 27},
        '時間_3行目': {'type': 'int', 'index': 28},
        '時間_4行目': {'type': 'int', 'index': 29},
        '時間_5行目': {'type': 'int', 'index': 30},
        '時間_6行目': {'type': 'int', 'index': 31},
        '単価_1行目': {'type': 'int', 'index': 32},
        '単価_2行目': {'type': 'int', 'index': 33},
        '単価_3行目': {'type': 'int', 'index': 34},
        '単価_4行目': {'type': 'int', 'index': 35},
        '単価_5行目': {'type': 'int', 'index': 36},
        '単価_6行目': {'type': 'int', 'index': 37},
        '回数_1行目': {'type': 'int', 'index': 38},
        '回数_2行目': {'type': 'int', 'index': 39},
        '回数_3行目': {'type': 'int', 'index': 40},
        '回数_4行目': {'type': 'int', 'index': 41},
        '回数_5行目': {'type': 'int', 'index': 42},
        '回数_6行目': {'type': 'int', 'index': 43},
        '回数合計': {'type': 'int', 'index': 44},
        '総時間数_1行目': {'type': 'int', 'index': 45},
        '総時間数_2行目': {'type': 'int', 'index': 46},
        '総時間数_3行目': {'type': 'int', 'index': 47},
        '総時間数_4行目': {'type': 'int', 'index': 48},
        '総時間数_5行目': {'type': 'int', 'index': 49},
        '総時間数_6行目': {'type': 'int', 'index': 50},
        '総時間合計': {'type': 'int', 'index': 51},
        '金額_1行目': {'type': 'int', 'index': 52},
        '金額_2行目': {'type': 'int', 'index': 53},
        '金額_3行目': {'type': 'int', 'index': 54},
        '金額_4行目': {'type': 'int', 'index': 55},
        '金額_5行目': {'type': 'int', 'index': 56},
        '金額_6行目': {'type': 'int', 'index': 57},
        '金額合計': {'type': 'int', 'index': 58},
        '商品名_推奨商品_1行目': {'type': 'text', 'index': 59},
        '商品名_推奨商品_2行目': {'type': 'text', 'index': 60},
        '商品名_推奨商品_3行目': {'type': 'text', 'index': 61},
        '商品名_推奨商品_4行目': {'type': 'text', 'index': 62},
        '商品名_推奨商品_5行目': {'type': 'text', 'index': 63},
        '商品名_推奨商品_6行目': {'type': 'text', 'index': 64},
        '種類_推奨商品_1行目': {'type': 'text', 'index': 65},
        '種類_推奨商品_2行目': {'type': 'text', 'index': 66},
        '種類_推奨商品_3行目': {'type': 'text', 'index': 67},
        '種類_推奨商品_4行目': {'type': 'text', 'index': 68},
        '種類_推奨商品_5行目': {'type': 'text', 'index': 69},
        '単価_推奨商品_1行目': {'type': 'int', 'index': 70},
        '単価_推奨商品_2行目': {'type': 'int', 'index': 71},
        '単価_推奨商品_3行目': {'type': 'int', 'index': 72},
        '単価_推奨商品_4行目': {'type': 'int', 'index': 73},
        '単価_推奨商品_5行目': {'type': 'int', 'index': 74},
        '単価_推奨商品_6行目': {'type': 'int', 'index': 75},
        '数量_推奨商品_1行目': {'type': 'int', 'index': 76},
        '数量_推奨商品_2行目': {'type': 'int', 'index': 77},
        '数量_推奨商品_3行目': {'type': 'int', 'index': 78},
        '数量_推奨商品_4行目': {'type': 'int', 'index': 79},
        '数量_推奨商品_5行目': {'type': 'int', 'index': 80},
        '数量_推奨商品_6行目': {'type': 'int', 'index': 81},
        '数量合計_推奨商品': {'type': 'int', 'index': 82},
        '金額_推奨商品_1行目': {'type': 'int', 'index': 83},
        '金額_推奨商品_2行目': {'type': 'int', 'index': 84},
        '金額_推奨商品_3行目': {'type': 'int', 'index': 85},
        '金額_推奨商品_4行目': {'type': 'int', 'index': 86},
        '金額_推奨商品_5行目': {'type': 'int', 'index': 87},
        '金額_推奨商品_6行目': {'type': 'int', 'index': 88},
        '金額合計_推奨商品': {'type': 'int', 'index': 89},
        'お支払い総合計金額_推奨商品': {'type': 'int', 'index': 90},
        'お支払い方法及びお支払い時期': {'type': 'text', 'index': 91},
        'クレジットカード支払い回数': {'type': 'int', 'index': 92},
        'ショッピングクレジット支払い回数': {'type': 'int', 'index': 93},
        'お支払い時期_1行目': {'type': 'text', 'index': 94},
        'お支払い時期_2行目': {'type': 'text', 'index': 95},
        'クレジット会社名': {'type': 'text', 'index': 96},
        '引き落とし開始': {'type': 'text', 'index': 97},
        '引き落とし日': {'type': 'text', 'index': 98},
        '金額_分割払手数料含む_1行目': {'type': 'int', 'index': 99},
        '金額_分割払手数料含む_2行目': {'type': 'int', 'index': 100},
        '初回': {'type': 'int', 'index': 101},
        '最終回': {'type': 'int', 'index': 102},
        '通常各回': {'type': 'int', 'index': 103}
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
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_sub.db')
conn.row_factory = sqlite3.Row

# データベースの値を読み込み
df_shionaga = pd.read_sql_query('select * from contract_shionaga', conn)
df_watabe = pd.read_sql_query('select * from contract_watabe', conn)
df_kashiwaba = pd.read_sql_query('select * from contract_kashiwaba', conn)
df_keino = pd.read_sql_query('select * from contract_keino', conn)
df_satou = pd.read_sql_query('select * from contract_satou', conn)

# データフレームを結合
df = pd.concat([df_shionaga, df_watabe, df_kashiwaba, df_keino, df_satou])

# 不要な値を含む行を削除
df = df.drop_duplicates(subset='No', keep='last')

# データベースとの接続を切断
conn.close()

# 関数db_insertの呼び出し
db_insert(df, db_map)