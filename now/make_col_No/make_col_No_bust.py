import pandas as pd
import sqlite3

# アップロードに用いるdbmap
db_map = {
    'table_name': 'questionnaire_bust_4',
    'column': {
        'id': {'type': 'int', 'index': 1},
        'No': {'type': 'text', 'index': 2},
        '来店年月日': {'type': 'text', 'index': 3},
        '氏名': {'type': 'text', 'index': 4},
        'フリガナ': {'type': 'text', 'index': 5},
        '生年月日': {'type': 'text', 'index': 6},
        '年齢': {'type': 'text', 'index': 7},
        '郵便番号': {'type': 'text', 'index': 8},
        '住所': {'type': 'text', 'index': 9},
        'TEL_自宅': {'type': 'text', 'index': 10},
        'TEL_携帯': {'type': 'text', 'index': 11},
        'メールアドレス': {'type': 'text', 'index': 12},
        '職業': {'type': 'text', 'index': 13},
        '血液型': {'type': 'text', 'index': 14},
        '結婚': {'type': 'text', 'index': 15},
        '家族構成': {'type': 'text', 'index': 16},
        'DM': {'type': 'text', 'index': 17},
        '知った理由': {'type': 'text', 'index': 18},
        '来店のきっかけ': {'type': 'text', 'index': 19},
        'アレルギー': {'type': 'text', 'index': 20},
        'アレルギー_有': {'type': 'text', 'index': 21},
        '治療': {'type': 'text', 'index': 22},
        '治療_有': {'type': 'text', 'index': 23},
        '服用薬': {'type': 'text', 'index': 24},
        '服用薬_有': {'type': 'text', 'index': 25},
        '健康状態': {'type': 'text', 'index': 26},
        '健康状態_悪い': {'type': 'text', 'index': 27},
        '精神状態': {'type': 'text', 'index': 28},
        '精神状態_悪い': {'type': 'text', 'index': 29},
        '月経': {'type': 'text', 'index': 30},
        '周期': {'type': 'text', 'index': 31},
        'ダイエット経験': {'type': 'text', 'index': 32},
        'かぶれの経験': {'type': 'text', 'index': 33},
        'かぶれ_有': {'type': 'text', 'index': 34},
        'バストの悩み': {'type': 'text', 'index': 35},
        'バストの悩み_その他': {'type': 'text', 'index': 36},
        '気になった時期': {'type': 'text', 'index': 37},
        '気になった理由': {'type': 'text', 'index': 38},
        '改善策': {'type': 'text', 'index': 39},
        '改善策_有': {'type': 'text', 'index': 40},
        '睡眠時間': {'type': 'text', 'index': 41},
        '平均睡眠時間': {'type': 'text', 'index': 42},
        '睡眠型': {'type': 'text', 'index': 43},
        '喫煙': {'type': 'text', 'index': 44},
        '喫煙_本数': {'type': 'text', 'index': 45},
        '出産経験': {'type': 'text', 'index': 46},
        '授乳経験': {'type': 'text', 'index': 47},
        'いつまでに': {'type': 'text', 'index': 48},
        '美しくなりたい理由': {'type': 'text', 'index': 49},
        'どんな状態に': {'type': 'text', 'index': 50},
        'どの程度の施術_コース': {'type': 'text', 'index': 51},
        'どの程度の施術_予算': {'type': 'text', 'index': 52},
        '予算金額': {'type': 'text', 'index': 53},
        'バストケア経験': {'type': 'text', 'index': 54},
        'バストケア経験_サロン': {'type': 'text', 'index': 55},
        'バストケア経験_費用': {'type': 'text', 'index': 56},
        'バストケア経験_期間': {'type': 'text', 'index': 57},
        'バストケア経験_結果': {'type': 'text', 'index': 58},
        'エステサロン経験': {'type': 'text', 'index': 59},
        'エステサロン経験_コース': {'type': 'text', 'index': 60},
        'エステサロン経験_サロン': {'type': 'text', 'index': 61},
        'エステサロン経験_費用': {'type': 'text', 'index': 62},
        'エステサロン経験_期間': {'type': 'text', 'index': 63},
        'エステサロン経験_結果': {'type': 'text', 'index': 64},
        '該当項目': {'type': 'text', 'index': 65},
        'カウンセリング': {'type': 'text', 'index': 66},
        'カウンセリング_その他': {'type': 'text', 'index': 67},
        '他のコース': {'type': 'text', 'index': 68},
        '他のコース_その他': {'type': 'text', 'index': 69},
        '現在のサイズ': {'type': 'text', 'index': 70},
        '理想のサイズ': {'type': 'text', 'index': 71},
        'AYAを選んだ理由': {'type': 'text', 'index': 72},
        '契約': {'type': 'text', 'index': 73},
        '契約内容': {'type': 'text', 'index': 74},
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
# データベースの値を読み込み
df_cus = pd.read_sql_query('select * from customers', conn)
# データベースとの接続を切断
conn.close()

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')
# データベースの値を読み込み
df_body = pd.read_sql_query('select * from questionnaire_bust_4', conn)
# データベースとの接続を切断
conn.close()

# 不足している列を追加
df_body.insert(loc = 1, column = 'No', value = None)

# df_hairremovalにdf_cusに含まれるid(No)を追加
for index, row in df_body.iterrows():
    for i, r in df_cus.iterrows():
        flag = 0
        if row[3].replace(' ', '').replace('　', '') == r[1].replace(' ', '').replace('　', ''):
            flag += 1
        # if row[4] == r[2]:
        #     flag += 1
        # if row[7] == r[3]:
        #     flag += 1
        # if row[8] == r[4]:
        #     flag += 1
        # if row[10] == r[5]:
        #     flag += 1
        if flag == 1:
            df_body.iloc[index,1] = r[0]
            print('match')
            break

# 関数db_insertの呼び出し
db_insert(df_body, db_map) 