import pandas as pd
import sqlite3

# アップロードに用いるdbmap
db_map = {
    'table_name': 'questionnaire_hairremoval_4',
    'column': {
        'id': {'type': 'int', 'index': 1},
        'No': {'type': 'text', 'index': 2},
        '来店年月日': {'type': 'text', 'index': 3},
        '氏名': {'type': 'text', 'index': 4},
        'フリガナ': {'type': 'text', 'index': 5},
        '生年月日': {'type': 'text', 'index': 6},
        '既婚・未婚': {'type': 'text', 'index': 7},
        '郵便番号': {'type': 'text', 'index': 8},
        '住所': {'type': 'text', 'index': 9},
        'TEL_自宅': {'type': 'text', 'index': 10},
        'TEL_携帯': {'type': 'text', 'index': 11},
        'メールアドレス': {'type': 'text', 'index': 12},
        '職業': {'type': 'text', 'index': 13},
        '趣味': {'type': 'text', 'index': 14},
        '知った理由': {'type': 'text', 'index': 15},
        '知った理由_その他': {'type': 'text', 'index': 16},
        'DM': {'type': 'text', 'index': 17},
        '電話対応': {'type': 'text', 'index': 18},
        '第一印象': {'type': 'text', 'index': 19},
        '脱毛経験': {'type': 'text', 'index': 20},
        '脱毛経験_サロン名': {'type': 'text', 'index': 21},
        '脱毛経験_期間': {'type': 'text', 'index': 22},
        '脱毛経験_時期': {'type': 'text', 'index': 23},
        '脱毛経験_方法': {'type': 'text', 'index': 24},
        '脱毛経験_方法_その他': {'type': 'text', 'index': 25},
        '脱毛経験_部位': {'type': 'text', 'index': 26},
        '脱毛経験_部位_その他': {'type': 'text', 'index': 27},
        '脱毛経験_料金': {'type': 'text', 'index': 28},
        '自己処理部位': {'type': 'text', 'index': 29},
        '自己処理部位_その他': {'type': 'text', 'index': 30},
        '自己処理方法': {'type': 'text', 'index': 31},
        '自己処理方法_その他': {'type': 'text', 'index': 32},
        '希望脱毛箇所': {'type': 'text', 'index': 33},
        '希望脱毛箇所_その他': {'type': 'text', 'index': 34},
        '日焼け': {'type': 'text', 'index': 35},
        '日焼け_いつ': {'type': 'text', 'index': 36},
        '日焼け_理由': {'type': 'text', 'index': 37},
        '期待・要望': {'type': 'text', 'index': 38},
        '身近でムダ毛': {'type': 'text', 'index': 39},
        '身近でムダ毛_人数': {'type': 'text', 'index': 40},
        '興味のあること': {'type': 'text', 'index': 41},
        '興味のあること_その他': {'type': 'text', 'index': 42},
        '取り入れてほしいこと': {'type': 'text', 'index': 43},
        '取り入れてほしいこと_その他': {'type': 'text', 'index': 44},
        '妊娠の予定': {'type': 'text', 'index': 45},
        '病気治療': {'type': 'text', 'index': 46},
        '病気治療_病名': {'type': 'text', 'index': 47},
        'アレルギー': {'type': 'text', 'index': 48},
        'アレルギー_有': {'type': 'text', 'index': 49},
        'アレルギー_その他': {'type': 'text', 'index': 50},
        'ペースメーカー': {'type': 'text', 'index': 51},
        '薬・サプリメント': {'type': 'text', 'index': 52},
        'お肌のタイプ': {'type': 'text', 'index': 53},
        'お肌のタイプ_その他': {'type': 'text', 'index': 54},
        '生理': {'type': 'text', 'index': 55},
        '生理周期': {'type': 'text', 'index': 56},
        '生理痛': {'type': 'text', 'index': 57},
        '生理時の服用薬': {'type': 'text', 'index': 58},
        '生理時の服用薬_有': {'type': 'text', 'index': 59},
        '最終月経': {'type': 'text', 'index': 60},
        '授乳中': {'type': 'text', 'index': 61},
        '医師からの注意': {'type': 'text', 'index': 62},
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
df_hairremoval = pd.read_sql_query('select * from questionnaire_hairremoval_4', conn)
# データベースとの接続を切断
conn.close()

# 不足している列を追加
df_hairremoval.insert(loc = 1, column = 'No', value = None)

# df_hairremovalにdf_cusに含まれるid(No)を追加
for index, row in df_hairremoval.iterrows():
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
            df_hairremoval.iloc[index,1] = r[0]
            print('match')
            break

# 関数db_insertの呼び出し
db_insert(df_hairremoval, db_map) 