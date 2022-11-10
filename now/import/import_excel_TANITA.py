import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'TANITA',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        '測定日': {'type': 'text', 'index': 3},
        '測定時間': {'type': 'text', 'index': 4},
        'ID': {'type': 'text', 'index': 5},
        '年齢': {'type': 'text', 'index': 6},
        '身長': {'type': 'text', 'index': 7},
        '着衣量': {'type': 'text', 'index': 8},
        'スタンダード・アスリート': {'type': 'text', 'index': 9},
        '性別': {'type': 'text', 'index': 10},
        '体重_結果': {'type': 'text', 'index': 11},
        '体重_標準範囲': {'type': 'text', 'index': 12},
        '体重_目標値': {'type': 'text', 'index': 13},
        '体重_差': {'type': 'text', 'index': 14},
        '体脂肪率_結果': {'type': 'text', 'index': 15},
        '体脂肪率_標準範囲': {'type': 'text', 'index': 16},
        '体脂肪率_目標値': {'type': 'text', 'index': 17},
        '体脂肪率_差': {'type': 'text', 'index': 18},
        '脂肪量_結果': {'type': 'text', 'index': 19},
        '脂肪量_標準範囲': {'type': 'text', 'index': 20},
        '脂肪量_目標値': {'type': 'text', 'index': 21},
        '脂肪量_差': {'type': 'text', 'index': 22},
        '除脂肪量_結果': {'type': 'text', 'index': 23},
        '除脂肪量_標準範囲': {'type': 'text', 'index': 24},
        '除脂肪量_目標値': {'type': 'text', 'index': 25},
        '除脂肪量_差': {'type': 'text', 'index': 26},
        '筋肉量_結果': {'type': 'text', 'index': 27},
        '筋肉量_標準範囲': {'type': 'text', 'index': 28},
        '筋肉量_目標値': {'type': 'text', 'index': 29},
        '筋肉量_差': {'type': 'text', 'index': 30},
        '体水分量_結果': {'type': 'text', 'index': 31},
        '体水分量_標準範囲': {'type': 'text', 'index': 32},
        '体水分量_目標値': {'type': 'text', 'index': 33},
        '体水分量_差': {'type': 'text', 'index': 34},
        '推定骨量_結果': {'type': 'text', 'index': 35},
        '推定骨量_標準範囲': {'type': 'text', 'index': 36},
        '推定骨量_目標値': {'type': 'text', 'index': 37},
        '推定骨量_差': {'type': 'text', 'index': 38},
        'BMI': {'type': 'text', 'index': 39},
        'アスリート指数': {'type': 'text', 'index': 40},
        '基礎代謝量': {'type': 'text', 'index': 41},
        '内臓脂肪レベル': {'type': 'text', 'index': 42},
        '筋肉総合評価_体幹部': {'type': 'text', 'index': 43},
        '筋肉総合評価_右腕': {'type': 'text', 'index': 44},
        '筋肉総合評価_右脚': {'type': 'text', 'index': 45},
        '筋肉総合評価_左腕': {'type': 'text', 'index': 46},
        '筋肉総合評価_左脚': {'type': 'text', 'index': 47},
        '体脂肪総合評価_体幹部_割合': {'type': 'text', 'index': 48},
        '体脂肪総合評価_体幹部_kg': {'type': 'text', 'index': 49},
        '体脂肪総合評価_右腕_割合': {'type': 'text', 'index': 50},
        '体脂肪総合評価_右腕_kg': {'type': 'text', 'index': 51},
        '体脂肪総合評価_右脚_割合': {'type': 'text', 'index': 52},
        '体脂肪総合評価_右脚_kg': {'type': 'text', 'index': 53},
        '体脂肪総合評価_左腕_割合': {'type': 'text', 'index': 54},
        '体脂肪総合評価_左腕_kg': {'type': 'text', 'index': 55},
        '体脂肪総合評価_左脚_割合': {'type': 'text', 'index': 56},
        '体脂肪総合評価_左脚_kg': {'type': 'text', 'index': 57}
    }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
    # conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_K.db')
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


# excelファイルの内容を読み込み
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(潮永さん）/【K】TANITA.xlsx', sheet_name='Sheet1', header=4)
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【K】TANITA.xlsx', sheet_name='Sheet1', header=4)
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【K】TANITA.xlsx', sheet_name='Sheet1', header=4)
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【K】TANITA.xlsx', sheet_name='Sheet1', header=4)
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【K】TANITA.xlsx', sheet_name='Sheet1', header=4)

#データフレームを結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])

wb = wb.rename(columns={'会員番号' : 'No'})

# 不要な値を含む行を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last')

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)