import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'ACCUNIQ',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        '測定日': {'type': 'text', 'index': 3},
        '測定時間': {'type': 'text', 'index': 4},
        '身長': {'type': 'text', 'index': 5},
        '体重': {'type': 'text', 'index': 6},
        '性別': {'type': 'text', 'index': 7},
        '年齢': {'type': 'text', 'index': 8},
        '体成分分析_体重': {'type': 'text', 'index': 9},
        '体成分分析_標準体重': {'type': 'text', 'index': 10},
        '体成分分析_美容体重': {'type': 'text', 'index': 11},
        '体成分分析_除脂肪量': {'type': 'text', 'index': 12},
        '体成分分析_除脂肪量_評価': {'type': 'text', 'index': 13},
        '体成分分析_筋肉量': {'type': 'text', 'index': 14},
        '体成分分析_筋肉量_評価': {'type': 'text', 'index': 15},
        '体成分分析_体水分量': {'type': 'text', 'index': 16},
        '体成分分析_体水分量_評価': {'type': 'text', 'index': 17},
        '体成分分析_タンパク質量': {'type': 'text', 'index': 18},
        '体成分分析_タンパク質量_評価': {'type': 'text', 'index': 19},
        '体成分分析_ミネラル': {'type': 'text', 'index': 20},
        '体成分分析_ミネラル_評価': {'type': 'text', 'index': 21},
        '体成分分析_体脂肪量': {'type': 'text', 'index': 22},
        '体成分分析_体脂肪量_評価': {'type': 'text', 'index': 23},
        '適正範囲_体重': {'type': 'text', 'index': 24},
        '適正範囲_除脂肪量': {'type': 'text', 'index': 25},
        '適正範囲_体脂肪量': {'type': 'text', 'index': 26},
        '適正範囲_筋肉量': {'type': 'text', 'index': 27},
        '適正範囲_骨格筋量': {'type': 'text', 'index': 28},
        '適正範囲_体水分量': {'type': 'text', 'index': 29},
        '適正範囲_タンパク質量': {'type': 'text', 'index': 30},
        '適正範囲_ミネラル': {'type': 'text', 'index': 31},
        '身体健康度_体重': {'type': 'text', 'index': 32},
        '身体健康度_体重_評価': {'type': 'text', 'index': 33},
        '身体健康度_骨格筋量': {'type': 'text', 'index': 34},
        '身体健康度_骨格筋量_評価': {'type': 'text', 'index': 35},
        '身体健康度_体脂肪量': {'type': 'text', 'index': 36},
        '身体健康度_体脂肪量_評価': {'type': 'text', 'index': 37},
        '肥満分析_BMI': {'type': 'text', 'index': 38},
        '肥満分析_BMI_評価': {'type': 'text', 'index': 39},
        '肥満分析_体脂肪率': {'type': 'text', 'index': 40},
        '肥満分析_体脂肪率_評価': {'type': 'text', 'index': 41},
        '腹部肥満分析_腹部肥満率': {'type': 'text', 'index': 42},
        '腹部肥満分析_腹部肥満率_評価': {'type': 'text', 'index': 43},
        '腹部肥満分析_内臓脂肪レベル': {'type': 'text', 'index': 44},
        '腹部肥満分析_内臓脂肪レベル_評価': {'type': 'text', 'index': 45},
        '部位別筋肉量_胴体': {'type': 'text', 'index': 46},
        '部位別筋肉量_胴体_評価': {'type': 'text', 'index': 47},
        '部位別筋肉量_左腕': {'type': 'text', 'index': 48},
        '部位別筋肉量_左腕_評価': {'type': 'text', 'index': 49},
        '部位別筋肉量_左脚': {'type': 'text', 'index': 50},
        '部位別筋肉量_左脚_評価': {'type': 'text', 'index': 51},
        '部位別筋肉量_右腕': {'type': 'text', 'index': 52},
        '部位別筋肉量_右腕_評価': {'type': 'text', 'index': 53},
        '部位別筋肉量_右脚': {'type': 'text', 'index': 54},
        '部位別筋肉量_右脚_評価': {'type': 'text', 'index': 55},
        '部位別脂肪量_胴体': {'type': 'text', 'index': 56},
        '部位別脂肪量_胴体_評価': {'type': 'text', 'index': 57},
        '部位別脂肪量_左腕': {'type': 'text', 'index': 58},
        '部位別脂肪量_左腕_評価': {'type': 'text', 'index': 59},
        '部位別脂肪量_左脚': {'type': 'text', 'index': 60},
        '部位別脂肪量_左脚_評価': {'type': 'text', 'index': 61},
        '部位別脂肪量_右腕': {'type': 'text', 'index': 62},
        '部位別脂肪量_右腕_評価': {'type': 'text', 'index': 63},
        '部位別脂肪量_右脚': {'type': 'text', 'index': 64},
        '部位別脂肪量_右脚_評価': {'type': 'text', 'index': 65}
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


# 各excelファイルの内容を読み込み
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(潮永さん）/【K】ACCUNIQ.xlsx', sheet_name='Sheet1', header=4)
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【K】ACCUNIQ.xlsx', sheet_name='Sheet1', header=4)
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【K】ACCUNIQ.xlsx', sheet_name='Sheet1', header=4)
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【K】ACCUNIQ.xlsx', sheet_name='Sheet1', header=4)
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【K】ACCUNIQ.xlsx', sheet_name='Sheet1', header=4)

# データフレーム結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])
wb = wb.rename(columns={'会員番号' : 'No'})

# 不要な値を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last') 

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)