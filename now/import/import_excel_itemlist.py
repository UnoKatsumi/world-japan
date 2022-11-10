import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'item_list',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        '日付_1行目': {'type': 'text', 'index': 3},
        '商品名_1行目': {'type': 'text', 'index': 4},
        '金額_1行目': {'type': 'text', 'index': 5},
        '個数_1行目': {'type': 'text', 'index': 6},
        '使用方法_1行目': {'type': 'text', 'index': 7},
        '日付_2行目': {'type': 'text', 'index': 8},
        '商品名_2行目': {'type': 'text', 'index': 9},
        '金額_2行目': {'type': 'text', 'index': 10},
        '個数_2行目': {'type': 'text', 'index': 11},
        '使用方法_2行目': {'type': 'text', 'index': 12},
        '日付_3行目': {'type': 'text', 'index': 13},
        '商品名_3行目': {'type': 'text', 'index': 14},
        '金額_3行目': {'type': 'text', 'index': 15},
        '個数_3行目': {'type': 'text', 'index': 16},
        '使用方法_3行目': {'type': 'text', 'index': 17},
        '日付_4行目': {'type': 'text', 'index': 18},
        '商品名_4行目': {'type': 'text', 'index': 19},
        '金額_4行目': {'type': 'text', 'index': 20},
        '個数_4行目': {'type': 'text', 'index': 21},
        '使用方法_4行目': {'type': 'text', 'index': 22},
        '日付_5行目': {'type': 'text', 'index': 23},
        '商品名_5行目': {'type': 'text', 'index': 24},
        '金額_5行目': {'type': 'text', 'index': 25},
        '個数_5行目': {'type': 'text', 'index': 26},
        '使用方法_5行目': {'type': 'text', 'index': 27},
        '日付_6行目': {'type': 'text', 'index': 28},
        '商品名_6行目': {'type': 'text', 'index': 29},
        '金額_6行目': {'type': 'text', 'index': 30},
        '個数_6行目': {'type': 'text', 'index': 31},
        '使用方法_6行目': {'type': 'text', 'index': 32},
        '日付_7行目': {'type': 'text', 'index': 33},
        '商品名_7行目': {'type': 'text', 'index': 34},
        '金額_7行目': {'type': 'text', 'index': 35},
        '個数_7行目': {'type': 'text', 'index': 36},
        '使用方法_7行目': {'type': 'text', 'index': 37},
        '日付_8行目': {'type': 'text', 'index': 38},
        '商品名_8行目': {'type': 'text', 'index': 39},
        '金額_8行目': {'type': 'text', 'index': 40},
        '個数_8行目': {'type': 'text', 'index': 41},
        '使用方法_8行目': {'type': 'text', 'index': 42},
        '日付_9行目': {'type': 'text', 'index': 43},
        '商品名_9行目': {'type': 'text', 'index': 44},
        '金額_9行目': {'type': 'text', 'index': 45},
        '個数_9行目': {'type': 'text', 'index': 46},
        '使用方法_9行目': {'type': 'text', 'index': 47},
        '日付_10行目': {'type': 'text', 'index': 48},
        '商品名_10行目': {'type': 'text', 'index': 49},
        '金額_10行目': {'type': 'text', 'index': 50},
        '個数_10行目': {'type': 'text', 'index': 51},
        '使用方法_10行目': {'type': 'text', 'index': 52},
        '日付_2枚目_1行目': {'type': 'text', 'index': 53},
        '商品名_2枚目_1行目': {'type': 'text', 'index': 54},
        '金額_2枚目_1行目': {'type': 'text', 'index': 55},
        '個数_2枚目_1行目': {'type': 'text', 'index': 56},
        '使用方法_2枚目_1行目': {'type': 'text', 'index': 57},
        '日付_2枚目_2行目': {'type': 'text', 'index': 58},
        '商品名_2枚目_2行目': {'type': 'text', 'index': 59},
        '金額_2枚目_2行目': {'type': 'text', 'index': 60},
        '個数_2枚目_2行目': {'type': 'text', 'index': 61},
        '使用方法_2枚目_2行目': {'type': 'text', 'index': 62},
        '日付_2枚目_3行目': {'type': 'text', 'index': 63},
        '商品名_2枚目_3行目': {'type': 'text', 'index': 64},
        '金額_2枚目_3行目': {'type': 'text', 'index': 65},
        '個数_2枚目_3行目': {'type': 'text', 'index': 66},
        '使用方法_2枚目_3行目': {'type': 'text', 'index': 67},
        '日付_2枚目_4行目': {'type': 'text', 'index': 68},
        '商品名_2枚目_4行目': {'type': 'text', 'index': 69},
        '金額_2枚目_4行目': {'type': 'text', 'index': 70},
        '個数_2枚目_4行目': {'type': 'text', 'index': 71},
        '使用方法_2枚目_4行目': {'type': 'text', 'index': 72},
        '日付_2枚目_5行目': {'type': 'text', 'index': 73},
        '商品名_2枚目_5行目': {'type': 'text', 'index': 74},
        '金額_2枚目_5行目': {'type': 'text', 'index': 75},
        '個数_2枚目_5行目': {'type': 'text', 'index': 76},
        '使用方法_2枚目_5行目': {'type': 'text', 'index': 77},
        '日付_2枚目_6行目': {'type': 'text', 'index': 78},
        '商品名_2枚目_6行目': {'type': 'text', 'index': 79},
        '金額_2枚目_6行目': {'type': 'text', 'index': 80},
        '個数_2枚目_6行目': {'type': 'text', 'index': 81},
        '使用方法_2枚目_6行目': {'type': 'text', 'index': 82},
        '日付_2枚目_7行目': {'type': 'text', 'index': 83},
        '商品名_2枚目_7行目': {'type': 'text', 'index': 84},
        '金額_2枚目_7行目': {'type': 'text', 'index': 85},
        '個数_2枚目_7行目': {'type': 'text', 'index': 86},
        '使用方法_2枚目_7行目': {'type': 'text', 'index': 87},
        '日付_2枚目_8行目': {'type': 'text', 'index': 88},
        '商品名_2枚目_8行目': {'type': 'text', 'index': 89},
        '金額_2枚目_8行目': {'type': 'text', 'index': 90},
        '個数_2枚目_8行目': {'type': 'text', 'index': 91},
        '使用方法_2枚目_8行目': {'type': 'text', 'index': 92},
        '日付_2枚目_9行目': {'type': 'text', 'index': 93},
        '商品名_2枚目_9行目': {'type': 'text', 'index': 94},
        '金額_2枚目_9行目': {'type': 'text', 'index': 95},
        '個数_2枚目_9行目': {'type': 'text', 'index': 96},
        '使用方法_2枚目_9行目': {'type': 'text', 'index': 97},
        '日付_2枚目_10行目': {'type': 'text', 'index': 98},
        '商品名_2枚目_10行目': {'type': 'text', 'index': 99},
        '金額_2枚目_10行目': {'type': 'text', 'index': 100},
        '個数_2枚目_10行目': {'type': 'text', 'index': 101},
        '使用方法_2枚目_10行目': {'type': 'text', 'index': 102},
        '日付_3枚目_1行目': {'type': 'text', 'index': 103},
        '商品名_3枚目_1行目': {'type': 'text', 'index': 104},
        '金額_3枚目_1行目': {'type': 'text', 'index': 105},
        '個数_3枚目_1行目': {'type': 'text', 'index': 106},
        '使用方法_3枚目_1行目': {'type': 'text', 'index': 107},
        '日付_3枚目_2行目': {'type': 'text', 'index': 108},
        '商品名_3枚目_2行目': {'type': 'text', 'index': 109},
        '金額_3枚目_2行目': {'type': 'text', 'index': 110},
        '個数_3枚目_2行目': {'type': 'text', 'index': 111},
        '使用方法_3枚目_2行目': {'type': 'text', 'index': 112},
        '日付_3枚目_3行目': {'type': 'text', 'index': 113},
        '商品名_3枚目_3行目': {'type': 'text', 'index': 114},
        '金額_3枚目_3行目': {'type': 'text', 'index': 115},
        '個数_3枚目_3行目': {'type': 'text', 'index': 116},
        '使用方法_3枚目_3行目': {'type': 'text', 'index': 117},
        '日付_3枚目_4行目': {'type': 'text', 'index': 118},
        '商品名_3枚目_4行目': {'type': 'text', 'index': 119},
        '金額_3枚目_4行目': {'type': 'text', 'index': 120},
        '個数_3枚目_4行目': {'type': 'text', 'index': 121},
        '使用方法_3枚目_4行目': {'type': 'text', 'index': 122},
        '日付_3枚目_5行目': {'type': 'text', 'index': 123},
        '商品名_3枚目_5行目': {'type': 'text', 'index': 124},
        '金額_3枚目_5行目': {'type': 'text', 'index': 125},
        '個数_3枚目_5行目': {'type': 'text', 'index': 126},
        '使用方法_3枚目_5行目': {'type': 'text', 'index': 127},
        '日付_3枚目_6行目': {'type': 'text', 'index': 128},
        '商品名_3枚目_6行目': {'type': 'text', 'index': 129},
        '金額_3枚目_6行目': {'type': 'text', 'index': 130},
        '個数_3枚目_6行目': {'type': 'text', 'index': 131},
        '使用方法_3枚目_6行目': {'type': 'text', 'index': 132},
        '日付_3枚目_7行目': {'type': 'text', 'index': 133},
        '商品名_3枚目_7行目': {'type': 'text', 'index': 134},
        '金額_3枚目_7行目': {'type': 'text', 'index': 135},
        '個数_3枚目_7行目': {'type': 'text', 'index': 136},
        '使用方法_3枚目_7行目': {'type': 'text', 'index': 137},
        '日付_3枚目_8行目': {'type': 'text', 'index': 138},
        '商品名_3枚目_8行目': {'type': 'text', 'index': 139},
        '金額_3枚目_8行目': {'type': 'text', 'index': 140},
        '個数_3枚目_8行目': {'type': 'text', 'index': 141},
        '使用方法_3枚目_8行目': {'type': 'text', 'index': 142},
        '日付_3枚目_9行目': {'type': 'text', 'index': 143},
        '商品名_3枚目_9行目': {'type': 'text', 'index': 144},
        '金額_3枚目_9行目': {'type': 'text', 'index': 145},
        '個数_3枚目_9行目': {'type': 'text', 'index': 146},
        '使用方法_3枚目_9行目': {'type': 'text', 'index': 147},
        '日付_3枚目_10行目': {'type': 'text', 'index': 148},
        '商品名_3枚目_10行目': {'type': 'text', 'index': 149},
        '金額_3枚目_10行目': {'type': 'text', 'index': 150},
        '個数_3枚目_10行目': {'type': 'text', 'index': 151},
        '使用方法_3枚目_10行目': {'type': 'text', 'index': 152},
        '日付_4枚目_1行目': {'type': 'text', 'index': 153},
        '商品名_4枚目_1行目': {'type': 'text', 'index': 154},
        '金額_4枚目_1行目': {'type': 'text', 'index': 155},
        '個数_4枚目_1行目': {'type': 'text', 'index': 156},
        '使用方法_4枚目_1行目': {'type': 'text', 'index': 157},
        '日付_4枚目_2行目': {'type': 'text', 'index': 158},
        '商品名_4枚目_2行目': {'type': 'text', 'index': 159},
        '金額_4枚目_2行目': {'type': 'text', 'index': 160},
        '個数_4枚目_2行目': {'type': 'text', 'index': 161},
        '使用方法_4枚目_2行目': {'type': 'text', 'index': 162},
        '日付_4枚目_3行目': {'type': 'text', 'index': 163},
        '商品名_4枚目_3行目': {'type': 'text', 'index': 164},
        '金額_4枚目_3行目': {'type': 'text', 'index': 165},
        '個数_4枚目_3行目': {'type': 'text', 'index': 166},
        '使用方法_4枚目_3行目': {'type': 'text', 'index': 167},
        '日付_4枚目_4行目': {'type': 'text', 'index': 168},
        '商品名_4枚目_4行目': {'type': 'text', 'index': 169},
        '金額_4枚目_4行目': {'type': 'text', 'index': 170},
        '個数_4枚目_4行目': {'type': 'text', 'index': 171},
        '使用方法_4枚目_4行目': {'type': 'text', 'index': 172},
    }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
    # conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_L.db')
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
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(潮永さん）/【L】商品管理リスト.xlsx', sheet_name='Sheet1', header=1)
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【L】商品管理リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_2.insert(loc = 152,column = '日付_4枚目_1行目', value = None)
wb_2.insert(loc = 153,column = '商品名_4枚目_1行目', value = None)
wb_2.insert(loc = 154,column = '金額_4枚目_1行目', value = None)
wb_2.insert(loc = 155,column = '個数_4枚目_1行目', value = None)
wb_2.insert(loc = 156,column = '使用方法_4枚目_1行目', value = None)
wb_2.insert(loc = 157,column = '日付_4枚目_2行目', value = None)
wb_2.insert(loc = 158,column = '商品名_4枚目_2行目', value = None)
wb_2.insert(loc = 159,column = '金額_4枚目_2行目', value = None)
wb_2.insert(loc = 160,column = '個数_4枚目_2行目', value = None)
wb_2.insert(loc = 161,column = '使用方法_4枚目_2行目', value = None)
wb_2.insert(loc = 162,column = '日付_4枚目_3行目', value = None)
wb_2.insert(loc = 163,column = '商品名_4枚目_3行目', value = None)
wb_2.insert(loc = 164,column = '金額_4枚目_3行目', value = None)
wb_2.insert(loc = 165,column = '個数_4枚目_3行目', value = None)
wb_2.insert(loc = 166,column = '使用方法_4枚目_3行目', value = None)
wb_2.insert(loc = 167,column = '日付_4枚目_4行目', value = None)
wb_2.insert(loc = 168,column = '商品名_4枚目_4行目', value = None)
wb_2.insert(loc = 169,column = '金額_4枚目_4行目', value = None)
wb_2.insert(loc = 170,column = '個数_4枚目_4行目', value = None)
wb_2.insert(loc = 171,column = '使用方法_4枚目_4行目', value = None)

# excelファイルの内容を読み込み
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【L】商品管理リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_3.insert(loc = 127,column = '日付_3枚目_6行目', value = None)
wb_3.insert(loc = 128,column = '商品名_3枚目_6行目', value = None)
wb_3.insert(loc = 129,column = '金額__3枚目_6行目', value = None)
wb_3.insert(loc = 130,column = '個数_3枚目_6行目', value = None)
wb_3.insert(loc = 131,column = '使用方法_3枚目_6行目', value = None)
wb_3.insert(loc = 132,column = '日付_3枚目_7行目', value = None)
wb_3.insert(loc = 133,column = '商品名_3枚目_7行目', value = None)
wb_3.insert(loc = 134,column = '金額_3枚目_7行目', value = None)
wb_3.insert(loc = 135,column = '個数_3枚目_7行目', value = None)
wb_3.insert(loc = 136,column = '使用方法_3枚目_7行目', value = None)
wb_3.insert(loc = 137,column = '日付_3枚目_8行目', value = None)
wb_3.insert(loc = 138,column = '商品名_3枚目_8行目', value = None)
wb_3.insert(loc = 139,column = '金額_3枚目_8行目', value = None)
wb_3.insert(loc = 140,column = '個数_3枚目_8行目', value = None)
wb_3.insert(loc = 141,column = '使用方法_3枚目_8行目', value = None)
wb_3.insert(loc = 142,column = '日付_3枚目_9行目', value = None)
wb_3.insert(loc = 143,column = '商品名_3枚目_9行目', value = None)
wb_3.insert(loc = 144,column = '金額_3枚目_9行目', value = None)
wb_3.insert(loc = 145,column = '個数_3枚目_9行目', value = None)
wb_3.insert(loc = 146,column = '使用方法_3枚目_9行目', value = None)
wb_3.insert(loc = 147,column = '日付_3枚目_10行目', value = None)
wb_3.insert(loc = 148,column = '商品名_3枚目_10行目', value = None)
wb_3.insert(loc = 149,column = '金額_3枚目_10行目', value = None)
wb_3.insert(loc = 150,column = '個数_3枚目_10行目', value = None)
wb_3.insert(loc = 151,column = '使用方法_3枚目_10行目', value = None)
wb_3.insert(loc = 152,column = '日付_4枚目_1行目', value = None)
wb_3.insert(loc = 153,column = '商品名_4枚目_1行目', value = None)
wb_3.insert(loc = 154,column = '金額_4枚目_1行目', value = None)
wb_3.insert(loc = 155,column = '個数_4枚目_1行目', value = None)
wb_3.insert(loc = 156,column = '使用方法_4枚目_1行目', value = None)
wb_3.insert(loc = 157,column = '日付_4枚目_2行目', value = None)
wb_3.insert(loc = 158,column = '商品名_4枚目_2行目', value = None)
wb_3.insert(loc = 159,column = '金額_4枚目_2行目', value = None)
wb_3.insert(loc = 160,column = '個数_4枚目_2行目', value = None)
wb_3.insert(loc = 161,column = '使用方法_4枚目_2行目', value = None)
wb_3.insert(loc = 162,column = '日付_4枚目_3行目', value = None)
wb_3.insert(loc = 163,column = '商品名_4枚目_3行目', value = None)
wb_3.insert(loc = 164,column = '金額_4枚目_3行目', value = None)
wb_3.insert(loc = 165,column = '個数_4枚目_3行目', value = None)
wb_3.insert(loc = 166,column = '使用方法_4枚目_3行目', value = None)
wb_3.insert(loc = 167,column = '日付_4枚目_4行目', value = None)
wb_3.insert(loc = 168,column = '商品名_4枚目_4行目', value = None)
wb_3.insert(loc = 169,column = '金額_4枚目_4行目', value = None)
wb_3.insert(loc = 170,column = '個数_4枚目_4行目', value = None)
wb_3.insert(loc = 171,column = '使用方法_4枚目_4行目', value = None)

# excelファイルの内容を読み込み
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【L】商品管理リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_4.insert(loc = 107,column = '日付_3枚目_2行目', value = None)
wb_4.insert(loc = 108,column = '商品名_3枚目_2行目', value = None)
wb_4.insert(loc = 109,column = '金額_3枚目_2行目', value = None)
wb_4.insert(loc = 110,column = '個数_3枚目_2行目', value = None)
wb_4.insert(loc = 111,column = '使用方法_3枚目_2行目', value = None)
wb_4.insert(loc = 112,column = '日付_3枚目_3行目', value = None)
wb_4.insert(loc = 113,column = '商品名_3枚目_3行目', value = None)
wb_4.insert(loc = 114,column = '金額_3枚目_3行目', value = None)
wb_4.insert(loc = 115,column = '個数_3枚目_3行目', value = None)
wb_4.insert(loc = 116,column = '使用方法_3枚目_3行目', value = None)
wb_4.insert(loc = 117,column = '日付_3枚目_4行目', value = None)
wb_4.insert(loc = 118,column = '商品名_3枚目_4行目', value = None)
wb_4.insert(loc = 119,column = '金額_3枚目_4行目', value = None)
wb_4.insert(loc = 120,column = '個数_3枚目_4行目', value = None)
wb_4.insert(loc = 121,column = '使用方法_3枚目_4行目', value = None)
wb_4.insert(loc = 122,column = '日付_3枚目_5行目', value = None)
wb_4.insert(loc = 123,column = '商品名_3枚目_5行目', value = None)
wb_4.insert(loc = 124,column = '金額_3枚目_5行目', value = None)
wb_4.insert(loc = 125,column = '個数_3枚目_5行目', value = None)
wb_4.insert(loc = 126,column = '使用方法_3枚目_5行目', value = None)
wb_4.insert(loc = 127,column = '日付_3枚目_6行目', value = None)
wb_4.insert(loc = 128,column = '商品名_3枚目_6行目', value = None)
wb_4.insert(loc = 129,column = '金額__3枚目_6行目', value = None)
wb_4.insert(loc = 130,column = '個数_3枚目_6行目', value = None)
wb_4.insert(loc = 131,column = '使用方法_3枚目_6行目', value = None)
wb_4.insert(loc = 132,column = '日付_3枚目_7行目', value = None)
wb_4.insert(loc = 133,column = '商品名_3枚目_7行目', value = None)
wb_4.insert(loc = 134,column = '金額_3枚目_7行目', value = None)
wb_4.insert(loc = 135,column = '個数_3枚目_7行目', value = None)
wb_4.insert(loc = 136,column = '使用方法_3枚目_7行目', value = None)
wb_4.insert(loc = 137,column = '日付_3枚目_8行目', value = None)
wb_4.insert(loc = 138,column = '商品名_3枚目_8行目', value = None)
wb_4.insert(loc = 139,column = '金額_3枚目_8行目', value = None)
wb_4.insert(loc = 140,column = '個数_3枚目_8行目', value = None)
wb_4.insert(loc = 141,column = '使用方法_3枚目_8行目', value = None)
wb_4.insert(loc = 142,column = '日付_3枚目_9行目', value = None)
wb_4.insert(loc = 143,column = '商品名_3枚目_9行目', value = None)
wb_4.insert(loc = 144,column = '金額_3枚目_9行目', value = None)
wb_4.insert(loc = 145,column = '個数_3枚目_9行目', value = None)
wb_4.insert(loc = 146,column = '使用方法_3枚目_9行目', value = None)
wb_4.insert(loc = 147,column = '日付_3枚目_10行目', value = None)
wb_4.insert(loc = 148,column = '商品名_3枚目_10行目', value = None)
wb_4.insert(loc = 149,column = '金額_3枚目_10行目', value = None)
wb_4.insert(loc = 150,column = '個数_3枚目_10行目', value = None)
wb_4.insert(loc = 151,column = '使用方法_3枚目_10行目', value = None)
wb_4.insert(loc = 152,column = '日付_4枚目_1行目', value = None)
wb_4.insert(loc = 153,column = '商品名_4枚目_1行目', value = None)
wb_4.insert(loc = 154,column = '金額_4枚目_1行目', value = None)
wb_4.insert(loc = 155,column = '個数_4枚目_1行目', value = None)
wb_4.insert(loc = 156,column = '使用方法_4枚目_1行目', value = None)
wb_4.insert(loc = 157,column = '日付_4枚目_2行目', value = None)
wb_4.insert(loc = 158,column = '商品名_4枚目_2行目', value = None)
wb_4.insert(loc = 159,column = '金額_4枚目_2行目', value = None)
wb_4.insert(loc = 160,column = '個数_4枚目_2行目', value = None)
wb_4.insert(loc = 161,column = '使用方法_4枚目_2行目', value = None)
wb_4.insert(loc = 162,column = '日付_4枚目_3行目', value = None)
wb_4.insert(loc = 163,column = '商品名_4枚目_3行目', value = None)
wb_4.insert(loc = 164,column = '金額_4枚目_3行目', value = None)
wb_4.insert(loc = 165,column = '個数_4枚目_3行目', value = None)
wb_4.insert(loc = 166,column = '使用方法_4枚目_3行目', value = None)
wb_4.insert(loc = 167,column = '日付_4枚目_4行目', value = None)
wb_4.insert(loc = 168,column = '商品名_4枚目_4行目', value = None)
wb_4.insert(loc = 169,column = '金額_4枚目_4行目', value = None)
wb_4.insert(loc = 170,column = '個数_4枚目_4行目', value = None)
wb_4.insert(loc = 171,column = '使用方法_4枚目_4行目', value = None)

# excelファイルの内容を読み込み
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【L】商品管理リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_5.insert(loc = 137,column = '日付_3枚目_8行目', value = None)
wb_5.insert(loc = 138,column = '商品名_3枚目_8行目', value = None)
wb_5.insert(loc = 139,column = '金額_3枚目_8行目', value = None)
wb_5.insert(loc = 140,column = '個数_3枚目_8行目', value = None)
wb_5.insert(loc = 141,column = '使用方法_3枚目_8行目', value = None)
wb_5.insert(loc = 142,column = '日付_3枚目_9行目', value = None)
wb_5.insert(loc = 143,column = '商品名_3枚目_9行目', value = None)
wb_5.insert(loc = 144,column = '金額_3枚目_9行目', value = None)
wb_5.insert(loc = 145,column = '個数_3枚目_9行目', value = None)
wb_5.insert(loc = 146,column = '使用方法_3枚目_9行目', value = None)
wb_5.insert(loc = 147,column = '日付_3枚目_10行目', value = None)
wb_5.insert(loc = 148,column = '商品名_3枚目_10行目', value = None)
wb_5.insert(loc = 149,column = '金額_3枚目_10行目', value = None)
wb_5.insert(loc = 150,column = '個数_3枚目_10行目', value = None)
wb_5.insert(loc = 151,column = '使用方法_3枚目_10行目', value = None)
wb_5.insert(loc = 152,column = '日付_4枚目_1行目', value = None)
wb_5.insert(loc = 153,column = '商品名_4枚目_1行目', value = None)
wb_5.insert(loc = 154,column = '金額_4枚目_1行目', value = None)
wb_5.insert(loc = 155,column = '個数_4枚目_1行目', value = None)
wb_5.insert(loc = 156,column = '使用方法_4枚目_1行目', value = None)
wb_5.insert(loc = 157,column = '日付_4枚目_2行目', value = None)
wb_5.insert(loc = 158,column = '商品名_4枚目_2行目', value = None)
wb_5.insert(loc = 159,column = '金額_4枚目_2行目', value = None)
wb_5.insert(loc = 160,column = '個数_4枚目_2行目', value = None)
wb_5.insert(loc = 161,column = '使用方法_4枚目_2行目', value = None)
wb_5.insert(loc = 162,column = '日付_4枚目_3行目', value = None)
wb_5.insert(loc = 163,column = '商品名_4枚目_3行目', value = None)
wb_5.insert(loc = 164,column = '金額_4枚目_3行目', value = None)
wb_5.insert(loc = 165,column = '個数_4枚目_3行目', value = None)
wb_5.insert(loc = 166,column = '使用方法_4枚目_3行目', value = None)
wb_5.insert(loc = 167,column = '日付_4枚目_4行目', value = None)
wb_5.insert(loc = 168,column = '商品名_4枚目_4行目', value = None)
wb_5.insert(loc = 169,column = '金額_4枚目_4行目', value = None)
wb_5.insert(loc = 170,column = '個数_4枚目_4行目', value = None)
wb_5.insert(loc = 171,column = '使用方法_4枚目_4行目', value = None)

#データフレームを結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])

wb = wb.rename(columns={'NO.' : 'No'})

# 不要な値を含む行を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last')

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)