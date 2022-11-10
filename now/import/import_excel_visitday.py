import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'visitday_list',
    'column': {
        'No': {'type': 'text', 'index': 1},
        'お名前': {'type': 'text', 'index': 2},
        'コース名': {'type': 'text', 'index': 3},
        '契約日': {'type': 'text', 'index': 4},
        '有効期限': {'type': 'text', 'index': 5},
        '単価': {'type': 'text', 'index': 6},
        'コース_1回目': {'type': 'text', 'index': 7},
        'コース_2回目': {'type': 'text', 'index': 8},
        'コース_3回目': {'type': 'text', 'index': 9},
        'コース_4回目': {'type': 'text', 'index': 10},
        'コース_5回目': {'type': 'text', 'index': 11},
        'コース_6回目': {'type': 'text', 'index': 12},
        'コース_7回目': {'type': 'text', 'index': 13},
        'コース_8回目': {'type': 'text', 'index': 14},
        'コース_9回目': {'type': 'text', 'index': 15},
        'コース_10回目': {'type': 'text', 'index': 16},
        'コース_11回目': {'type': 'text', 'index': 17},
        'コース_12回目': {'type': 'text', 'index': 18},
        'コース_13回目': {'type': 'text', 'index': 19},
        'コース_14回目': {'type': 'text', 'index': 20},
        'コース_15回目': {'type': 'text', 'index': 21},
        'コース_16回目': {'type': 'text', 'index': 22},
        'コース_17回目': {'type': 'text', 'index': 23},
        'コース_18回目': {'type': 'text', 'index': 24},
        'コース_19回目': {'type': 'text', 'index': 25},
        'コース_20回目': {'type': 'text', 'index': 26},
        'コース_21回目': {'type': 'text', 'index': 27},
        'コース_22回目': {'type': 'text', 'index': 28},
        'コース_23回目': {'type': 'text', 'index': 29},
        'コース_24回目': {'type': 'text', 'index': 30},
        'コース_25回目': {'type': 'text', 'index': 31},
        'コース_26回目': {'type': 'text', 'index': 32},
        'コース_27回目': {'type': 'text', 'index': 33},
        'コース_28回目': {'type': 'text', 'index': 34},
        'コース_29回目': {'type': 'text', 'index': 35},
        'コース_30回目': {'type': 'text', 'index': 36},
        'コース_31回目': {'type': 'text', 'index': 37},
        'コース_32回目': {'type': 'text', 'index': 38},
        'コース_33回目': {'type': 'text', 'index': 39},
        'コース_34回目': {'type': 'text', 'index': 40},
        'コース_35回目': {'type': 'text', 'index': 41},
        'コース_36回目': {'type': 'text', 'index': 42},
        'コース_37回目': {'type': 'text', 'index': 43},
        'コース_38回目': {'type': 'text', 'index': 44},
        'コース_39回目': {'type': 'text', 'index': 45},
        'コース_40回目': {'type': 'text', 'index': 46},
        'コース_41回目': {'type': 'text', 'index': 47},
        'コース_42回目': {'type': 'text', 'index': 48},
        'コース_43回目': {'type': 'text', 'index': 49},
        'コース_44回目': {'type': 'text', 'index': 50},
        'コース_45回目': {'type': 'text', 'index': 51},
        'コース_46回目': {'type': 'text', 'index': 52},
        'コース_47回目': {'type': 'text', 'index': 53},
        'コース_48回目': {'type': 'text', 'index': 54},
        'コース_49回目': {'type': 'text', 'index': 55},
        'コース_50回目': {'type': 'text', 'index': 56},
        'コース_51回目': {'type': 'text', 'index': 57},
        'コース_52回目': {'type': 'text', 'index': 58},
        'コース_53回目': {'type': 'text', 'index': 59},
        'コース_54回目': {'type': 'text', 'index': 60},
        'コース_55回目': {'type': 'text', 'index': 61},
        'コース_56回目': {'type': 'text', 'index': 62},
        'コース_57回目': {'type': 'text', 'index': 63},
        'コース_58回目': {'type': 'text', 'index': 64},
        'コース_59回目': {'type': 'text', 'index': 65},
        'コース_60回目': {'type': 'text', 'index': 66},
        'コース_61回目': {'type': 'text', 'index': 67},
        'コース_62回目': {'type': 'text', 'index': 68},
        'コース_63回目': {'type': 'text', 'index': 69},
        'コース_64回目': {'type': 'text', 'index': 70},
        'コース_65回目': {'type': 'text', 'index': 71},
        'コース_66回目': {'type': 'text', 'index': 72},
        'コース_67回目': {'type': 'text', 'index': 73},
        'コース_68回目': {'type': 'text', 'index': 74},
        'コース_69回目': {'type': 'text', 'index': 75},
        'コース_70回目': {'type': 'text', 'index': 76},
        'コース_71回目': {'type': 'text', 'index': 77},
        'コース_72回目': {'type': 'text', 'index': 78},
        'コース_73回目': {'type': 'text', 'index': 79},
        'コース_74回目': {'type': 'text', 'index': 80},
        'コース_75回目': {'type': 'text', 'index': 81},
        'コース_76回目': {'type': 'text', 'index': 82},
        'コース_77回目': {'type': 'text', 'index': 83},
        'コース_78回目': {'type': 'text', 'index': 84},
        'コース_79回目': {'type': 'text', 'index': 85},
        'コース_80回目': {'type': 'text', 'index': 86},
        'コース_81回目': {'type': 'text', 'index': 87},
        'コース_82回目': {'type': 'text', 'index': 88},
        'コース_83回目': {'type': 'text', 'index': 89},
        'コース_84回目': {'type': 'text', 'index': 90},
        'コース_85回目': {'type': 'text', 'index': 91},
        'コース_86回目': {'type': 'text', 'index': 92},
        'コース_87回目': {'type': 'text', 'index': 93},
        'コース_88回目': {'type': 'text', 'index': 94},
        'コース_89回目': {'type': 'text', 'index': 95},
        'コース_90回目': {'type': 'text', 'index': 96},
        'コース_91回目': {'type': 'text', 'index': 97},
        'コース_92回目': {'type': 'text', 'index': 98},
        'コース_93回目': {'type': 'text', 'index': 99},
        'コース_94回目': {'type': 'text', 'index': 100},
        'コース_95回目': {'type': 'text', 'index': 101},
        'コース_96回目': {'type': 'text', 'index': 102},
        'コース_97回目': {'type': 'text', 'index': 103},
        'コース_98回目': {'type': 'text', 'index': 104},
        'コース_99回目': {'type': 'text', 'index': 105},
        'コース_100回目': {'type': 'text', 'index': 106},
        'コース_101回目': {'type': 'text', 'index': 107},
        'コース_102回目': {'type': 'text', 'index': 108},
        'コース_103回目': {'type': 'text', 'index': 109},
        'コース_104回目': {'type': 'text', 'index': 110},
        'コース_105回目': {'type': 'text', 'index': 111},
        'コース_106回目': {'type': 'text', 'index': 112},
        'コース_107回目': {'type': 'text', 'index': 113},
        'コース_108回目': {'type': 'text', 'index': 114},
        'コース_109回目': {'type': 'text', 'index': 115},
        'コース_110回目': {'type': 'text', 'index': 116},
        'コース_111回目': {'type': 'text', 'index': 117},
        'コース_112回目': {'type': 'text', 'index': 118},
        'コース_113回目': {'type': 'text', 'index': 119},
        'コース_114回目': {'type': 'text', 'index': 120},
        'コース_115回目': {'type': 'text', 'index': 121},
        'コース_116回目': {'type': 'text', 'index': 122},
        'コース_117回目': {'type': 'text', 'index': 123},
        'コース_118回目': {'type': 'text', 'index': 124}
    }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
    # conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_H.db')
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
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(潮永さん）/【H】来店日リスト.xlsx', sheet_name='Sheet1', header=1)
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【H】来店日リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_2.insert(loc = 106,column = 'コース_101回目', value = '')
wb_2.insert(loc = 107,column = 'コース_102回目', value = '')
wb_2.insert(loc = 108,column = 'コース_103回目', value = '')
wb_2.insert(loc = 109,column = 'コース_104回目', value = '')
wb_2.insert(loc = 110,column = 'コース_105回目', value = '')
wb_2.insert(loc = 111,column = 'コース_106回目', value = '')
wb_2.insert(loc = 112,column = 'コース_107回目', value = '')
wb_2.insert(loc = 113,column = 'コース_108回目', value = '')
wb_2.insert(loc = 114,column = 'コース_109回目', value = '')
wb_2.insert(loc = 115,column = 'コース_110回目', value = '')
wb_2.insert(loc = 116,column = 'コース_111回目', value = '')
wb_2.insert(loc = 117,column = 'コース_112回目', value = '')
wb_2.insert(loc = 118,column = 'コース_113回目', value = '')
wb_2.insert(loc = 119,column = 'コース_114回目', value = '')
wb_2.insert(loc = 120,column = 'コース_115回目', value = '')
wb_2.insert(loc = 121,column = 'コース_116回目', value = '')
wb_2.insert(loc = 122,column = 'コース_117回目', value = '')
wb_2.insert(loc = 123,column = 'コース_118回目', value = '')

# excelファイルの内容を読み込み
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【H】来店日リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_3.insert(loc = 106,column = 'コース_101回目', value = '')
wb_3.insert(loc = 107,column = 'コース_102回目', value = '')
wb_3.insert(loc = 108,column = 'コース_103回目', value = '')
wb_3.insert(loc = 109,column = 'コース_104回目', value = '')
wb_3.insert(loc = 110,column = 'コース_105回目', value = '')
wb_3.insert(loc = 111,column = 'コース_106回目', value = '')
wb_3.insert(loc = 112,column = 'コース_107回目', value = '')
wb_3.insert(loc = 113,column = 'コース_108回目', value = '')
wb_3.insert(loc = 114,column = 'コース_109回目', value = '')
wb_3.insert(loc = 115,column = 'コース_110回目', value = '')
wb_3.insert(loc = 116,column = 'コース_111回目', value = '')
wb_3.insert(loc = 117,column = 'コース_112回目', value = '')
wb_3.insert(loc = 118,column = 'コース_113回目', value = '')
wb_3.insert(loc = 119,column = 'コース_114回目', value = '')
wb_3.insert(loc = 120,column = 'コース_115回目', value = '')
wb_3.insert(loc = 121,column = 'コース_116回目', value = '')
wb_3.insert(loc = 122,column = 'コース_117回目', value = '')
wb_3.insert(loc = 123,column = 'コース_118回目', value = '')

# excelファイルの内容を読み込み
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【H】来店日リスト.xlsx', sheet_name='Sheet1', header=1)
# 不足している列を追加
wb_4.insert(loc = 56,column = 'コース_51回目', value = '')
wb_4.insert(loc = 57,column = 'コース_52回目', value = '')
wb_4.insert(loc = 58,column = 'コース_53回目', value = '')
wb_4.insert(loc = 59,column = 'コース_54回目', value = '')
wb_4.insert(loc = 60,column = 'コース_55回目', value = '')
wb_4.insert(loc = 61,column = 'コース_56回目', value = '')
wb_4.insert(loc = 62,column = 'コース_57回目', value = '')
wb_4.insert(loc = 63,column = 'コース_58回目', value = '')
wb_4.insert(loc = 64,column = 'コース_59回目', value = '')
wb_4.insert(loc = 65,column = 'コース_60回目', value = '')
wb_4.insert(loc = 66,column = 'コース_61回目', value = '')
wb_4.insert(loc = 67,column = 'コース_62回目', value = '')
wb_4.insert(loc = 68,column = 'コース_63回目', value = '')
wb_4.insert(loc = 69,column = 'コース_64回目', value = '')
wb_4.insert(loc = 70,column = 'コース_65回目', value = '')
wb_4.insert(loc = 71,column = 'コース_66回目', value = '')
wb_4.insert(loc = 72,column = 'コース_67回目', value = '')
wb_4.insert(loc = 73,column = 'コース_68回目', value = '')
wb_4.insert(loc = 74,column = 'コース_69回目', value = '')
wb_4.insert(loc = 75,column = 'コース_70回目', value = '')
wb_4.insert(loc = 76,column = 'コース_71回目', value = '')
wb_4.insert(loc = 77,column = 'コース_72回目', value = '')
wb_4.insert(loc = 78,column = 'コース_73回目', value = '')
wb_4.insert(loc = 79,column = 'コース_74回目', value = '')
wb_4.insert(loc = 80,column = 'コース_75回目', value = '')
wb_4.insert(loc = 81,column = 'コース_76回目', value = '')
wb_4.insert(loc = 82,column = 'コース_77回目', value = '')
wb_4.insert(loc = 83,column = 'コース_78回目', value = '')
wb_4.insert(loc = 84,column = 'コース_79回目', value = '')
wb_4.insert(loc = 85,column = 'コース_80回目', value = '')
wb_4.insert(loc = 86,column = 'コース_81回目', value = '')
wb_4.insert(loc = 87,column = 'コース_82回目', value = '')
wb_4.insert(loc = 88,column = 'コース_83回目', value = '')
wb_4.insert(loc = 89,column = 'コース_84回目', value = '')
wb_4.insert(loc = 90,column = 'コース_85回目', value = '')
wb_4.insert(loc = 91,column = 'コース_86回目', value = '')
wb_4.insert(loc = 92,column = 'コース_87回目', value = '')
wb_4.insert(loc = 93,column = 'コース_88回目', value = '')
wb_4.insert(loc = 94,column = 'コース_89回目', value = '')
wb_4.insert(loc = 95,column = 'コース_90回目', value = '')
wb_4.insert(loc = 96,column = 'コース_91回目', value = '')
wb_4.insert(loc = 97,column = 'コース_92回目', value = '')
wb_4.insert(loc = 98,column = 'コース_93回目', value = '')
wb_4.insert(loc = 99,column = 'コース_94回目', value = '')
wb_4.insert(loc = 100,column = 'コース_95回目', value = '')
wb_4.insert(loc = 101,column = 'コース_96回目', value = '')
wb_4.insert(loc = 102,column = 'コース_97回目', value = '')
wb_4.insert(loc = 103,column = 'コース_98回目', value = '')
wb_4.insert(loc = 104,column = 'コース_99回目', value = '')
wb_4.insert(loc = 105,column = 'コース_100回目', value = '')
wb_4.insert(loc = 106,column = 'コース_101回目', value = '')
wb_4.insert(loc = 107,column = 'コース_102回目', value = '')
wb_4.insert(loc = 108,column = 'コース_103回目', value = '')
wb_4.insert(loc = 109,column = 'コース_104回目', value = '')
wb_4.insert(loc = 110,column = 'コース_105回目', value = '')
wb_4.insert(loc = 111,column = 'コース_106回目', value = '')
wb_4.insert(loc = 112,column = 'コース_107回目', value = '')
wb_4.insert(loc = 113,column = 'コース_108回目', value = '')
wb_4.insert(loc = 114,column = 'コース_109回目', value = '')
wb_4.insert(loc = 115,column = 'コース_110回目', value = '')
wb_4.insert(loc = 116,column = 'コース_111回目', value = '')
wb_4.insert(loc = 117,column = 'コース_112回目', value = '')
wb_4.insert(loc = 118,column = 'コース_113回目', value = '')
wb_4.insert(loc = 119,column = 'コース_114回目', value = '')
wb_4.insert(loc = 120,column = 'コース_115回目', value = '')
wb_4.insert(loc = 121,column = 'コース_116回目', value = '')
wb_4.insert(loc = 122,column = 'コース_117回目', value = '')
wb_4.insert(loc = 123,column = 'コース_118回目', value = '')

# excelファイルの内容を読み込み
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【H】来店日リスト.xlsx', sheet_name='Sheet1', header=1)

#データフレームを結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])

wb = wb.rename(columns={'NO.' : 'No'})

# 不要な値を含む行を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last')

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)