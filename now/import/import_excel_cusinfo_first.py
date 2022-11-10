import pandas as pd
import sqlite3

# アップロードに用いるdbmap
excel_db_map = {
    'table_name': 'customer_infomation_firsttime',
    'column': {
        'No': {'type': 'text', 'index': 1},
        '入会日_西暦': {'type': 'text', 'index': 2},
        'お名前': {'type': 'text', 'index': 3},
        'ご来店動機': {'type': 'text', 'index': 4},
        '通う理由': {'type': 'text', 'index': 5},
        '職業': {'type': 'text', 'index': 6},
        '性格': {'type': 'text', 'index': 7},
        '展望': {'type': 'text', 'index': 8},
        '来店ペース': {'type': 'text', 'index': 9},
        '施術箇所': {'type': 'text', 'index': 10},
        '商品': {'type': 'text', 'index': 11},
        '次回フォロー': {'type': 'text', 'index': 12},
        '定期購入': {'type': 'text', 'index': 13},
        '家族構成': {'type': 'text', 'index': 14},
        '彼氏': {'type': 'text', 'index': 15},
        '趣味': {'type': 'text', 'index': 16},
        '会社名': {'type': 'text', 'index': 17},
        '平均年収': {'type': 'int', 'index': 18},
        '勤務形態': {'type': 'text', 'index': 19},
        '休日の過ごし方': {'type': 'text', 'index': 20},
        'ローン': {'type': 'text', 'index': 21},
        '不労所得': {'type': 'text', 'index': 22},
        '支払い方法': {'type': 'text', 'index': 23},
        'DNA検査結果': {'type': 'text', 'index': 24}
    }
}


# データフレームをデータベースへアップロード
def db_insert(db, db_map):
    # conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_F.db')
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
wb_1 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(潮永さん）/【F】顧客情報(初回) .xlsx', sheet_name='Sheet1', header=1)
wb_2 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(渡部さん)/【F】顧客情報(初回) .xlsx', sheet_name='Sheet1', header=1)
wb_3 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/2フォーマット(柏葉さん)/【F】顧客情報(初回) .xlsx', sheet_name='Sheet1', header=1)
wb_4 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/慶野さん/【F】顧客情報(初回) .xlsx', sheet_name='Sheet1', header=1)
wb_5 = pd.read_excel('S:/個人作業用/渡邊/ワールドジャパン/学習用データ/サロンカルテデータ_相馬さん/佐藤さん/【F】顧客情報(初回) .xlsx', sheet_name='Sheet1', header=1)

# データフレームを結合
wb = pd.concat([wb_1, wb_2, wb_3, wb_4, wb_5])

wb = wb.rename(columns={'NO.' : 'No'})

# 不要な値を含む行を削除
wb = wb.dropna(subset=['No'])
wb = wb.drop_duplicates(subset='No', keep='last')

# 関数db_insertの呼び出し
db_insert(wb, excel_db_map)