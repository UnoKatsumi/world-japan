import pandas as pd
import numpy as np

# excelファイルを読み込み
df = pd.read_excel('S:\個人作業用\渡邊\ワールドジャパン\classify_all.xlsx', sheet_name='有効データ', engine='openpyxl')
path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_all_custom.xlsx'
sheet = '有効データ'

# 特定のカラムにおける文字列を数値に変換
for index, row in df.iterrows():
    if row[4] == '会社員' :
        df.loc[index, '職業'] = 1
    elif row[4] == '自営業' :
        df.loc[index, '職業'] = 2
    elif row[4] == 'パート・アルバイト' :
        df.loc[index, '職業'] = 3
    elif row[4] == '主婦' :
        df.loc[index, '職業'] = 4
    elif row[4] == '学生' :
        df.loc[index, '職業'] = 5
    elif row[4] == '無職' :
        df.loc[index, '職業'] = 6
    else :
        df.loc[index, '職業'] = 7
    if row[5] == '既婚' :
        df.loc[index, '結婚'] = 1
    elif row[5] == '未婚' :
        df.loc[index, '結婚'] = 0
    if row[6] == 'HP' :
        df.loc[index, '知った理由'] = 1
    elif row[6] == 'WEB' :
        df.loc[index, '知った理由'] = 2
    elif row[6] == 'HPB' :
        df.loc[index, '知った理由'] = 3
    elif row[6] == 'チラシ' :
        df.loc[index, '知った理由'] = 4
    elif row[6] == '紹介' :
        df.loc[index, '知った理由'] = 5
    else :
        df.loc[index, '知った理由'] = 6

df = df.set_index('氏名')

# excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df.to_excel(writer, sheet_name = sheet)
