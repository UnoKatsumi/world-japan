import pandas as pd

# excelファイルを読み込み
df = pd.read_excel('S:\個人作業用\渡邊\ワールドジャパン\classify_body.xlsx', sheet_name='有効データ_body', engine='openpyxl')
path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_body_custom.xlsx'
sheet = '有効データ_body'

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
    if row[5] == 'A' :
        df.loc[index, '血液型'] = 1
    elif row[5] == 'B' :
        df.loc[index, '血液型'] = 2
    elif row[5] == 'AB' :
        df.loc[index, '血液型'] = 3
    elif row[5] == 'O' :
        df.loc[index, '血液型'] = 4
    if row[6] == '既婚' :
        df.loc[index, '結婚'] = 1
    elif row[6] == '未婚' :
        df.loc[index, '結婚'] = 0
    if row[7] == 'HP' :
        df.loc[index, '知った理由'] = 1
    elif row[7] == 'WEB' :
        df.loc[index, '知った理由'] = 2
    elif row[7] == 'HPB' :
        df.loc[index, '知った理由'] = 3
    elif row[7] == 'チラシ' :
        df.loc[index, '知った理由'] = 4
    elif row[7] == '紹介' :
        df.loc[index, '知った理由'] = 5
    else :
        df.loc[index, '知った理由'] = 6

df = df.set_index('氏名')

# excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df.to_excel(writer, sheet_name = sheet)
