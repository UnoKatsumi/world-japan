# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

# ハッシュチェインによる警告を無視
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_all.xlsx'
sheet1 = '有効データ'
sheet2 = 'エラーデータ(body)'
sheet3 = 'エラーデータ(bust)'
sheet4 = 'エラーデータ(facial)'
sheet5 = 'エラーデータ(hair removal)'

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   body   #############################
#body_4のデータを抽出
df_body_4 = pd.read_sql_query('select * from questionnaire_body_4', conn)
df_bd4 = df_body_4.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚', '知った理由']]

# ############################   bust   #############################
#bust_4のデータを抽出
df_bust_4 = pd.read_sql_query('select * from questionnaire_bust_4', conn)
df_bs4 = df_bust_4.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚', '知った理由']]

# ############################   facial   #############################
#facial_4のデータを抽出
df_facial_4 = pd.read_sql_query('select * from questionnaire_facial_4', conn)
df_fa4 = df_facial_4.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚', '知った理由']]

# ############################   hair removal   #############################
#hairremoval_4のデータを抽出
df_hairremoval_4 = pd.read_sql_query('select * from questionnaire_hairremoval_4', conn)
df_hr4 = df_hairremoval_4.loc[:,['氏名', '来店年月日', '生年月日', '職業', '既婚・未婚', '知った理由']]
#hairremoval_4の'年齢'が欠損しているため年齢を計算して設定
list_hr4_age = []
list_hr4_age_err = []
year_4 = []
for row in df_hr4.iterrows():
	if type(row[1][1]) == str and type(row[1][2]) == str :
		if ':' in row[1][1] and ':' in row[1][2] :
			date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
			date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		year_4.append(abs(old))
	else :
		year_4.append('')
df_hr4.insert(loc = 3, column = '年齢', value = year_4)
#hairremoval_4の'既婚・未婚'カラムの名前を'結婚'へ変更
df_hr4 = df_hr4.rename(columns ={'既婚・未婚' : '結婚'})

###各dfにエラー値を保持するカラムを追加
df_bd4.insert(loc = 7, column = 'エラー', value = '')
df_bs4.insert(loc = 7, column = 'エラー', value = '')
df_fa4.insert(loc = 7, column = 'エラー', value = '')
df_hr4.insert(loc = 7, column = 'エラー', value = '')


###データのクリーニング
########## body4 ##########
#'氏名'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[0] == '' :
		df_bd4.iloc[index, 7] += 'a'
	elif type(row[0]) is not str :
		df_bd4.iloc[index, 7] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[1] == '' :
		df_bd4.iloc[index, 7] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bd4.iloc[index, 7] += 'd'
	else :
		df_bd4.iloc[index, 7] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[2] == '' :
		df_bd4.iloc[index, 7] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bd4.iloc[index, 7] += 'g'
	else :
		df_bd4.iloc[index, 7] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[3] == '' : 
		df_bd4.iloc[index, 7] +='i'
	elif type(row[3]) != int :
		df_bd4.iloc[index, 7] += 'j'
#'職業'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[4] == '':
		df_bd4.iloc[index, 7] += 'k'
	elif type(row[4]) != str :
		df_bd4.iloc[index, 7] += 'l'
#'結婚'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[5] == '' :
		df_bd4.iloc[index, 7] += 'm'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_bd4.iloc[index, 7] += 'n'
	else :
		df_bd4.iloc[index, 7] += 'o'
#'知った理由'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[6] == '' :
		df_bd4.iloc[index, 7] += 'p'
	elif type(row[6]) != str :
		df_bd4.iloc[index, 7] += 'q'

########## bust4 ##########
#'氏名'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[0] == '' :
		df_bs4.iloc[index, 7] += 'a'
	elif type(row[0]) is not str :
		df_bs4.iloc[index, 7] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[1] == '' :
		df_bs4.iloc[index, 7] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bs4.iloc[index, 7] += 'd'
	else :
		df_bs4.iloc[index, 7] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[2] == '' :
		df_bs4.iloc[index, 7] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bs4.iloc[index, 7] += 'g'
	else :
		df_bs4.iloc[index, 7] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[3] == '' : 
		df_bs4.iloc[index, 7] +='i'
	elif type(row[3]) != int :
		df_bs4.iloc[index, 7] += 'j'
#'職業'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[4] == '':
		df_bs4.iloc[index, 7] += 'k'
	elif type(row[4]) != str :
		df_bs4.iloc[index, 7] += 'l'
#'結婚'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[5] == '' :
		df_bs4.iloc[index, 7] += 'm'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_bs4.iloc[index, 7] += 'n'
	else :
		df_bs4.iloc[index, 7] += 'o'
#'知った理由'に例外がないか判定
for index, row in df_bs4.iterrows():
	if row[6] == '' :
		df_bs4.iloc[index, 7] += 'p'
	elif type(row[6]) != str :
		df_bs4.iloc[index, 7] += 'q'

########## facial4 ##########
#'氏名'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[0] == '' :
		df_fa4.iloc[index, 7] += 'a'
	elif type(row[0]) is not str :
		df_fa4.iloc[index, 7] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[1] == '' :
		df_fa4.iloc[index, 7] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_fa4.iloc[index, 7] += 'd'
	else :
		df_fa4.iloc[index, 7] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[2] == '' :
		df_fa4.iloc[index, 7] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_fa4.iloc[index, 7] += 'g'
	else :
		df_fa4.iloc[index, 7] += 'h'
#'年齢'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[3] == '' : 
		df_fa4.iloc[index, 7] +='i'
	elif type(row[3]) != int :
		df_fa4.iloc[index, 7] += 'j'
#'職業'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[4] == '':
		df_fa4.iloc[index, 7] += 'k'
	elif type(row[4]) != str :
		df_fa4.iloc[index, 7] += 'l'
#'結婚'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[5] == '' :
		df_fa4.iloc[index, 7] += 'm'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_fa4.iloc[index, 7] += 'n'
	else :
		df_fa4.iloc[index, 7] += 'o'
#'知った理由'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[6] == '' :
		df_fa4.iloc[index, 7] += 'p'
	elif type(row[6]) != str :
		df_fa4.iloc[index, 7] += 'q'

########## hairremoval4 ##########
#'氏名'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[0] == '' :
		df_hr4.iloc[index, 7] += 'a'
	elif type(row[0]) is not str :
		df_hr4.iloc[index, 7] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[1] == '' :
		df_hr4.iloc[index, 7] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_hr4.iloc[index, 7] += 'd'
	else :
		df_hr4.iloc[index, 7] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[2] == '' :
		df_hr4.iloc[index, 7] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_hr4.iloc[index, 7] += 'g'
	else :
		df_hr4.iloc[index, 7] += 'h'
#'年齢'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[3] == '' : 
		df_hr4.iloc[index, 7] +='i'
	elif type(row[3]) != int :
		df_hr4.iloc[index, 7] += 'j'
#'職業'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[4] == '':
		df_hr4.iloc[index, 7] += 'k'
	elif type(row[4]) != str :
		df_hr4.iloc[index, 7] += 'l'
#'結婚'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[5] == '' :
		df_hr4.iloc[index, 7] += 'm'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_hr4.iloc[index, 7] += 'n'
	else :
		df_hr4.iloc[index, 7] += 'o'
#'知った理由'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[6] == '' :
		df_hr4.iloc[index, 7] += 'p'
	elif type(row[6]) != str :
		df_hr4.iloc[index, 7] += 'q'

#データの結合
list_use = []
list_use_bd4 = []
list_use_bs4 = []
list_use_fa4 = []
list_use_hr4 = []
list_err_bd4 = []
list_err_bs4 = []
list_err_fa4 = []
list_err_hr4 = []
for index, row in df_bd4.iterrows():
	if row[7] == '':
		list_use_bd4.append(row[0:7])
	else :
		list_err_bd4.append(row[:])
for index, row in df_bs4.iterrows():
	if row[7] == '':
		list_use_bs4.append(row[0:7])
	else :
		list_err_bs4.append(row[:])
for index, row in df_fa4.iterrows():
	if row[7] == '':
		list_use_fa4.append(row[0:7])
	else :
		list_err_fa4.append(row[:])
for index, row in df_hr4.iterrows():
	if row[7] == '':
		list_use_hr4.append(row[0:7])
	else :
		list_err_hr4.append(row[:])

df_err_bd = pd.DataFrame(list_err_bd4)
df_err_bd = df_err_bd.drop_duplicates()
df_err_bs = pd.DataFrame(list_err_bs4)
df_err_bs = df_err_bs.drop_duplicates()
df_err_fa = pd.DataFrame(list_err_fa4)
df_err_fa = df_err_fa.drop_duplicates()
df_err_hr = pd.DataFrame(list_err_hr4)
df_err_hr = df_err_hr.drop_duplicates()

for index, row in df_err_bd.iterrows():
	if 'a' in row[7] :
		df_err_bd.loc[index, '氏名'] = '欠損'
	elif 'b' in row[7] :
		df_err_bd.loc[index, '氏名'] = '不正値'
	if 'c' in row[7] :
		df_err_bd.loc[index, '来店年月日'] = '欠損'
	elif 'd' in row[7] :
		df_err_bd.loc[index, '来店年月日'] = '不正値'
	elif 'e' in row[7] :
		df_err_bd.loc[index, '来店年月日'] = 'エラー値'
	else :
		df_err_bd.loc[index, '来店年月日'] = ''
	if 'f' in row[7] :
		df_err_bd.loc[index, '生年月日'] = '欠損'
	elif 'g' in row[7] :
		df_err_bd.loc[index, '生年月日'] = '不正値'
	elif 'h' in row[7] :
		df_err_bd.loc[index, '生年月日'] = 'エラー値'
	else :
		df_err_bd.loc[index, '生年月日'] = ''
	if 'i' in row[7] :
		df_err_bd.loc[index, '年齢'] = '欠損'
	elif 'j' in row[7] :
		df_err_bd.loc[index, '年齢'] = '不正値'
	else :
		df_err_bd.loc[index, '年齢'] = ''
	if 'k' in row[7] :
		df_err_bd.loc[index, '職業'] = '欠損'
	elif 'l' in row[7] :
		df_err_bd.loc[index, '職業'] = '不正値'
	else :
		df_err_bd.loc[index, '職業'] = ''
	if 'm' in row[7] :
		df_err_bd.loc[index, '結婚'] = '欠損'
	elif 'n' in row[7] :
		df_err_bd.loc[index, '結婚'] = '不正値'
	elif 'o' in row[7] :
		df_err_bd.loc[index, '結婚'] = 'エラー値'
	else :
		df_err_bd.loc[index, '結婚'] = ''
	if 'p' in row[7] :
		df_err_bd.loc[index, '知った理由'] = '欠損'
	elif 'q' in row[7] :
		df_err_bd.loc[index, '知った理由'] = '不正値'
	else :
		df_err_bd.loc[index, '知った理由'] = ''

for index, row in df_err_bs.iterrows():
	if 'a' in row[7] :
		df_err_bs.loc[index, '氏名'] = '欠損'
	elif 'b' in row[7] :
		df_err_bs.loc[index, '氏名'] = '不正値'
	if 'c' in row[7] :
		df_err_bs.loc[index, '来店年月日'] = '欠損'
	elif 'd' in row[7] :
		df_err_bs.loc[index, '来店年月日'] = '不正値'
	elif 'e' in row[7] :
		df_err_bs.loc[index, '来店年月日'] = 'エラー値'
	else :
		df_err_bs.loc[index, '来店年月日'] = ''
	if 'f' in row[7] :
		df_err_bs.loc[index, '生年月日'] = '欠損'
	elif 'g' in row[7] :
		df_err_bs.loc[index, '生年月日'] = '不正値'
	elif 'h' in row[7] :
		df_err_bs.loc[index, '生年月日'] = 'エラー値'
	else :
		df_err_bs.loc[index, '生年月日'] = ''
	if 'i' in row[7] :
		df_err_bs.loc[index, '年齢'] = '欠損'
	elif 'j' in row[7] :
		df_err_bs.loc[index, '年齢'] = '不正値'
	else :
		df_err_bs.loc[index, '年齢'] = ''
	if 'k' in row[7] :
		df_err_bs.loc[index, '職業'] = '欠損'
	elif 'l' in row[7] :
		df_err_bs.loc[index, '職業'] = '不正値'
	else :
		df_err_bs.loc[index, '職業'] = ''
	if 'm' in row[7] :
		df_err_bs.loc[index, '結婚'] = '欠損'
	elif 'n' in row[7] :
		df_err_bs.loc[index, '結婚'] = '不正値'
	elif 'o' in row[7] :
		df_err_bs.loc[index, '結婚'] = 'エラー値'
	else :
		df_err_bs.loc[index, '結婚'] = ''
	if 'p' in row[7] :
		df_err_bs.loc[index, '知った理由'] = '欠損'
	elif 'q' in row[7] :
		df_err_bs.loc[index, '知った理由'] = '不正値'
	else :
		df_err_bs.loc[index, '知った理由'] = ''

for index, row in df_err_fa.iterrows():
	if 'a' in row[7] :
		df_err_fa.loc[index, '氏名'] = '欠損'
	elif 'b' in row[7] :
		df_err_fa.loc[index, '氏名'] = '不正値'
	if 'c' in row[7] :
		df_err_fa.loc[index, '来店年月日'] = '欠損'
	elif 'd' in row[7] :
		df_err_fa.loc[index, '来店年月日'] = '不正値'
	elif 'e' in row[7] :
		df_err_fa.loc[index, '来店年月日'] = 'エラー値'
	else :
		df_err_fa.loc[index, '来店年月日'] = ''
	if 'f' in row[7] :
		df_err_fa.loc[index, '生年月日'] = '欠損'
	elif 'g' in row[7] :
		df_err_fa.loc[index, '生年月日'] = '不正値'
	elif 'h' in row[7] :
		df_err_fa.loc[index, '生年月日'] = 'エラー値'
	else :
		df_err_fa.loc[index, '生年月日'] = ''
	if 'i' in row[7] :
		df_err_fa.loc[index, '年齢'] = '欠損'
	elif 'j' in row[7] :
		df_err_fa.loc[index, '年齢'] = '不正値'
	else :
		df_err_fa.loc[index, '年齢'] = ''
	if 'k' in row[7] :
		df_err_fa.loc[index, '職業'] = '欠損'
	elif 'l' in row[7] :
		df_err_fa.loc[index, '職業'] = '不正値'
	else :
		df_err_fa.loc[index, '職業'] = ''
	if 'm' in row[7] :
		df_err_fa.loc[index, '結婚'] = '欠損'
	elif 'n' in row[7] :
		df_err_fa.loc[index, '結婚'] = '不正値'
	elif 'o' in row[7] :
		df_err_fa.loc[index, '結婚'] = 'エラー値'
	else :
		df_err_fa.loc[index, '結婚'] = ''
	if 'p' in row[7] :
		df_err_fa.loc[index, '知った理由'] = '欠損'
	elif 'q' in row[7] :
		df_err_fa.loc[index, '知った理由'] = '不正値'
	else :
		df_err_fa.loc[index, '知った理由'] = ''

for index, row in df_err_hr.iterrows():
	if 'a' in row[7] :
		df_err_hr.loc[index, '氏名'] = '欠損'
	elif 'b' in row[7] :
		df_err_hr.loc[index, '氏名'] = '不正値'
	if 'c' in row[7] :
		df_err_hr.loc[index, '来店年月日'] = '欠損'
	elif 'd' in row[7] :
		df_err_hr.loc[index, '来店年月日'] = '不正値'
	elif 'e' in row[7] :
		df_err_hr.loc[index, '来店年月日'] = 'エラー値'
	else :
		df_err_hr.loc[index, '来店年月日'] = ''
	if 'f' in row[7] :
		df_err_hr.loc[index, '生年月日'] = '欠損'
	elif 'g' in row[7] :
		df_err_hr.loc[index, '生年月日'] = '不正値'
	elif 'h' in row[7] :
		df_err_hr.loc[index, '生年月日'] = 'エラー値'
	else :
		df_err_hr.loc[index, '生年月日'] = ''
	if 'i' in row[7] :
		df_err_hr.loc[index, '年齢'] = '欠損'
	elif 'j' in row[7] :
		df_err_hr.loc[index, '年齢'] = '不正値'
	else :
		df_err_hr.loc[index, '年齢'] = ''
	if 'k' in row[7] :
		df_err_hr.loc[index, '職業'] = '欠損'
	elif 'l' in row[7] :
		df_err_hr.loc[index, '職業'] = '不正値'
	else :
		df_err_hr.loc[index, '職業'] = ''
	if 'm' in row[7] :
		df_err_hr.loc[index, '結婚'] = '欠損'
	elif 'n' in row[7] :
		df_err_hr.loc[index, '結婚'] = '不正値'
	elif 'o' in row[7] :
		df_err_hr.loc[index, '結婚'] = 'エラー値'
	else :
		df_err_hr.loc[index, '結婚'] = ''
	if 'p' in row[7] :
		df_err_hr.loc[index, '知った理由'] = '欠損'
	elif 'q' in row[7] :
		df_err_hr.loc[index, '知った理由'] = '不正値'
	else :
		df_err_hr.loc[index, '知った理由'] = ''

list_use = list_use_bd4 + list_use_bs4 + list_use_fa4 + list_use_hr4

df_use = pd.DataFrame(list_use)
df_use = df_use.drop_duplicates()
df_use = df_use.set_index('氏名')

#Excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df_use.to_excel(writer, sheet_name = sheet1)
	df_err_bd.to_excel(writer, sheet_name = sheet2)
	df_err_bs.to_excel(writer, sheet_name = sheet3)
	df_err_fa.to_excel(writer, sheet_name = sheet4)
	df_err_hr.to_excel(writer, sheet_name = sheet5)

conn.commit()
conn.close()
