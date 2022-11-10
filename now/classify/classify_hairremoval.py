# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

# ハッシュチェインの警告を無視
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_hairremoval.xlsx'
sheet = '有効データ_hairremoval'

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   facial   #############################
#hairremoval_1のデータを抽出
df_hairremoval_1 = pd.read_sql_query('select * from questionnaire_hairremoval', conn)
df_hr1 = df_hairremoval_1.loc[:,['氏名', '来店年月日', '生年月日', '職業', '既婚・未婚', '知った理由', '脱毛経験', '自己処理部位', '自己処理方法', '希望脱毛箇所']]
#hairremoval_1の'年齢'が欠損しているため年齢を計算して設定
list_hr1_age = []
list_hr1_age_err = []
year_1 = []
for row in df_hr1.iterrows():
	if type(row[1][1]) == str and type(row[1][2]) == str :
		if ':' in row[1][1] and ':' in row[1][2] :
			date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
			date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		year_1.append(abs(old))
	else :
		year_1.append('')
df_hr1.insert(loc = 3, column = '年齢', value = year_1)
#hairremoval_1の'既婚・未婚'カラムの名前を'結婚'へ変更
df_hr1 = df_hr1.rename(columns ={'既婚・未婚' : '結婚'})
#hairremoval_2のデータを抽出
df_hairremoval_2 = pd.read_sql_query('select * from questionnaire_hairremoval_2', conn)
df_hr2 = df_hairremoval_2.loc[:,['氏名', '来店年月日', '生年月日', '職業', '既婚・未婚', '知った理由', '脱毛経験', '自己処理部位', '自己処理方法', '希望脱毛箇所']]
#hairremoval_2の'年齢'が欠損しているため年齢を計算して設定
list_hr2_age = []
list_hr2_age_err = []
year_2 = []
for row in df_hr2.iterrows():
	if type(row[1][1]) == str and type(row[1][2]) == str :
		if ':' in row[1][1] and ':' in row[1][2] :
			date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
			date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		year_2.append(abs(old))
	else :
		year_2.append('')
df_hr2.insert(loc = 3, column = '年齢', value = year_2)
#hairremoval_2の'既婚・未婚'カラムの名前を'結婚'へ変更
df_hr2 = df_hr2.rename(columns ={'既婚・未婚' : '結婚'})
#hairremoval_3のデータを抽出
df_hairremoval_3 = pd.read_sql_query('select * from questionnaire_hairremoval_3', conn)
df_hr3 = df_hairremoval_3.loc[:,['氏名', '来店年月日', '生年月日', '職業', '既婚・未婚', '知った理由', '脱毛経験', '自己処理部位', '自己処理方法', '希望脱毛箇所']]
#hairremoval_3の'年齢'が欠損しているため年齢を計算して設定
list_hr3_age = []
list_hr3_age_err = []
year_3 = []
for row in df_hr3.iterrows():
	if type(row[1][1]) == str and type(row[1][2]) == str :
		if ':' in row[1][1] and ':' in row[1][2] :
			date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
			date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		year_3.append(abs(old))
	else :
		year_3.append('')
df_hr3.insert(loc = 3, column = '年齢', value = year_3)
#hairremoval_3の'既婚・未婚'カラムの名前を'結婚'へ変更
df_hr3 = df_hr3.rename(columns ={'既婚・未婚' : '結婚'})
#hairremoval_4のデータを抽出
df_hairremoval_4 = pd.read_sql_query('select * from questionnaire_hairremoval_4', conn)
df_hr4 = df_hairremoval_4.loc[:,['氏名', '来店年月日', '生年月日', '職業', '既婚・未婚', '知った理由', '脱毛経験', '自己処理部位', '自己処理方法', '希望脱毛箇所']]
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
###dfにエラー値を保持するカラムを追加
df_hr1.insert(loc = 11, column = 'エラー', value = '')
df_hr2.insert(loc = 11, column = 'エラー', value = '')
df_hr3.insert(loc = 11, column = 'エラー', value = '')
df_hr4.insert(loc = 11, column = 'エラー', value = '')


###データのクリーニング
########## facial_1 ##########
#'氏名'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[0] == '' :
		df_hr1.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_hr1.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[1] == '' :
		df_hr1.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_hr1.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_hr1.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[2] == '' :
		df_hr1.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_hr1.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_hr1.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[3] == '' : 
		df_hr1.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_hr1.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[4] == '':
		df_hr1.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_hr1.iloc[index, 11] += 'l'
#'結婚'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[5] == '' :
		df_hr1.iloc[index, 11] += 'm'
	elif type(row[5]) != str :
		df_hr1.iloc[index, 11] += 'n'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_hr1.iloc[index, 11] += 'l'
#'知った理由'に例外がないか判定
for index, row in df_hr1.iterrows():
	if row[6] == '' :
		df_hr1.iloc[index, 11] += 'o'
	elif row[6] != 'HP' and row[6] != 'WEB' and row[6] != 'HPB' and row[6] != 'チラシ' and row[6] != '紹介' and row[6] != 'その他' :
		df_hr1.iloc[index, 11] += 'p'
#'脱毛経験'に例外がないか判定
for index, row in df_hr1.iterrows():
	if type(row[7]) != str :
		df_hr1.iloc[index, 11] += 'q'
	elif type(row[7]) == str :
	  if row[7] != 'ある' and row[7] != 'ない' :
		  df_hr1.iloc[index, 11] += 'r'
#'自己処理部位'に例外がないか判定
for index, row in df_hr1.iterrows():
	if type(row[8]) != str :
		df_hr1.iloc[index, 11] += 's'
#'自己処理方法'に例外がないか判定
for index, row in df_hr1.iterrows():
	if type(row[9]) != str :
		df_hr1.iloc[index, 11] += 't'
#'希望脱毛箇所'に例外がないか判定
for index, row in df_hr1.iterrows():
	if type(row[10]) != str :
		df_hr1.iloc[index, 11] += 'u'
########## facial_2 ##########
#'氏名'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[0] == '' :
		df_hr2.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_hr2.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[1] == '' :
		df_hr2.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_hr2.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_hr2.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[2] == '' :
		df_hr2.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_hr2.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_hr2.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[3] == '' : 
		df_hr2.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_hr2.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[4] == '':
		df_hr2.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_hr2.iloc[index, 11] += 'l'
#'結婚'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[5] == '' :
		df_hr2.iloc[index, 11] += 'm'
	elif type(row[5]) != str :
		df_hr2.iloc[index, 11] += 'n'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_hr2.iloc[index, 11] += 'l'
#'知った理由'に例外がないか判定
for index, row in df_hr2.iterrows():
	if row[6] == '' :
		df_hr2.iloc[index, 11] += 'o'
	elif row[6] != 'HP' and row[6] != 'WEB' and row[6] != 'HPB' and row[6] != 'チラシ' and row[6] != '紹介' and row[6] != 'その他' :
		df_hr2.iloc[index, 11] += 'p'
#'脱毛経験'に例外がないか判定
for index, row in df_hr2.iterrows():
	if type(row[7]) != str :
		df_hr2.iloc[index, 11] += 'q'
	elif type(row[7]) == str :
	  if row[7] != 'ある' and row[7] != 'ない' :
		  df_hr2.iloc[index, 11] += 'r'
#'自己処理部位'に例外がないか判定
for index, row in df_hr2.iterrows():
	if type(row[8]) != str :
		df_hr2.iloc[index, 11] += 's'
#'自己処理方法'に例外がないか判定
for index, row in df_hr2.iterrows():
	if type(row[9]) != str :
		df_hr2.iloc[index, 11] += 't'
#'希望脱毛箇所'に例外がないか判定
for index, row in df_hr2.iterrows():
	if type(row[10]) != str :
		df_hr2.iloc[index, 11] += 'u'
########## facial_3 ##########
#'氏名'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[0] == '' :
		df_hr3.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_hr3.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[1] == '' :
		df_hr3.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_hr3.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_hr3.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[2] == '' :
		df_hr3.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_hr3.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_hr3.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[3] == '' : 
		df_hr3.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_hr3.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[4] == '':
		df_hr3.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_hr3.iloc[index, 11] += 'l'
#'結婚'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[5] == '' :
		df_hr3.iloc[index, 11] += 'm'
	elif type(row[5]) != str :
		df_hr3.iloc[index, 11] += 'n'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_hr3.iloc[index, 11] += 'l'
#'知った理由'に例外がないか判定
for index, row in df_hr3.iterrows():
	if row[6] == '' :
		df_hr3.iloc[index, 11] += 'o'
	elif row[6] != 'HP' and row[6] != 'WEB' and row[6] != 'HPB' and row[6] != 'チラシ' and row[6] != '紹介' and row[6] != 'その他' :
		df_hr3.iloc[index, 11] += 'p'
#'脱毛経験'に例外がないか判定
for index, row in df_hr3.iterrows():
	if type(row[7]) != str :
		df_hr3.iloc[index, 11] += 'q'
	elif type(row[7]) == str :
	  if row[7] != 'ある' and row[7] != 'ない' :
		  df_hr3.iloc[index, 11] += 'r'
#'自己処理部位'に例外がないか判定
for index, row in df_hr3.iterrows():
	if type(row[8]) != str :
		df_hr3.iloc[index, 11] += 's'
#'自己処理方法'に例外がないか判定
for index, row in df_hr3.iterrows():
	if type(row[9]) != str :
		df_hr3.iloc[index, 11] += 't'
#'希望脱毛箇所'に例外がないか判定
for index, row in df_hr3.iterrows():
	if type(row[10]) != str :
		df_hr3.iloc[index, 11] += 'u'
########## facial4 ##########
#'氏名'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[0] == '' :
		df_hr4.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_hr4.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[1] == '' :
		df_hr4.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_hr4.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_hr4.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[2] == '' :
		df_hr4.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_hr4.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_hr4.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[3] == '' : 
		df_hr4.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_hr4.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[4] == '':
		df_hr4.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_hr4.iloc[index, 11] += 'l'
#'結婚'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[5] == '' :
		df_hr4.iloc[index, 11] += 'm'
	elif type(row[5]) != str :
		df_hr4.iloc[index, 11] += 'n'
	elif type(row[5]) == str :
		if row[5] != '既婚' and row[5] != '未婚' :
			df_hr4.iloc[index, 11] += 'l'
#'知った理由'に例外がないか判定
for index, row in df_hr4.iterrows():
	if row[6] == '' :
		df_hr4.iloc[index, 11] += 'o'
	elif row[6] != 'HP' and row[6] != 'WEB' and row[6] != 'HPB' and row[6] != 'チラシ' and row[6] != '紹介' and row[6] != 'その他' :
		df_hr4.iloc[index, 11] += 'p'
#'脱毛経験'に例外がないか判定
for index, row in df_hr4.iterrows():
	if type(row[7]) != str :
		df_hr4.iloc[index, 11] += 'q'
	elif type(row[7]) == str :
	  if row[7] != 'ある' and row[7] != 'ない' :
		  df_hr4.iloc[index, 11] += 'r'
#'自己処理部位'に例外がないか判定
for index, row in df_hr4.iterrows():
	if type(row[8]) != str :
		df_hr4.iloc[index, 11] += 's'
#'自己処理方法'に例外がないか判定
for index, row in df_hr4.iterrows():
	if type(row[9]) != str :
		df_hr4.iloc[index, 11] += 't'
#'希望脱毛箇所'に例外がないか判定
for index, row in df_hr4.iterrows():
	if type(row[10]) != str :
		df_hr4.iloc[index, 11] += 'u'

#データの結合
list_use_hr1 = []
for index, row in df_hr1.iterrows():
	if row[11] == '':
		list_use_hr1.append(row[0:11])
list_use_hr2 = []
for index, row in df_hr2.iterrows():
	if row[11] == '':
		list_use_hr2.append(row[0:11])
list_use_hr3 = []
for index, row in df_hr3.iterrows():
	if row[11] == '':
		list_use_hr3.append(row[0:11])
list_use_hr4 = []
for index, row in df_hr4.iterrows():
	if row[11] == '':
		list_use_hr4.append(row[0:11])
list_use_hr = list_use_hr1 + list_use_hr2 + list_use_hr3 + list_use_hr4
df_use = pd.DataFrame(list_use_hr)
df_use = df_use.drop_duplicates()
df_use = df_use.set_index('氏名')

#Excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df_use.to_excel(writer, sheet_name = sheet)

# データベースとの接続を切断
conn.close()
