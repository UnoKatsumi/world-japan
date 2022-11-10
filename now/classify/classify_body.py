# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

# ハッシュチェインの警告を無視
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_body.xlsx'
sheet = '有効データ_body'

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   body   #############################
#bodyのデータを抽出
df_body_1 = pd.read_sql_query('select * from questionnaire_body', conn)
df_body_2 = pd.read_sql_query('select * from questionnaire_body_2', conn)
df_body_3 = pd.read_sql_query('select * from questionnaire_body_3', conn)
df_body_4 = pd.read_sql_query('select * from questionnaire_body_4', conn)
df_bd1 = df_body_1.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '身長', '体重', '目標体重']]
df_bd2 = df_body_2.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '身長', '体重', '目標体重']]
df_bd3 = df_body_3.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '身長', '体重', '目標体重']]
df_bd4 = df_body_4.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '身長', '体重', '目標体重']]

###各dfにエラー値を保持するカラムを追加
df_bd1.insert(loc = 11, column = 'エラー', value = '')
df_bd2.insert(loc = 11, column = 'エラー', value = '')
df_bd3.insert(loc = 11, column = 'エラー', value = '')
df_bd4.insert(loc = 11, column = 'エラー', value = '')


###データのクリーニング
########## body1 ##########
#'氏名'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[0] == '' :
		df_bd1.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_bd1.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[1] == '' :
		df_bd1.iloc[index, 11] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bd1.iloc[index, 11] += 'd'
	else :
		df_bd1.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[2] == '' :
		df_bd1.iloc[index, 11] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bd1.iloc[index, 11] += 'g'
	else :
			df_bd1.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[3] == '' : 
		df_bd1.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_bd1.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[4] == '':
		df_bd1.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_bd1.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[5] == '' :
		df_bd1.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_bd1.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[6] == '' :
		df_bd1.iloc[index, 11] += 'o'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_bd1.iloc[index, 11] += 'p'
	else : 
		df_bd1.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[7] == '' :
		df_bd1.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_bd1.iloc[index, 11] += 's'
#'身長'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[8] == '' :
		df_bd1.iloc[index, 11] += 't'
	elif type(row[8]) != int :
		df_bd1.iloc[index, 11] += 'u'
#'体重'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[9] == '' :
		df_bd1.iloc[index, 11] += 'v'
	elif type(row[9]) != int :
		df_bd1.iloc[index, 11] += 'w'
#'目標体重'に例外がないか判定
for index, row in df_bd1.iterrows():
	if row[10] == '' :
		df_bd1.iloc[index, 11] += 'x'
	elif type(row[10]) != int :
		df_bd1.iloc[index, 11] += 'y'
########## body_2 ##########
#'氏名'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[0] == '' :
		df_bd2.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_bd2.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[1] == '' :
		df_bd2.iloc[index, 11] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bd2.iloc[index, 11] += 'd'
	else :
		df_bd2.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[2] == '' :
		df_bd2.iloc[index, 11] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bd2.iloc[index, 11] += 'g'
	else :
			df_bd2.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[3] == '' : 
		df_bd2.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_bd2.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[4] == '':
		df_bd2.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_bd2.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[5] == '' :
		df_bd2.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_bd2.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[6] == '' :
		df_bd2.iloc[index, 11] += 'o'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_bd2.iloc[index, 11] += 'p'
	else : 
		df_bd2.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[7] == '' :
		df_bd2.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_bd2.iloc[index, 11] += 's'
#'身長'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[8] == '' :
		df_bd2.iloc[index, 11] += 't'
	elif type(row[8]) != int :
		df_bd2.iloc[index, 11] += 'u'
#'体重'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[9] == '' :
		df_bd2.iloc[index, 11] += 'v'
	elif type(row[9]) != int :
		df_bd2.iloc[index, 11] += 'w'
#'目標体重'に例外がないか判定
for index, row in df_bd2.iterrows():
	if row[10] == '' :
		df_bd2.iloc[index, 11] += 'x'
	elif type(row[10]) != int :
		df_bd2.iloc[index, 11] += 'y'
########## body_3 ##########
#'氏名'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[0] == '' :
		df_bd3.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_bd3.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[1] == '' :
		df_bd3.iloc[index, 11] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bd3.iloc[index, 11] += 'd'
	else :
		df_bd3.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[2] == '' :
		df_bd3.iloc[index, 11] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bd3.iloc[index, 11] += 'g'
	else :
			df_bd3.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[3] == '' : 
		df_bd3.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_bd3.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[4] == '':
		df_bd3.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_bd3.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[5] == '' :
		df_bd3.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_bd3.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[6] == '' :
		df_bd3.iloc[index, 11] += 'o'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_bd3.iloc[index, 11] += 'p'
	else : 
		df_bd3.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[7] == '' :
		df_bd3.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_bd3.iloc[index, 11] += 's'
#'身長'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[8] == '' :
		df_bd3.iloc[index, 11] += 't'
	elif type(row[8]) != int :
		df_bd3.iloc[index, 11] += 'u'
#'体重'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[9] == '' :
		df_bd3.iloc[index, 11] += 'v'
	elif type(row[9]) != int :
		df_bd3.iloc[index, 11] += 'w'
#'目標体重'に例外がないか判定
for index, row in df_bd3.iterrows():
	if row[10] == '' :
		df_bd3.iloc[index, 11] += 'x'
	elif type(row[10]) != int :
		df_bd3.iloc[index, 11] += 'y'
########## body_4 ##########
#'氏名'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[0] == '' :
		df_bd4.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_bd4.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[1] == '' :
		df_bd4.iloc[index, 11] += 'c'
	elif type(row[1]) is str :
		if len(row[1]) != 19 :
			df_bd4.iloc[index, 11] += 'd'
	else :
		df_bd4.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[2] == '' :
		df_bd4.iloc[index, 11] += 'f'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_bd4.iloc[index, 11] += 'g'
	else :
			df_bd4.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[3] == '' : 
		df_bd4.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_bd4.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[4] == '':
		df_bd4.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_bd4.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[5] == '' :
		df_bd4.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_bd4.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[6] == '' :
		df_bd4.iloc[index, 11] += 'o'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_bd4.iloc[index, 11] += 'p'
	else : 
		df_bd4.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[7] == '' :
		df_bd4.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_bd4.iloc[index, 11] += 's'
#'身長'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[8] == '' :
		df_bd4.iloc[index, 11] += 't'
	elif type(row[8]) != int :
		df_bd4.iloc[index, 11] += 'u'
#'体重'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[9] == '' :
		df_bd4.iloc[index, 11] += 'v'
	elif type(row[9]) != int :
		df_bd4.iloc[index, 11] += 'w'
#'目標体重'に例外がないか判定
for index, row in df_bd4.iterrows():
	if row[10] == '' :
		df_bd4.iloc[index, 11] += 'x'
	elif type(row[10]) != int :
		df_bd4.iloc[index, 11] += 'y'

#有効データの抽出
list_use_bd1 = []
for index, row in df_bd1.iterrows():
	if row[11] == '':
		list_use_bd1.append(row[0:11])
list_use_bd2 = []
for index, row in df_bd2.iterrows():
	if row[11] == '':
		list_use_bd2.append(row[0:11])
list_use_bd3 = []
for index, row in df_bd3.iterrows():
	if row[11] == '':
		list_use_bd3.append(row[0:11])
list_use_bd4 = []
for index, row in df_bd4.iterrows():
	if row[11] == '':
		list_use_bd4.append(row[0:11])

list_use_bd = list_use_bd1 + list_use_bd2 + list_use_bd3 + list_use_bd4

df_use = pd.DataFrame(list_use_bd)
df_use = df_use.drop_duplicates()
df_use = df_use.set_index('氏名')

#Excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df_use.to_excel(writer, sheet_name = sheet)

# データベースとの接続を切断
conn.close()
