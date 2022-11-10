# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

# ハッシュチェインの警告を無視
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_facial.xlsx'
sheet = '有効データ_facial'

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   facial   #############################
#facialのデータを抽出
df_facial_1 = pd.read_sql_query('select * from questionnaire_facial', conn)
df_fa1 = df_facial_1.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '肌気になるところ', '手入れ_朝', '手入れ_夜']]
df_facial_2 = pd.read_sql_query('select * from questionnaire_facial_2', conn)
df_fa2 = df_facial_2.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '肌気になるところ', '手入れ_朝', '手入れ_夜']]
df_facial_3 = pd.read_sql_query('select * from questionnaire_facial_3', conn)
df_fa3 = df_facial_3.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '肌気になるところ', '手入れ_朝', '手入れ_夜']]
df_facial_4 = pd.read_sql_query('select * from questionnaire_facial_4', conn)
df_fa4 = df_facial_4.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '血液型', '結婚', '知った理由', '肌気になるところ', '手入れ_朝', '手入れ_夜']]

###dfにエラー値を保持するカラムを追加
df_fa1.insert(loc = 11, column = 'エラー', value = '')
df_fa2.insert(loc = 11, column = 'エラー', value = '')
df_fa3.insert(loc = 11, column = 'エラー', value = '')
df_fa4.insert(loc = 11, column = 'エラー', value = '')

###データのクリーニング
########## facial_1 ##########
#'氏名'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[0] == '' :
		df_fa1.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_fa1.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[1] == '' :
		df_fa1.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_fa1.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_fa1.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[2] == '' :
		df_fa1.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_fa1.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_fa1.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[3] == '' : 
		df_fa1.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_fa1.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[4] == '':
		df_fa1.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_fa1.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[5] == '' :
		df_fa1.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_fa1.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[6] == '' :
		df_fa1.iloc[index, 11] += 'o'
	elif type(row[6]) != str :
		df_fa1.iloc[index, 11] += 'p'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_fa1.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_fa1.iterrows():
	if row[7] == '' :
		df_fa1.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_fa1.iloc[index, 11] += 's'
#'肌気になるところ'に例外がないか判定
for index, row in df_fa1.iterrows():
	if type(row[8]) != str :
		df_fa1.iloc[index, 11] += 'u'
#'手入れ_朝'に例外がないか判定
for index, row in df_fa1.iterrows():
	if type(row[9]) != str :
		df_fa1.iloc[index, 11] += 'w'
#'手入れ_夜'に例外がないか判定
for index, row in df_fa1.iterrows():
	if type(row[10]) != str :
		df_fa1.iloc[index, 11] += 'y'
########## facial_2 ##########
#'氏名'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[0] == '' :
		df_fa2.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_fa2.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[1] == '' :
		df_fa2.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_fa2.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_fa2.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[2] == '' :
		df_fa2.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_fa2.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_fa2.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[3] == '' : 
		df_fa2.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_fa2.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[4] == '':
		df_fa2.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_fa2.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[5] == '' :
		df_fa2.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_fa2.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[6] == '' :
		df_fa2.iloc[index, 11] += 'o'
	elif type(row[6]) != str :
		df_fa2.iloc[index, 11] += 'p'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_fa2.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_fa2.iterrows():
	if row[7] == '' :
		df_fa2.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_fa2.iloc[index, 11] += 's'
#'肌気になるところ'に例外がないか判定
for index, row in df_fa2.iterrows():
	if type(row[8]) != str :
		df_fa2.iloc[index, 11] += 'u'
#'手入れ_朝'に例外がないか判定
for index, row in df_fa2.iterrows():
	if type(row[9]) != str :
		df_fa2.iloc[index, 11] += 'w'
#'手入れ_夜'に例外がないか判定
for index, row in df_fa2.iterrows():
	if type(row[10]) != str :
		df_fa2.iloc[index, 11] += 'y'
########## facial_3 ##########
#'氏名'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[0] == '' :
		df_fa3.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_fa3.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[1] == '' :
		df_fa3.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_fa3.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_fa3.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[2] == '' :
		df_fa3.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_fa3.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_fa3.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[3] == '' : 
		df_fa3.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_fa3.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[4] == '':
		df_fa3.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_fa3.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[5] == '' :
		df_fa3.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_fa3.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[6] == '' :
		df_fa3.iloc[index, 11] += 'o'
	elif type(row[6]) != str :
		df_fa3.iloc[index, 11] += 'p'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_fa3.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_fa3.iterrows():
	if row[7] == '' :
		df_fa3.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_fa3.iloc[index, 11] += 's'
#'肌気になるところ'に例外がないか判定
for index, row in df_fa3.iterrows():
	if type(row[8]) != str :
		df_fa3.iloc[index, 11] += 'u'
#'手入れ_朝'に例外がないか判定
for index, row in df_fa3.iterrows():
	if type(row[9]) != str :
		df_fa3.iloc[index, 11] += 'w'
#'手入れ_夜'に例外がないか判定
for index, row in df_fa3.iterrows():
	if type(row[10]) != str :
		df_fa3.iloc[index, 11] += 'y'
########## facial4 ##########
#'氏名'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[0] == '' :
		df_fa4.iloc[index, 11] += 'a'
	elif type(row[0]) != str :
		df_fa4.iloc[index, 11] += 'b'
#'来店年月日'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[1] == '' :
		df_fa4.iloc[index, 11] += 'c'
	elif type(row[1]) != str :
		df_fa4.iloc[index, 11] += 'd'
	elif type(row[1]) == str :
		if len(row[1]) != 19 :
			df_fa4.iloc[index, 11] += 'e'
#'生年月日'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[2] == '' :
		df_fa4.iloc[index, 11] += 'f'
	elif type(row[2]) != str :
		df_fa4.iloc[index, 11] += 'g'
	elif type(row[2]) == str :
		if len(row[2]) != 19 :
			df_fa4.iloc[index, 11] += 'h'
#'年齢'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[3] == '' : 
		df_fa4.iloc[index, 11] +='i'
	elif type(row[3]) != int :
		df_fa4.iloc[index, 11] += 'j'
#'職業'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[4] == '':
		df_fa4.iloc[index, 11] += 'k'
	elif type(row[4]) != str :
		df_fa4.iloc[index, 11] += 'l'
#'血液型'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[5] == '' :
		df_fa4.iloc[index, 11] += 'm'
	elif row[5]!='A' and row[5]!='B' and row[5]!='AB' and row[5]!='O' :
		df_fa4.iloc[index, 11] += 'n'
#'結婚'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[6] == '' :
		df_fa4.iloc[index, 11] += 'o'
	elif type(row[6]) != str :
		df_fa4.iloc[index, 11] += 'p'
	elif type(row[6]) == str :
		if row[6] != '既婚' and row[6] != '未婚' :
			df_fa4.iloc[index, 11] += 'q'
#'知った理由'に例外がないか判定
for index, row in df_fa4.iterrows():
	if row[7] == '' :
		df_fa4.iloc[index, 11] += 'r'
	elif row[7] != 'HP' and row[7] != 'WEB' and row[7] != 'HPB' and row[7] != 'チラシ' and row[7] != '紹介' and row[7] != 'その他' :
		df_fa4.iloc[index, 11] += 's'
#'肌気になるところ'に例外がないか判定
for index, row in df_fa4.iterrows():
	if type(row[8]) != str :
		df_fa4.iloc[index, 11] += 'u'
#'手入れ_朝'に例外がないか判定
for index, row in df_fa4.iterrows():
	if type(row[9]) != str :
		df_fa4.iloc[index, 11] += 'w'
#'手入れ_夜'に例外がないか判定
for index, row in df_fa4.iterrows():
	if type(row[10]) != str :
		df_fa4.iloc[index, 11] += 'y'

#データの結合
list_use_fa1 = []
for index, row in df_fa1.iterrows():
	if row[11] == '':
		list_use_fa1.append(row[0:11])
list_use_fa2 = []
for index, row in df_fa2.iterrows():
	if row[11] == '':
		list_use_fa2.append(row[0:11])
list_use_fa3 = []
for index, row in df_fa3.iterrows():
	if row[11] == '':
		list_use_fa3.append(row[0:11])
list_use_fa4 = []
for index, row in df_fa4.iterrows():
	if row[11] == '':
		list_use_fa4.append(row[0:11])
list_use_fa = list_use_fa1 + list_use_fa2 + list_use_fa3 + list_use_fa4
df_use = pd.DataFrame(list_use_fa)
df_use = df_use.drop_duplicates()
df_use = df_use.set_index('氏名')

#Excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df_use.to_excel(writer, sheet_name = sheet)

# データベースとの接続を切断
conn.close()
