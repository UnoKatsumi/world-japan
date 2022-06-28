# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

# pd.set_option('display.max_rows', None)
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\有効データ_all.xlsx'
sheet = 'Sheet1'

conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   body   #############################
# #body_1のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_body', conn)
# df_bd1 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bd1 = df_bd1.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bd1 = df_bd1.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_bd1) < len(df_err_bd1)) :
# 	df_err_bd1_y = df_err_bd1[df_err_bd1.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bd1_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bd1_y.loc[df_err_bd1_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_bd1_y = df_err_bd1.iloc[0:0]

# #body_2のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_body_2', conn)
# df_bd2 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bd2 = df_bd2.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bd2 = df_bd2.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_bd2) < len(df_err_bd2)) :
# 	df_err_bd2_y = df_err_bd2[df_err_bd2.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bd2_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bd2_y.loc[df_err_bd2_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_bd2_y = df_err_bd2.iloc[0:0]

# #body_3のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_body_3', conn)
# df_bd3 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bd3 = df_bd3.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bd3 = df_bd3.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_bd3) < len(df_err_bd3)) :
# 	df_err_bd3_y = df_err_bd3[df_err_bd3.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bd3_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bd3_y.loc[df_err_bd3_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_bd3_y = df_err_bd3.iloc[0:0]

#body_4のデータから各カラムに欠損値(nan)のない有効なデータを抽出
df = pd.read_sql_query('select * from questionnaire_body_4', conn)
df_bd4 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
df_err_bd4 = df_bd4.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
df_bd4 = df_bd4.dropna()
##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
if (len(df_bd4) < len(df_err_bd4)) :
	df_err_bd4_y = df_err_bd4[df_err_bd4.isnull().any(axis = 1)]
	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
	for row in df_err_bd4_y.iterrows():
		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		df_err_bd4_y.loc[df_err_bd4_y['氏名'] == row[1][0], '年齢'] = abs(old)
else :
	df_err_bd4_y = df_err_bd4.iloc[0:0]

# #bodyの結合
# df_err_bd_y = pd.merge(df_err_bd1_y, df_err_bd2_y, how = 'outer')
# df_err_bd_y = pd.merge(df_err_bd_y, df_err_bd3_y, how = 'outer')
# df_bd = pd.merge(df_bd1, df_bd2, how = 'outer')
# df_bd = pd.merge(df_bd, df_bd3, how = 'outer')
# df_bd = pd.merge(df_bd, df_err_bd_y, how = 'outer')

# df_err_bd_y = df_err_bd4_y
df_bd = df_bd4
# df_bd = pd.merge(df_bd, df_err_bd_y, how = 'outer')



############################   bust   #############################
# #bust_1のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_bust', conn)
# df_bs1 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bs1 = df_bs1.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bs1 = df_bs1.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# #print(len(df_bs1), len(df_err_bs1))
# if (len(df_bs1) < len(df_err_bs1)) :
# 	df_err_bs1_y = df_err_bs1[df_err_bs1.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bs1_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bs1_y.loc[df_err_bs1_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_bs1_y = df_err_bs1.iloc[0:0]

# #bust_2のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_bust_2', conn)
# df_bs2 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bs2 = df_bs2.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bs2 = df_bs2.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_bs2) < len(df_err_bs2)) :
# 	df_err_bs2_y = df_err_bs2[df_err_bs2.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bs2_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bs2_y.loc[df_err_bs2_y['氏名'] == row[1][0], '年齢'] = abs(old)
# 		print(df_err_bs2_y)
# else :
# 	df_err_bs2_y = df_err_bs2.iloc[0:0]

# #bust_3のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_bust_3', conn)
# df_bs3 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_bs3 = df_bs3.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_bs3 = df_bs3.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_bs3) < len(df_err_bs3)) :
# 	df_err_bs3_y = df_err_bs3[df_err_bs3.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_bs3_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_bs3_y.loc[df_err_bs3_y['氏名'] == row[1][0], '年齢'] = abs(old)
# 		print(df_err_bs3_y)
# else :
# 	df_err_bs3_y = df_err_bs3.iloc[0:0]

#bust_4のデータから各カラムに欠損値(nan)のない有効なデータを抽出
df = pd.read_sql_query('select * from questionnaire_bust_4', conn)
df_bs4 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
df_err_bs4 = df_bs4.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
df_bs4 = df_bs4.dropna()
##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
if (len(df_bs4) < len(df_err_bs4)) :
	df_err_bs4_y = df_err_bs4[df_err_bs4.isnull().any(axis = 1)]
	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
	for row in df_err_bs4_y.iterrows():
		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		df_err_bs4_y.loc[df_err_bs4_y['氏名'] == row[1][0], '年齢'] = abs(old)
		print(df_err_bs4_y)
else :
	df_err_bs4_y = df_err_bs4.iloc[0:0]

# #bustの結合
# df_err_bs_y = pd.merge(df_err_bs1_y, df_err_bs2_y, how = 'outer')
# df_err_bs_y = pd.merge(df_err_bs_y, df_err_bs3_y, how = 'outer')
# df_bs = pd.merge(df_bs1, df_bs2, how = 'outer')
# df_bs = pd.merge(df_bs, df_bs3, how = 'outer')
# df_bs = pd.merge(df_bs, df_err_bs_y, how = 'outer')

# df_err_bs_y = df_err_bs4_y
df_bs = df_bs4
# df_bs = pd.merge(df_err_bs_y, df_bs, how = 'outer')


############################   facial   #############################
# #facial_1のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_facial', conn)
# df_fa1 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_fa1 = df_fa1.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_fa1 = df_fa1.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_fa1) < len(df_err_fa1)) :
# 	df_err_fa1_y = df_err_fa1[df_err_fa1.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_fa1_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_fa1_y.loc[df_err_fa1_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_fa1_y = df_err_fa1.iloc[0:0]

# #facial_2のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_facial_2', conn)
# df_fa2 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_fa2 = df_fa2.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_fa2 = df_fa2.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_fa2) < len(df_err_fa2)) :
# 	df_err_fa2_y = df_err_fa2[df_err_fa2.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_fa2_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_fa2_y.loc[df_err_fa2_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_fa2_y = df_err_fa2.iloc[0:0]

# #facial_3のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_facial_3', conn)
# df_fa3 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
# df_err_fa3 = df_fa3.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
# df_fa3 = df_fa3.dropna()
# ##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
# if (len(df_fa3) < len(df_err_fa3)) :
# 	df_err_fa3_y = df_err_fa3[df_err_fa3.isnull().any(axis = 1)]
# 	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
# 	for row in df_err_fa3_y.iterrows():
# 		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 		mmod = monthmod(date_1, date_2)
# 		old = mmod[0].months//12
# 		df_err_fa3_y.loc[df_err_fa3_y['氏名'] == row[1][0], '年齢'] = abs(old)
# else :
# 	df_err_fa3_y = df_err_fa3.iloc[0:0]

#facial_4のデータから各カラムに欠損値(nan)のない有効なデータを抽出
df = pd.read_sql_query('select * from questionnaire_facial_4', conn)
df_fa4 = df.loc[:,['氏名', '来店年月日', '生年月日' ,'年齢', '職業', '結婚']] #, '知った理由', 'DM']]
df_err_fa4 = df_fa4.dropna(subset = ['来店年月日', '生年月日', '職業', '結婚']) #, '知った理由', 'DM'])
df_fa4 = df_fa4.dropna()
##年齢のみが欠損しているデータの年齢を求めて有効データとして保持する
if (len(df_fa4) < len(df_err_fa4)) :
	df_err_fa4_y = df_err_fa4[df_err_fa4.isnull().any(axis = 1)]
	#年齢が欠損しているデータは来店年月日と生年月日から年齢を計算して年齢として設定
	for row in df_err_fa4_y.iterrows():
		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
		mmod = monthmod(date_1, date_2)
		old = mmod[0].months//12
		df_err_fa4_y.loc[df_err_fa4_y['氏名'] == row[1][0], '年齢'] = abs(old)
else :
	df_err_fa4_y = df_err_fa4.iloc[0:0]

#facialの結合
# df_err_fa_y = pd.merge(df_err_fa1_y, df_err_fa2_y, how = 'outer')
# df_err_fa_y = pd.merge(df_err_fa_y, df_err_fa3_y, how = 'outer')
# df_fa = pd.merge(df_fa1, df_fa2, how = 'outer')
# df_fa = pd.merge(df_fa, df_fa3, how = 'outer')
# df_fa = pd.merge(df_fa, df_err_fa_y, how = 'outer')

# df_err_fa_y = df_err_fa4
df_fa = df_fa4
# df_fa = pd.merge(df_err_fa_y, df_fa, how = 'outer')


############################   hair removal   #############################
# #hairremoval_1のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_hairremoval', conn)
# df_hr1 = df.loc[:,['氏名', '来店年月日', '生年月日' , '職業', '既婚・未婚']] #, '知った理由', 'DM']]
# df_hr1 = df_hr1.dropna()
# year_1 = []
# #年齢が欠損しているため来店年月日と生年月日から年齢を計算して年齢として設定
# for row in df_hr1.iterrows():
# 	date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 	date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 	mmod = monthmod(date_1, date_2)
# 	old = mmod[0].months//12
# 	year_1.append(abs(old))
# df_hr1.insert(loc = 3, column = '年齢', value = year_1)

# #hairremoval_2のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_hairremoval_2', conn)
# df_hr2 = df.loc[:,['氏名', '来店年月日', '生年月日' , '職業', '既婚・未婚']] #, '知った理由', 'DM']]
# df_hr2 = df_hr2.dropna()
# year_2 = []
# #年齢が欠損しているため来店年月日と生年月日から年齢を計算して年齢として設定
# for row in df_hr2.iterrows():
# 	date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 	date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 	mmod = monthmod(date_1, date_2)
# 	old = mmod[0].months//12
# 	year_2.append(abs(old))
# df_hr2.insert(loc = 3, column = '年齢', value = year_2)

# #hairremoval_3のデータから各カラムに欠損値(nan)のない有効なデータを抽出
# df = pd.read_sql_query('select * from questionnaire_hairremoval_3', conn)
# df_hr3 = df.loc[:,['氏名', '来店年月日', '生年月日' , '職業', '既婚・未婚']] #, '知った理由', 'DM']]
# df_hr3 = df_hr3.dropna()
# year_3 = []
# #年齢が欠損しているため来店年月日と生年月日から年齢を計算して年齢として設定
# for row in df_hr3.iterrows():
# 	date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
# 	date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
# 	mmod = monthmod(date_1, date_2)
# 	old = mmod[0].months//12
# 	year_3.append(abs(old))
# df_hr3.insert(loc = 3, column = '年齢', value = year_3)

#hairremoval_4のデータから各カラムに欠損値(nan)のない有効なデータを抽出
df = pd.read_sql_query('select * from questionnaire_hairremoval_4', conn)
df_hr4 = df.loc[:,['氏名', '来店年月日', '生年月日' , '職業', '既婚・未婚']] #, '知った理由', 'DM']]
df_hr4 = df_hr4.dropna()
year_4 = []
#年齢が欠損しているため来店年月日と生年月日から年齢を計算して年齢として設定
for row in df_hr4.iterrows():
	if ':' in row[1][1] and ':' in row[1][2] :
		date_1 = datetime.strptime(row[1][1], '%Y-%m-%d %H:%M:%S')
		date_2 = datetime.strptime(row[1][2], '%Y-%m-%d %H:%M:%S')
	mmod = monthmod(date_1, date_2)
	old = mmod[0].months//12
	year_4.append(abs(old))
df_hr4.insert(loc = 3, column = '年齢', value = year_4)

#hairremovalの結合
# df_hr = pd.merge(df_hr1, df_hr2, how = 'outer')
# df_hr = pd.merge(df_hr, df_hr3, how = 'outer')
# df_hr = df_hr.rename( columns = {'既婚・未婚' : '結婚'} ) #'脱毛経験' : 'エステ体験', 

df_hr = df_hr4
df_hr = df_hr.rename( columns = {'既婚・未婚' : '結婚'} ) #'脱毛経験' : 'エステ体験', 


###最終的な結合
df_use = pd.merge(df_bd, df_bs, how = 'outer')
df_use = pd.merge(df_use, df_fa, how = 'outer')
df_use = pd.merge(df_use, df_hr, how = 'outer')
df_use = df_use.drop_duplicates()

df_use.to_excel(path, sheet_name = sheet)

conn.commit()
conn.close()
