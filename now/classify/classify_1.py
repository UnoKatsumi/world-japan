# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
from datetime import datetime
from monthdelta import monthmod
from openpyxl import utils

#pd.set_option('display.max_rows', None)
pd.set_option('mode.chained_assignment', None)

path = 'S:\個人作業用\渡邊\ワールドジャパン\有効データ_all.xlsx'
sheet = 'Sheet1'

conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')

############################   body   #############################
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

#bodyの結合
# df_err_bd_y = pd.merge(df_err_bd1_y, df_err_bd2_y, how = 'outer')
# df_err_bd_y = pd.merge(df_err_bd_y, df_err_bd3_y, how = 'outer')
# df_bd = pd.merge(df_bd1, df_bd2, how = 'outer')
# df_bd = pd.merge(df_bd, df_bd3, how = 'outer')
# df_bd = pd.merge(df_bd, df_err_bd_y, how = 'outer')

# df_err_bd_y = df_err_bd4_y
df_bd = df_bd4
# df_bd = pd.merge(df_err_bd_y, df_bd4, how = 'outer')


############################   bust   #############################
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

#bustの結合
# df_err_bs_y = pd.merge(df_err_bs1_y, df_err_bs2_y, how = 'outer')
# df_err_bs_y = pd.merge(df_err_bs_y, df_err_bs3_y, how = 'outer')
# df_bs = pd.merge(df_bs1, df_bs2, how = 'outer')
# df_bs = pd.merge(df_bs, df_bs3, how = 'outer')
# df_bs = pd.merge(df_bs, df_err_bs_y, how = 'outer')

# df_err_bs_y = df_err_bs4_y
df_bs = df_bs4
# df_bs = pd.merge(df_bs, df_err_bs_y, how = 'outer')


############################   facial   #############################
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

# df_err_fa_y = df_err_fa4_y
df_fa = df_fa4
# df_fa = pd.merge(df_fa, df_err_fa_y, how = 'outer')


############################   hair removal   #############################
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
