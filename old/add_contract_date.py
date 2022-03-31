# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
import datetime
from openpyxl import utils

#pd.set_option('display.max_rows', None)

path = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/2フォーマット/【E】エステティックサービス契約書.xlsx'
sheet = '【E】エステティックサービス契約書'
path_dat = 'S:\個人作業用\那須\ワールドジャパン\有効データ候補_契約日.xlsx'
sheet_dat = 'Sheet1'

#conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')

#df = pd.read_sql_query('select * from visit_date', conn)
#df_period = pd.read_sql_query('select * from visit_date', conn)
#df_period_2 = pd.read_sql_query('select * from visit_date_2', conn)


#sql_1 = 'update visit_date set '
#sql_2 = 'id=?,No=?,お名前=?,コース名=?,契約日=?,有効期限=?,単価（＠）=?,回目_1=?,回目_2=?,回目_3=?,回目_4=?,回目_5=?,回目_6=?,回目_7=?,回目_8=?,回目_9=?,回目_10=?,回目_11=?,回目_12=?,回目_13=?,回目_14=?,回目_15=?,回目_16=?,回目_17=?,回目_18=?,回目_19=?,回目_20=?,回目_21=?,回目_22=?,回目_23=?,回目_24=?,回目_25=?,回目_26=?,回目_27=?,回目_28=?,回目_29=?,回目_30=?,回目_31=?,回目_32=?,回目_33=?,回目_34=?,回目_35=?,回目_36=?,回目_37=?,回目_38=?,回目_39=?,回目_40=?,回目_41=?,回目_42=?,回目_43=?,回目_44=?,回目_45=?,回目_46=?,回目_47=?,回目_48=?,回目_49=?,回目_50=? where id='
#sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

#c = conn.cursor()

df_contract = pd.read_excel(path, sheet_name = sheet, header = 1)
df_dat = pd.read_excel(path_dat, sheet_name = sheet_dat)

df_contract = df_contract.fillna(0)
df_dat = df_dat.fillna(0)

ary_err = []

for row in df_dat.iterrows():
	i = 0
	num = 0
	col = row[1]
	#print(col[3])
	df_con = df_contract[df_contract['お名前'] == col[3]]  #各名前の行のデータフレームを取得
	##コース名から回数を取得
	course = col[4]
	spl = course.split('　')
	for word in spl :
		if '回' in word:
			word = word[:-1]
			num = int(word)
			break
	##回数の記載がない場合は次のループへ
	if num == 0:
		ary_err.append(col[0])
		continue
	##同じ名前のレコードから該当のレコードを探す
	for row_con in df_con.iterrows():
		col_con = row_con[1]
		##各回数を足す。足せない場合は次のループへ
		try :
			count = col_con[31] + col_con[34] + col_con[33] + col_con[34]
		except Exception as e:
			ary_err.append(col[0])
			continue
		#print(count)
		if num == count :
			print('true')
			i += 1
		else :
			print('false')
	#print(i, col[3])
	if i > 1 :
		ary_err.append(col[0])
	else :
		ind = df_dat.index[df_dat['rec_id'] == col[0]]
		#print(col_con[3])
		#df_dat['契約日'].loc[ind] = utils.datetime.from_excel(col_con[3])
		df_dat['契約日'].loc[ind] = col_con[3]

df_dat.to_excel('S:\個人作業用\那須\ワールドジャパン\有効データ.xlsx')


print(df_dat)
print(ary_err)

'''
	
		df_con = df_con.reset_index()
		df_con = df_con.drop('index',axis = 1)
		print(df_con)
		#print(df['契約日'].loc[0])
		if len(df_con) < 2 and df['契約日'].loc[0] != '0':
			contract_date = df_con['契約日'].loc[0]
			#print(contract_date)
			if type(contract_date) == int or type(contract_date) == float:
				contract_date = utils.datetime.from_excel(contract_date)
			#print(contract_date)
			year = str(contract_date)[0:4]
			print(year)
			for i in range(7,57):
				if len(row[1][i]) > 3:
					if(row[1][i][5:7] < row[1][i - 1][5:7]):
						year = int(year) + 1
					#print(2)
					contract_date = row[1][i]
					mon_day = contract_date[4:]
					row[1][i] = str(year) + mon_day
					print(row[1][i])
			contract_date = df_con['契約日'].loc[0]
			if type(contract_date) == int or type(contract_date) == float:
				contract_date = utils.datetime.from_excel(contract_date)
			row[1][4] = contract_date
		tup = tuple(row[1].T.values.tolist())
		#sql_3 = str(row[1][0])
		#c.execute(sql_1 + sql_2 + sql_3, tup)
'''



#df.to_excel(path, sheet_name = sheet)

#conn.commit()
#conn.close()
