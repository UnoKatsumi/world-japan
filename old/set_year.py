# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3
import datetime

#pd.set_option('display.max_rows', None)

conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')

df = pd.read_sql_query('select * from visit_date', conn)
#df_period = pd.read_sql_query('select * from visit_date', conn)
#df_period_2 = pd.read_sql_query('select * from visit_date_2', conn)


sql_1 = 'update visit_date set '
sql_2 = 'id=?,No=?,お名前=?,コース名=?,契約日=?,有効期限=?,単価（＠）=?,回目_1=?,回目_2=?,回目_3=?,回目_4=?,回目_5=?,回目_6=?,回目_7=?,回目_8=?,回目_9=?,回目_10=?,回目_11=?,回目_12=?,回目_13=?,回目_14=?,回目_15=?,回目_16=?,回目_17=?,回目_18=?,回目_19=?,回目_20=?,回目_21=?,回目_22=?,回目_23=?,回目_24=?,回目_25=?,回目_26=?,回目_27=?,回目_28=?,回目_29=?,回目_30=?,回目_31=?,回目_32=?,回目_33=?,回目_34=?,回目_35=?,回目_36=?,回目_37=?,回目_38=?,回目_39=?,回目_40=?,回目_41=?,回目_42=?,回目_43=?,回目_44=?,回目_45=?,回目_46=?,回目_47=?,回目_48=?,回目_49=?,回目_50=? where id='
#sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

c = conn.cursor()

#print(df_body)
#print(df_body_2)
#print(df)
#df_period = df_period.drop('id', axis = 1)
#df_period_2 = df_period_2.drop('id', axis = 1)
#df = df.drop('id', axis = 1)

#df = pd.merge(df_body, df_body_2, suffixes = ['_1', '_2'], how = 'outer')


#print(df)

for row in df.iterrows():
	print(row[1][2])
	#print(len(row[1][4]) > 3 & len(row[1][3]) > 3)
	if len(row[1][7]) > 3 and len(row[1][4]) > 3:
		#print(1)
		contract_date = row[1][4]
		year = contract_date[0:4]
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
	tup = tuple(row[1].T.values.tolist())
	sql_3 = str(row[1][0])
	c.execute(sql_1 + sql_2 + sql_3, tup)



#df.to_excel(path, sheet_name = sheet)

conn.commit()
conn.close()