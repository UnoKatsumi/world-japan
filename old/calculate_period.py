import pandas as pd
import sqlite3
import datetime

#pd.set_option('display.max_rows', None)

path = 'S:\個人作業用\那須\ワールドジャパン\merge_body.xlsx'
sheet = 'Sheet2'


conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')

df = pd.read_sql_query('select * from visit_date', conn)
#df_period = pd.read_sql_query('select * from visit_date', conn)
#df_period_2 = pd.read_sql_query('select * from visit_date_2', conn)

#print(df_body)
#print(df_body_2)
#print(df)
#df_period = df_period.drop('id', axis = 1)
#df_period_2 = df_period_2.drop('id', axis = 1)
#df = df.drop('id', axis = 1)

#df = pd.merge(df_body, df_body_2, suffixes = ['_1', '_2'], how = 'outer')

merge = []

for row in df.iterrows():
	print(row[1][1])
	cnt = 0
	if len(row[1][2]) != 1:
		#print(1)
		for i in range(6,56):
			if len(row[1][i]) > 3:
					cnt += 1
					print(cnt)
					print(len(row[1][i]))
		if cnt > 0 :
			period = datetime.datetime.strptime(row[1][cnt + 5], '%Y-%m-%d %H:%M:%S') - datetime.datetime.strptime(row[1][6], '%Y-%m-%d %H:%M:%S')
			#print(1)
			merge.append([row[1][1], row[1][2], cnt, period.days])
			print(merge[-1])

print(merge)

#df.to_excel(path, sheet_name = sheet)

conn.close()