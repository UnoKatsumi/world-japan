import os
import pandas as pd
import sqlite3

#pd.set_option('display.max_rows', None)

path = 'S:\個人作業用\那須\ワールドジャパン\有効データ.xlsx'
path_cnt = 'S:\個人作業用\那須\ワールドジャパン\有効データ候補_回数.xlsx'
path_dat = 'S:\個人作業用\那須\ワールドジャパン\有効データ候補_契約日.xlsx'
sheet = 'Sheet1'


conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')

df = pd.read_sql_query('select * from visit_date', conn)

i = 0
f = 0
ary_i = []
ary_num = []
ary_cnt = []
ary_dat = []

for course in df['コース名']:
    if course != 0:
        if course != '0':
            print(course)
            spl = course.split('　')
            if df['契約日'].loc[i] != '0' :
                for word in spl :
                    if '回' in word:
                        word = word[:-1]
                        #print(word ,i)
                        num = int(word)
                        ary_num.append(num)
                        ary_i.append(i)
                        f = 1
                if f == 0 :
                    ary_cnt.append(i)
            else :
                ary_dat.append(i)
    i += 1
    f = 0

df_cnt = df.loc[ary_cnt]
df_dat = df.loc[ary_dat]

df2 = df.loc[ary_i]

#print(df2)

ary_2 = []

for rec in ary_i:
    num = 0
    for i in range(1,51):
        #print(df.loc[rec]['回目_' + str(i)])
        if df.loc[rec]['回目_' + str(i)] != '0.0' :
            if df.loc[rec]['回目_' + str(i)] != '0' :
                num = i
    #print(num)
    #print(df.loc[rec])
    ary_2.append(num)

#print(len(ary_2))

df2['回数_予定'] = ary_num
df2['回数_来店'] = ary_2

df3 = df2[df2['回数_予定'] <= df2['回数_来店']]
print(df3)

#i = 0

'''
for row in df2.iterrows():
    #print(row)
    if int(row['回数']) == ayr_num[i]:
        print(row['回数'],ayr_num[i])

for name in df2['お名前']:
    
    print(name,ary_num[i], ary_2[i])
    i += 1
'''

df3.to_excel(path, sheet_name = sheet)
df_cnt.to_excel(path_cnt, sheet_name = sheet)
df_dat.to_excel(path_dat, sheet_name = sheet)

conn.close()