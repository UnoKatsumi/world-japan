import os
import pandas as pd
import sqlite3

pd.set_option('display.max_rows', None)

path = 'S:\個人作業用\那須\ワールドジャパン\merge.xlsx'
sheet = 'Sheet1'


conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')

df_body = pd.read_sql_query('select * from questionnaire_body', conn)
df_bust = pd.read_sql_query('select * from questionnaire_bust', conn)
df_facial = pd.read_sql_query('select * from questionnaire_facial', conn)
df_hair = pd.read_sql_query('select * from questionnaire_hairremoval', conn)

df_body_2 = pd.read_sql_query('select * from questionnaire_body_2', conn)
df_bust_2 = pd.read_sql_query('select * from questionnaire_bust_2', conn)
df_facial_2 = pd.read_sql_query('select * from questionnaire_facial_2', conn)
df_hair_2 = pd.read_sql_query('select * from questionnaire_hairremoval_2', conn)
#print(df_body_2)

df_body = df_body.drop('id', axis = 1)
df_bust = df_bust.drop('id', axis = 1)
df_facial = df_facial.drop('id', axis = 1)
df_hair = df_hair.drop('id', axis = 1)

df_body_2 = df_body_2.drop('id', axis = 1)
df_bust_2 = df_bust_2.drop('id', axis = 1)
df_facial_2 = df_facial_2.drop('id', axis = 1)
df_hair_2 = df_hair_2.drop('id', axis = 1)

df1 = pd.merge(df_body, df_bust, suffixes = ['_body', '_bust'], how = 'outer')
df2 = pd.merge(df1, df_facial, suffixes = ['', '_facial'], how = 'outer')
df3 = pd.merge(df2, df_hair, suffixes = ['', '_hair'], how = 'outer')

'''
df4 = pd.merge(df_body_2, df_bust_2, on = ['氏名', 'フリガナ'], suffixes = ['_body', '_bust'], how = 'outer')
df5 = pd.merge(df4, df_facial_2, on = ['氏名', 'フリガナ'], suffixes = ['', '_facial'], how = 'outer')
df6 = pd.merge(df5, df_hair_2, on = ['氏名', 'フリガナ'], suffixes = ['', '_hair'], how = 'outer')
'''

#df7 = pd.merge(df3, df6, suffixes = ['_1', '_2'], how = 'outer')

df3.to_excel(path, sheet_name = sheet)

conn.close()