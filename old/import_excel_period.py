# -*- coding: utf-8 -*-

import os
import pandas as pd
import openpyxl as excel
from openpyxl import utils
import sqlite3
import datetime
import math

#pd.set_option('display.max_rows', None)

base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/2フォーマット/'
#base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/2フォーマット 2/'
file = '【H】来店日リスト.xlsx'

sheet = 'Sheet1'
#use_cols = [1,2,3]

sql_1 = 'insert into visit_date '
sql_2 = '(No,お名前,コース名,契約日,有効期限,単価（＠）,回目_1,回目_2,回目_3,回目_4,回目_5,回目_6,回目_7,回目_8,回目_9,回目_10,回目_11,回目_12,回目_13,回目_14,回目_15,回目_16,回目_17,回目_18,回目_19,回目_20,回目_21,回目_22,回目_23,回目_24,回目_25,回目_26,回目_27,回目_28,回目_29,回目_30,回目_31,回目_32,回目_33,回目_34,回目_35,回目_36,回目_37,回目_38,回目_39,回目_40,回目_41,回目_42,回目_43,回目_44,回目_45,回目_46,回目_47,回目_48,回目_49,回目_50) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

conn = sqlite3.connect('S:/個人作業用/那須/sqlite3/salon_A.db')
c = conn.cursor()

path = base_dir + file

df = pd.read_excel(path, sheet_name = sheet, header = 1) #, usecols = use_cols
df = df.drop_duplicates()

df = df.reset_index()
df = df.drop('index', axis = 1)
df = df.fillna(0)

i = 0

for row in df.iterrows():
	if row[1][2] != 0:
		for i in range(6,56):
			if type(row[1][i]) is int:
				if row[1][i] > 0:
					row[1][i] = utils.datetime.from_excel(row[1][i])
					#print(row[1][i])
			elif type(row[1][i]) is float:
				if row[1][i] > 0:
					row[1][i] = int(row[1][i])
					row[1][i] = utils.datetime.from_excel(row[1][i])
	tup = tuple(row[1].T.values.tolist())

	c.execute(sql_1 + sql_2 + sql_3, tup)


conn.commit()
conn.close()
