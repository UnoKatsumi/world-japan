# -*- coding: utf-8 -*-

###記入はしてあるが、名前などの基本情報が別シートにあるファイルの名前補填###
####(脱毛は大変なので手動でしてくださいすいません)####

import pandas as pd
import openpyxl as excel
#import sqlite3

#データフレームをprintした時にすべての行を表示するように設定
pd.set_option('display.max_rows', None)

use_cols = [1, 2, 3]
use_cols_e3 = [2, 3, 4]

##cehck_name.pyで出力したテキストファイルから名前補填が必用なファイルを読み出す
#dare = '2フォーマット(柏葉さん)'
dare = '佐藤さん'
f = open('S:/個人作業用/那須/ワールドジャパン/変更_' + dare + '.txt','r', encoding = 'UTF-8')
excel_list = f.readlines()
f.close()


ary_name = ['【ボディ】','【バスト】','【フェイシャル】','【脱毛】']

#転記先保存用
i = 1
n = 0
for path in excel_list :
	path = path.rstrip("\n")
	if path in ary_name:
		i += 1
		print(ary_name[i - 2])
		continue

	#print('################' + name + '################')

	#ブックを開く
	bk = excel.load_workbook(path)

	j = 1	#転記元保存用
	f = 0	#e3かどうか
	flug = 0

	#どのシートに名前などが記載されているか調べる
	for sheet_name in ary_name :
		j += 1

		#シートを開く
		st = bk.worksheets[j]

		#d3,e3(名前)が空欄でないシートをpadasで開き、要素が3列ある場合、そのシートを転記元とする
		d3 = st["d3"].value
		e3 = st["e3"].value
		if d3 != None :
			df = pd.read_excel(path, sheet_name = j, usecols = use_cols)
			if len(df.columns) == 3 :
				print(path, ary_name[i-2], ary_name[j-2])
				flug = 1
				break
		else :
			if e3 != None :
				df = pd.read_excel(path, sheet_name = j, usecols = use_cols_e3)
				if len(df.columns) == 3 :
					print(path, ary_name[i-2], ary_name[j-2])
					flug = 1
					f = 1
					break
	if flug != 1:
		print(path, ary_name[i - 2], "Not found")

	st_data = []
	st2_data = []

	#転記先を開く
	st2 = bk.worksheets[i]

	#名前などの情報を転記
	if f != 1 :
		for m in range(2,18):
			st_data.append(st["D" + str(m)].value)
	else :
		for m in range(2,18):
			st_data.append(st["E" + str(m)].value)
	
	for m in range(2,18):
			st2["D" + str(m)].value = st_data[m-2]
	
	#上書き保存
	bk.save(path)
	bk.close()









