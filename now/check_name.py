# -*- coding: utf-8 -*-

###記入はしてあるが、名前などの基本情報が別シートにあるファイルをチェック、ファイル書き出し###

import pandas as pd
import os
import openpyxl as excel

pd.set_option('display.max_rows', None)

#dareに学習用データの格納ディレクトリを入れることでどのディレクトリのデータの変更リストかわかりやすくなります
dare = '2フォーマット(柏葉さん)'
base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/' + dare + '/【A】'

#データフレームの読み込みに用いる列
use_cols = [1,2,3]

#ディレクトリ内のファイルを取得
files = os.listdir(base_dir)

i = 2

ary_name = ['【ボディ】','【バスト】','【フェイシャル】','【脱毛】']
ary = [[] for n in range(4)]	#後から用いる空配列を定義

#シートごとに各ファイルを処理
for name in ary_name :
	print('################' + name + '################')

	for path in files :
		#なぜか取得したファイル名の最初に~$が入っていることがあるため、それを除去
		#取得したファイル名とベースディレクトリを結合
		if '~$' in path :
			path = base_dir + '/' + path[2:]
			print(path)
		else :
			path = base_dir + '/' + path
			print(path)
		
		#ファイルの拡張子をチェック
		name,ext = os.path.splitext(path)
		if ext != '.xlsx' :
			continue

		#ファイルの対象シートを開き、名前のセル(d3)の値を取得
		bk = excel.load_workbook(path)
		st = bk.worksheets[i]
		d3 = st["d3"].value

		#d3がnoneの場合、そのシートは未記入であると判定するが、それよりも下の項目に入力がないか調べる
		if d3 == None :
			df = pd.read_excel(path, sheet_name = i, usecols = use_cols)
			#d3がnoneにもかかわらずそのほかの項目に記入がある場合は、変更リストに追加
			if len(df.columns) == 3 :
				print('################### find ###################')
				ary[i-2].append(path)
	i += 1

i = 0

#各保存ディレクトリごとのテキストファイルを開く
f = open('S:\個人作業用\宇野\ワールドジャパン\変更_' + dare + '.txt','w',encoding = 'UTF-8')

#変更リストにファイルリストを出力
for ar in ary :
	f.write(ary_name[i] + '\n')
	i += 1
	for line in ar:
		f.write(line + '\n')
f.close()




