# -*- coding: utf-8 -*-

import os
import pandas as pd
import openpyxl as excel
from openpyxl import utils
import sqlite3
import datetime

#ベースディレクトリ
base_dir = 'S:/個人作業用/渡邊/ワールドジャパン/重複なしデータ/【A】'

#用いるシート
sheet = '【アンケートフェイシャル】'
sheet_spe = ' 【アンケートフェイシャル】'	#なぜか時どきシート名の先頭にスペースが入っていることがあるためそれに対応 (このスクリプトの後に別のスクリプトを書いてて気づいたんですが普通にシート番号の指定でいけるので本当に無駄です、修正できずすいません)
use_cols = [1,2,3]

#アップロードに用いるクエリ
sql_1 = 'insert into questionnaire_facial_4 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, 来店のきっかけ, クーポン, 来店の目的, 専門の相談, 専門の相談_有, 専門の相談_感想, 希望, 希望時期, 質問, 肌気になるところ, 手入れ_朝, 手入れ_朝_その他, 手入れ_夜, 手入れ_夜_その他, 化粧品メーカー, 化粧品メーカー_その他, 化粧品_結果, 手入れ_感想, 美容代, エステサロン経験, エステサロン経験_コース, エステサロン経験_サロン, エステサロン経験_費用, エステサロン経験_期間, エステサロン経験_結果, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')
c = conn.cursor()

err_list = []

#ディレクトリ内のファイルを取得
files = os.listdir(base_dir)

#各ファイルごとに処理
for path in files :
	#ファイル名や拡張子のチェック
	if '~$' in path :
		path = base_dir + '/' + path[2:]
		print(path)
	else :
		path = base_dir + '/' + path
		print(path)
	
	name,ext = os.path.splitext(path)
	if ext != '.xlsx' :
		continue

	#ブックを開く
	bk = excel.load_workbook(path)

	flag = 0

	#シート名の先頭に空欄があるかによって場合分け
	try :
		st = bk[sheet]	
	except Exception as e :
		print('sheet:')
		print(e)
		flag = 1
	if flag == 1 :
		try :
			st = bk[sheet]
		except Exception as e :
			print('sheet_spe:')
			print(e)
			continue
	
	#名前のセルとフォーマット違いの際にずれが生じるセルを取得
	d3 = st["d3"].value
	a72 = st["a72"].value

	#シートが記入済みの場合に処理を続行
	if d3 != None :
		#各シートをデータフレームとして読込み、各列にラベル付け
		if flag == 0:
			df = pd.read_excel(path, sheet_name = sheet, usecols = use_cols)
		elif flag == 1:
			df = pd.read_excel(path, sheet_name = sheet_spe, usecols = use_cols)
		df.columns = ['index', 'detail', 'answer']

		#質問内容と回答の列のみ抽出、nanの含まれる行を削除
		df2 = df[['detail','answer']]
		df3 = df2.dropna()

		#アンケート項目の結合に用いる文字列を定義
		ans_skin = ''
		ans_morn = ''
		ans_nigt = ''

		#フォーマット違いのファイルは処理する行を変更
		if a72 != None:
			lim = 42
		else :
			lim = 41
		
		#各シートの行ごとに処理
		for i in df3['detail'].index:
			#肌・気になっているところの項目をDBに格納できるように結合、以下も同様に結合
			if i >= 25 and i <= lim :
				ans_skin = df3['detail'].loc[i] + ',' + ans_skin
			#手入れ方法・朝
			elif i >= 43 and i <= 48 :
				ans_morn = df3['detail'].loc[i] + ',' + ans_morn
			#手入れ方法・夜
			elif i >= 50 and i <= 55 :
				ans_nigt = df3['detail'].loc[i] + ',' + ans_nigt

		ans_skin = ans_skin.rstrip(',')
		ans_nigt = ans_nigt.rstrip(',')
		ans_morn = ans_morn.rstrip(',')

		#結合によって必要なくなった行を削除
		if lim == 42 :
			drop_list = list(range(26,43)) + list(range(44,49)) + list(range(51,56))
		else :
			drop_list = list(range(26,42)) + list(range(43,48)) + list(range(50,55))
		df4 = df.drop(drop_list)

		#来店年月日と生年月日の記入形式が正しくない場合に修正する
		d2 = st["d2"].value
		d5 = st["d5"].value

		if (d2 != None) :
			if ('.' in str(d2)) :
				d2 = d2.split('.')
				for i in range(0,3) :
					d2[i]= int(d2[i])
				if (d2[0] > 28) :
					d2 = datetime.datetime(d2[0] - 12 + 2000, d2[1], d2[2])
				else :
					d2 = datetime.datetime(d2[0] + 18 + 2000, d2[1], d2[2])
			else :
				try :
					d2 = utils.datetime.from_excel(d2)
				except Exception as e:
					print('date:')
					print(e)
		if (d5 != None) :
			if ('.' in str(d5)) :
				d5 = d5.split('.')
				for i in range(0,3) :
					d5[i]= int(d5[i])
				if (d5[0] > 28) :
					d5 = datetime.datetime(d5[0] - 12 + 2000, d5[1], d5[2])
				else :
					d5 = datetime.datetime(d5[0] + 18 + 2000, d5[1], d5[2])
			else :
				try :
					d5 = utils.datetime.from_excel(d5)
				except Exception as e:
					print('date:')
					print(e)

		#アップロードするために結合した項目を再格納
		df4['answer'][0] = d2
		df4['answer'][3] = d5
		if lim == 42 :
			df4['answer'][25] = ans_skin
			df4['answer'][43] = ans_morn
			df4['answer'][50] = ans_nigt
		else :
			df4['answer'][25] = ans_skin
			df4['answer'][42] = ans_morn
			df4['answer'][49] = ans_nigt

		#データフレームをタプルに変換
		df5 = df4['answer'].T.values.tolist()
		df5_tup = tuple(df5)

		#アップロードの実行
		if len(df5_tup) < 44 :
			err_list.append(path)
		try :
			c.execute(sql_1 + sql_2 + sql_3, df5_tup)
		except Exception as e:
			print('db:')
			print(e)
			continue
		print('######  Completed inserting to DB successfully!  #####')
		

conn.commit()
conn.close()

with open('例外.txt', 'w') as f :
	f.write('\n'.join(err_list))