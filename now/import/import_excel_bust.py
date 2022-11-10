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
sheet = '【アンケートバスト】'
sheet_spe = ' 【アンケートバスト】'	#なぜか時どきシート名の先頭にスペースが入っていることがあるためそれに対応 (このスクリプトの後に別のスクリプトを書いてて気づいたんですが普通にシート番号の指定でいけるので本当に無駄です、修正できずすいません)
use_cols = [1,2,3]

#アップロードに用いるクエリ
sql_1 = 'insert into questionnaire_bust_4 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, 来店のきっかけ, アレルギー, アレルギー_有, 治療, 治療_有, 服用薬, 服用薬_有, 健康状態, 健康状態_悪い, 精神状態, 精神状態_悪い, 月経, 周期, ダイエット経験, かぶれの経験, かぶれ_有, バストの悩み, バストの悩み_その他, 気になった時期, 気になった理由, 改善策, 改善策_有, 睡眠時間, 平均睡眠時間, 睡眠型, 喫煙, 喫煙_本数, 出産経験, 授乳経験, いつまでに, 美しくなりたい理由, どんな状態に, どの程度の施術_コース, どの程度の施術_予算, 予算金額, バストケア経験, バストケア経験_サロン, バストケア経験_費用, バストケア経験_期間, バストケア経験_結果, エステサロン経験, エステサロン経験_コース, エステサロン経験_サロン, エステサロン経験_費用, エステサロン経験_期間, エステサロン経験_結果, 該当項目,カウンセリング, カウンセリング_その他, 他のコース, 他のコース_その他, 現在のサイズ, 理想のサイズ, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')
c = conn.cursor()

#ディレクトリ内のファイルを取得
files = os.listdir(base_dir)

#各ファイルごとに処理
for path in files :
	#ファイル名や拡張子のチェック
	if '~$' in path :
		path = base_dir + '/' + path[2:]
		# print(path)
	else :
		path = base_dir + '/' + path
		# print(path)
	
	name,ext = os.path.splitext(path)
	if ext != '.xlsx' :
		continue

	#ブックを開く
	bk = excel.load_workbook(path)

	flag = 0

	#シート名の先頭に空欄があるかによって場合分け
	try :
		st = bk[sheet_spe]	
	except Exception as e :
		# print('sheet_spe:')
		# print(e)
		flag = 1
	if flag == 1 :
		try :
			st = bk[sheet]
		except Exception as e :
			# print('sheet:')
			# print(e)
			continue

	#名前のセルを取得
	d3 = st["d3"].value

	#シートが記入済みの場合に処理を続行
	if d3 != None :
		#各シートをデータフレームとして読込み、各列にラベル付け
		if flag == 0:
			df = pd.read_excel(path, sheet_name = sheet_spe, usecols = use_cols)
		elif flag == 1:
			df = pd.read_excel(path, sheet_name = sheet, usecols = use_cols)
		df.columns = ['index', 'detail', 'answer']

		#質問内容と回答の列のみ抽出、nanの含まれる行を削除
		df2 = df[['detail','answer']]
		df3 = df2.dropna()

		#アンケート項目の結合に用いる文字列を定義
		ans_bust = ''
		ans_appl = ''
		ans_coun = ''
		ans_cose = ''

		#各シートの行ごとに処理
		for i in df3['detail'].index:
			#バストの悩みの項目をDBに格納できるように結合、以下も同様に結合
			if i >= 32 and i <= 37 :
				ans_bust = df3['detail'].loc[i] + ',' + ans_bust
			#該当項目
			elif i >= 67 and i <= 73 :
				ans_appl = df3['detail'].loc[i] + ',' + ans_appl
			#カウンセリング
			elif i >= 74 and i <= 77 :
				ans_coun = df3['detail'].loc[i] + ',' + ans_coun
			#コース
			elif i >= 79 and i <= 86 :
				ans_cose = df3['detail'].loc[i] + ',' + ans_cose

		ans_bust = ans_bust.rstrip(',')
		ans_coun = ans_coun.rstrip(',')
		ans_appl = ans_appl.rstrip(',')
		ans_cose = ans_cose.rstrip(',')

		#結合によって必要なくなった行を削除
		drop_list = list(range(33,38)) + list(range(68,74)) + list(range(75,78)) + list(range(80,87))
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
		df4['answer'][32] = ans_bust
		df4['answer'][67] = ans_appl
		df4['answer'][74] = ans_coun
		df4['answer'][79] = ans_cose

		#データフレームをタプルに変換
		df5 = df4['answer'].T.values.tolist()
		df5_tup = tuple(df5)
		print(df5_tup)

		# #アップロードの実行
		# try :
		# 	c.execute(sql_1 + sql_2 + sql_3, df5_tup)
		# except Exception as e:
		# 	print('db:')
		# 	print(e)
		# 	continue

		print('######  Completed inserting to DB successfully!  #####')

conn.commit()
conn.close()
