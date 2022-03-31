# -*- coding: utf-8 -*-

import os
import pandas as pd
import openpyxl as excel
from openpyxl import utils
import sqlite3
from datetime import datetime, timedelta

#ベースディレクトリ
base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/佐藤さん/【A】'

#用いるシート
sheet = '【アンケート脱毛】'
sheet_spe = ' 【アンケート脱毛】'	#なぜか時どきシート名の先頭にスペースが入っていることがあるためそれに対応 (このスクリプトの後に別のスクリプトを書いてて気づいたんですが普通にシート番号の指定でいけるので本当に無駄です、修正できずすいません)
use_cols = [1,2,3]

#アップロードに用いるクエリ
sql_1 = 'insert into questionnaire_hairremoval_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 既婚・未婚, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 趣味, 知った理由, 知った理由_その他, DM, 電話対応, 第一印象, 脱毛経験, 脱毛経験_サロン名, 脱毛経験_期間, 脱毛経験_時期, 脱毛経験_方法, 脱毛経験_方法_その他, 脱毛経験_部位, 脱毛経験_部位_その他, 脱毛経験_料金, 自己処理部位, 自己処理部位_その他, 自己処理方法, 自己処理方法_その他, 希望脱毛箇所, 希望脱毛箇所_その他, 日焼け, 日焼け_いつ, 日焼け_理由, 期待・要望, 身近でムダ毛, 身近でムダ毛_人数, 興味のあること, 興味のあること_その他, 取り入れてほしいこと, 取り入れてほしいこと_その他, 妊娠の予定, 病気治療, 病気治療_病名, アレルギー, アレルギー_有, アレルギー_その他, ペースメーカー, 薬・サプリメント, お肌のタイプ, お肌のタイプ_その他, 生理, 生理周期, 生理痛, 生理時の服用薬, 生理時の服用薬_有, 最終月経, 授乳中, 医師からの注意) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

conn = sqlite3.connect('S:/個人作業用/那須/ワールドジャパン/sqlite3/salon_A.db')
c = conn.cursor()

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
			st = bk[sheet_spe]
		except Exception as e :
			print('sheet_spe:')
			print(e)
			continue

	#名前のセルを取得
	d3 = st["d3"].value

	#シートが記入済みの場合に処理を続行
	if d3 != None :
		#各シートをデータフレームとして読込み、各列にラベル付け
		if flag == 0:
			df = pd.read_excel(path, sheet_name = sheet, usecols = use_cols)
		elif flag == 1:
			df = pd.read_excel(path, sheet_name = sheet_spe, usecols = use_cols)
		df.columns = ['index', 'detail', 'answer']

		#回答中の・を,に変換
		rep_num = [21, 23, 26, 28, 30, 38, 40, 46, 50]
		for i in rep_num :
			text = st["d" + str(i + 2)].value
			if text != None :
				rep_text = text.replace('・',',')
				if text[-1] == ',' :
					df['answer'][i] = text[:-1]
				else :
					df['answer'][i] = text

		#来店年月日と生年月日の記入形式が正しくない場合に修正する
		d2 = st["d2"].value
		d5 = st["d5"].value
		typ_d2 = type(d2)
		typ_d5 = type(d5)

		if not(typ_d2 is datetime):
			try:
				d2 = datetime(1899, 12, 30) + timedelta(days=d2)
			except Exception as e:
				print('date_d2:')
				print(e)
		if not(typ_d5 is datetime):
			try:
				d5 = datetime(1899, 12, 30) + timedelta(days=d5)
			except Exception as e:
				print('date_d5:')
				print(e)

		#アップロードするために修正した日付を再格納
		df['answer'][0] = d2
		df['answer'][3] = d5

		#データフレームをタプルに変換
		df2 = df['answer'].T.values.tolist()
		df2_tup = tuple(df2)

		#アップロードの実行
		try :
			c.execute(sql_1 + sql_2 + sql_3, df2_tup)
		except Exception as e:
			print('db:')
			print(e)
			continue
		print('######  Completed inserting to DB successfully!  #####')
		

conn.commit()
conn.close()
