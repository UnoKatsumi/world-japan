# -*- coding: utf-8 -*-

import os
import pandas as pd
import openpyxl as excel
from openpyxl import utils
import sqlite3
import datetime

#ベースディレクトリ
base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/慶野さん/【A】'

#用いるシート
sheet = '【アンケートボディ】'
sheet_spe = ' 【アンケートボディ】'	#なぜか時どきシート名の先頭にスペースが入っていることがあるためそれに対応 (このスクリプトの後に別のスクリプトを書いてて気づいたんですが普通にシート番号の指定でいけるので本当に無駄です、修正できずすいません)
use_cols = [1,2,3]

#アップロードに用いるクエリ
sql_1 = 'insert into questionnaire_body_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, クーポン, 希望コース, 身長, 体重, 目標体重, いつまでに, エステ体験, 施術内容, 結果, 通院回数, 美容に使う金額, 健康状態, 体調, 不調, 治療中・その他, 体質, 体質_その他, アレルギー, アレルギー_有, アレルギー_その他, かぶれ, かぶれ_有, 生理, 周期, 周期_不順, 周期_その他, 常用薬品, 常用薬品_有, 常用薬品_その他, 睡眠, 平均睡眠時間, 性格, 運動, 食事, 嗜好品, 食品嗜好, 体型・気になる部分, 肌・気になる部分, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

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

		#質問内容と回答の列のみ抽出、nanの含まれる行を削除
		df2 = df[['detail','answer']]
		df3 = df2.dropna()

		#アンケート項目の結合に用いる文字列を定義
		ans_cond = ''
		ans_cond_oth = ''
		ans_cons = ''
		ans_cons_oth = ''
		ans_algy = ''
		ans_algy_oth = ''
		ans_cycl = ''
		ans_cycl_oth = ''
		ans_medi = ''
		ans_medi_oth = ''
		ans_pers = ''
		ans_luxu = ''
		ans_meal = ''
		ans_shap = ''
		ans_shap_oth = ''
		ans_skin = ''
		ans_skin_oth = ''

		#各シートの行ごとに処理
		for i in df3['detail'].index:
			if(len(df) == 129):
				#体調の項目をDBに格納できるように結合、以下も同様に結合
				if i >= 29 and i <= 38 :
					ans_cond = df3['detail'].loc[i] + ',' + ans_cond
				elif i == 39:
					ans_cond_oth = df3['answer'].loc[i]
				#体質
				elif i >= 41 and i <= 54 :
					ans_cons = df3['detail'].loc[i] + ',' + ans_cons
				elif i == 55:
					ans_cons_oth = df3['answer'].loc[i]
				#アレルギー
				elif i >= 58 and i <= 62 :
					ans_algy = df3['detail'].loc[i] + ',' + ans_algy
				elif i == 63:
					ans_algy_oth = df3['answer'].loc[i]
				#周期
				elif i == 69 :
					ans_cycl = df3['detail'].loc[i]
				elif i == 39:
					ans_cycl_oth = df3['answer'].loc[i]
				#常用薬品
				elif i >= 73 and i <= 76 :
					ans_medi = df3['detail'].loc[i] + ',' + ans_medi
				elif i == 77:
					ans_medi_oth = df3['answer'].loc[i]
				#性格
				elif i >= 81 and i <= 85 :
					ans_pers = df3['detail'].loc[i] + ',' + ans_pers
				#嗜好品
				elif i >= 88 and i <= 97 :
					ans_luxu = df3['detail'].loc[i] + str(df2['answer'].loc[i + 1]) + ',' + ans_luxu
				#食品嗜好
				elif i >= 98 and i <= 103 :
					ans_meal = df3['detail'].loc[i] + str(df2['answer'].loc[i]) + str(df2['answer'].loc[i + 1]) + ',' + ans_meal
				#体型・気になる部分
				elif i >= 104 and i <= 113 :
					ans_shap = df3['detail'].loc[i] + ',' + ans_shap
				elif i == 114:
					ans_shap_oth = df3['answer'].loc[i]
				#肌・気になる部分
				elif i >= 115 and i <= 124 :
					ans_skin = df3['detail'].loc[i] + ',' + ans_skin
				elif i == 125:
					ans_skin_oth = df3['answer'].loc[i]
			elif(len(df) == 128) :
				if i >= 29 and i <= 38 :
					ans_cond = df3['detail'].loc[i] + ',' + ans_cond
				elif i == 39:
					ans_cond_oth = df3['answer'].loc[i]
				#体質
				elif i >= 41 and i <= 54 :
					ans_cons = df3['detail'].loc[i] + ',' + ans_cons
				elif i == 55:
					ans_cons_oth = df3['answer'].loc[i]
				#アレルギー
				elif i >= 58 and i <= 62 :
					ans_algy = df3['detail'].loc[i] + ',' + ans_algy
				elif i == 63:
					ans_algy_oth = df3['answer'].loc[i]
				#周期
				elif i == 69 :
					ans_cycl = df3['detail'].loc[i]
				elif i == 39:
					ans_cycl_oth = df3['answer'].loc[i]
				#常用薬品
				elif i >= 73 and i <= 76 :
					ans_medi = df3['detail'].loc[i] + ',' + ans_medi
				elif i == 77:
					ans_medi_oth = df3['answer'].loc[i]
				#性格
				elif i >= 81 and i <= 84 :
					ans_pers = df3['detail'].loc[i] + ',' + ans_pers
				#嗜好品
				elif i >= 87 and i <= 96 :
					ans_luxu = df3['detail'].loc[i] + str(df2['answer'].loc[i + 1]) + ',' + ans_luxu
				#食品嗜好
				elif i >= 97 and i <= 102 :
					ans_meal = df3['detail'].loc[i] + str(df2['answer'].loc[i]) + str(df2['answer'].loc[i + 1]) + ',' + ans_meal
				#体型・気になる部分
				elif i >= 103 and i <= 112 :
					ans_shap = df3['detail'].loc[i] + ',' + ans_shap
				elif i == 113:
					ans_shap_oth = df3['answer'].loc[i]
				#肌・気になる部分
				elif i >= 114 and i <= 123 :
					ans_skin = df3['detail'].loc[i] + ',' + ans_skin
				elif i == 124:
					ans_skin_oth = df3['answer'].loc[i]


		ans_cond = ans_cond + ans_cond_oth
		ans_cons = ans_cons + ans_cons_oth
		ans_algy = ans_algy + ans_algy_oth
		ans_cycl = ans_cycl + ans_cycl_oth
		ans_medi = ans_medi + ans_medi_oth
		ans_shap = ans_shap + ans_shap_oth
		ans_skin = ans_skin + ans_skin_oth

		ans_cond = ans_cond.rstrip(',')
		ans_cons = ans_cons.rstrip(',')
		ans_algy = ans_algy.rstrip(',')
		ans_cycl = ans_cycl.rstrip(',')
		ans_medi = ans_medi.rstrip(',')
		ans_luxu = ans_luxu.rstrip(',')
		ans_meal = ans_meal.rstrip(',')
		ans_pers = ans_pers.rstrip(',')
		ans_shap = ans_shap.rstrip(',')
		ans_skin = ans_skin.rstrip(',')

		#結合によって必要なくなった行を削除
		if(len(df) == 129) :
			drop_list = list(range(30,40)) + list(range(42,56)) + list(range(59,64)) + [70] + list(range(74,78)) + list(range(82,86)) + list(range(89,98)) + list(range(99,104)) + list(range(105,115)) + list(range(116,126))
		elif(len(df) == 128) :
			drop_list = list(range(30,40)) + list(range(42,56)) + list(range(59,64)) + [70] + list(range(74,78)) + list(range(82,85)) + list(range(88,97)) + list(range(98,103)) + list(range(104,114)) + list(range(115,125))

		df4 = df.drop(drop_list)

		#来店年月日の記入形式が正しくない場合に修正する
		d2 = st["d2"].value

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

		#アップロードするために結合した項目を再格納
		if(len(df) == 128) :
			df4['answer'][0] = d2
			df4['answer'][29] = ans_cond
			df4['answer'][41] = ans_cons
			df4['answer'][58] = ans_algy
			df4['answer'][69] = ans_cycl
			df4['answer'][73] = ans_medi
			df4['answer'][29] = ans_cond
			df4['answer'][87] = ans_luxu
			df4['answer'][97] = ans_meal
			df4['answer'][103] = ans_shap
			df4['answer'][114] = ans_skin
		elif(len(df) == 129) :
			df4['answer'][0] = d2
			df4['answer'][29] = ans_cond
			df4['answer'][41] = ans_cons
			df4['answer'][58] = ans_algy
			df4['answer'][69] = ans_cycl
			df4['answer'][73] = ans_medi
			df4['answer'][29] = ans_cond
			df4['answer'][88] = ans_luxu
			df4['answer'][98] = ans_meal
			df4['answer'][104] = ans_shap
			df4['answer'][115] = ans_skin

		#データフレームをタプルに変換
		df5 = df4['answer'].T.values.tolist()
		df5_tup = tuple(df5)

		#アップロードの実行
		try :
			c.execute(sql_1 + sql_2 + sql_3, df5_tup)
		except Exception as e:
			print('db:')
			print(e)
			continue
		print('######  Completed inserting to DB successfully!  #####')
		

conn.commit()
conn.close()
