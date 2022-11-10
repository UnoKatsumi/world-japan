# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')


# 指定したカラムに値が入力されている行数をカウントし出力
def sql_count(col_name):
  sql = "select count(" + col_name + ") from questionnaire_body_4"
  df = pd.read_sql(sql, conn)
  print(df)


# 関数sql_countを呼び出し
sql_count("id")
sql_count("来店年月日")
sql_count("氏名")
sql_count("フリガナ")
sql_count("生年月日")
sql_count("年齢")
sql_count("郵便番号")
sql_count("住所")
sql_count("TEL_自宅")
sql_count("TEL_携帯")
sql_count("メールアドレス")
sql_count("職業")
sql_count("血液型")
sql_count("結婚")
sql_count("家族構成")
sql_count("DM")
sql_count("知った理由")
sql_count("クーポン")
sql_count("希望コース")
sql_count("身長")
sql_count("体重")
sql_count("目標体重")
sql_count("いつまでに")
sql_count("エステ体験")
sql_count("施術内容")
sql_count("結果")
sql_count("通院回数")
sql_count("美容に使う金額")
sql_count("健康状態")
sql_count("体調")
sql_count("不調")
sql_count("治療中・その他")
sql_count("体質")
sql_count("体質_その他")
sql_count("アレルギー")
sql_count("アレルギー_有")
sql_count("アレルギー_その他")
sql_count("かぶれ")
sql_count("かぶれ_有")
sql_count("生理")
sql_count("周期")
sql_count("周期_不順")
sql_count("周期_その他")
sql_count("常用薬品")
sql_count("常用薬品_有")
sql_count("常用薬品_その他")
sql_count("睡眠")
sql_count("平均睡眠時間")
sql_count("性格")
sql_count("運動")
sql_count("食事")
sql_count("嗜好品")
sql_count("食品嗜好")
sql_count("体型・気になる部分")
sql_count("肌・気になる部分")
sql_count("AYAを選んだ理由")
sql_count("契約")
sql_count("契約内容")

# データベースとの接続を切断
conn.close()
