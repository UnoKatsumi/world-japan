# -*- coding: utf-8 -*-

import pandas as pd
import sqlite3

# データベースへ接続
conn = sqlite3.connect('S:/個人作業用/渡邊/ワールドジャパン/sqlite3/salon_A.db')


# 指定したカラムに値が入力されている行数をカウントし出力
def sql_count(col_name):
  sql = "select count(" + col_name + ") from questionnaire_bust_4"
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
sql_count("来店のきっかけ")
sql_count("アレルギー")
sql_count("アレルギー_有")
sql_count("治療")
sql_count("治療_有")
sql_count("服用薬")
sql_count("服用薬_有")
sql_count("健康状態")
sql_count("健康状態_悪い")
sql_count("精神状態")
sql_count("精神状態_悪い")
sql_count("月経")
sql_count("周期")
sql_count("ダイエット経験")
sql_count("かぶれの経験")
sql_count("かぶれ_有")
sql_count("バストの悩み")
sql_count("バストの悩み_その他")
sql_count("気になった時期")
sql_count("気になった理由")
sql_count("改善策")
sql_count("改善策_有")
sql_count("睡眠時間")
sql_count("平均睡眠時間")
sql_count("睡眠型")
sql_count("喫煙")
sql_count("喫煙_本数")
sql_count("出産経験")
sql_count("授乳経験")
sql_count("いつまでに")
sql_count("美しくなりたい理由")
sql_count("どんな状態に")
sql_count("どの程度の施術_コース")
sql_count("どの程度の施術_予算")
sql_count("予算金額")
sql_count("バストケア経験")
sql_count("バストケア経験_サロン")
sql_count("バストケア経験_費用")
sql_count("バストケア経験_期間")
sql_count("バストケア経験_結果")
sql_count("エステサロン経験")
sql_count("エステサロン経験_コース")
sql_count("エステサロン経験_サロン")
sql_count("エステサロン経験_費用")
sql_count("エステサロン経験_期間")
sql_count("エステサロン経験_結果")
sql_count("該当項目")
sql_count("カウンセリング")
sql_count("カウンセリング_その他")
sql_count("他のコース")
sql_count("他のコース_その他")
sql_count("現在のサイズ")
sql_count("理想のサイズ")
sql_count("AYAを選んだ理由")
sql_count("契約")
sql_count("契約内容")

# データベースとの接続を切断
conn.close()
