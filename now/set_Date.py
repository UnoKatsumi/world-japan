# -*- coding: utf-8 -*-
import sqlite3
import pandas as pd
import datetime
from openpyxl import utils

conn = sqlite3.connect('S:/個人作業用/那須/ワールドジャパン/sqlite3/salon_A.db')
c = conn.cursor()

############################   body   ############################
'''
df = pd.read_sql_query('select * from questionnaire_body_3', conn)
df = df.drop('id',axis = 1)
df = df.drop_duplicates()

c.execute('delete from questionnaire_body_3')
#conn.commit()


sql_1 = 'insert into questionnaire_body_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, クーポン, 希望コース, 身長, 体重, 目標体重, いつまでに, エステ体験, 施術内容, 結果, 通院回数, 美容に使う金額, 健康状態, 体調, 不調, 治療中・その他, 体質, 体質_その他, アレルギー, アレルギー_有, アレルギー_その他, かぶれ, かぶれ_有, 生理, 周期, 周期_不順, 周期_その他, 常用薬品, 常用薬品_有, 常用薬品_その他, 睡眠, 平均睡眠時間, 性格, 運動, 食事, 嗜好品, 食品嗜好, 体型・気になる部分, 肌・気になる部分, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

##来店年月日修正
for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[0]
    try :
        date_list = date.split('-')
    except Exception as e :
        continue
    year = int(date_list[0])
    if year > 2021 :
        year = year - 30
        date = str(year) + date[4:]
        tup_list = list(tup)
        tup_list[0] = date
        tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)

##生年月日修正_諸々修正
for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[3]
    if(type(date) == str) :
        if ('.' in date) :
            date_list = date.split('.')
            date = datetime.datetime(int(date_list[0]) - 12 + 2000, int(date_list[1]), int(date_list[2]))
            tup_list = list(tup)
            tup_list[3] = date
            tup = tuple(tup_list)
        elif (len(date) == 5) :
            date = utils.datetime.from_excel(int(date))
            tup_list = list(tup)
            tup_list[3] = date
            tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)


##生年月日修正_減算
for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[3]
    if(type(date) == str) :
        if('-' in date):
            date_list = date.split('-')
            year = int(date_list[0])
            if(year > 2021) :
                year = year - 63
                date = str(year) + date[4:]
                tup_list = list(tup)
                tup_list[3] = date
                tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)
'''
'''
############################   bust   ############################
df = pd.read_sql_query('select * from questionnaire_bust_3', conn)
df = df.drop('id',axis = 1)
df = df.drop_duplicates()

c.execute('delete from questionnaire_bust_3')
#conn.commit()

sql_1 = 'insert into questionnaire_bust_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, 来店のきっかけ, アレルギー, アレルギー_有, 治療, 治療_有, 服用薬, 服用薬_有, 健康状態, 健康状態_悪い, 精神状態, 精神状態_悪い, 月経, 周期, ダイエット経験, かぶれの経験, かぶれ_有, バストの悩み, バストの悩み_その他, 気になった時期, 気になった理由, 改善策, 改善策_有, 睡眠時間, 平均睡眠時間, 睡眠型, 喫煙, 喫煙_本数, 出産経験, 授乳経験, いつまでに, 美しくなりたい理由, どんな状態に, どの程度の施術_コース, どの程度の施術_予算, 予算金額, バストケア経験, バストケア経験_サロン, バストケア経験_費用, バストケア経験_期間, バストケア経験_結果, エステサロン経験, エステサロン経験_コース, エステサロン経験_サロン, エステサロン経験_費用, エステサロン経験_期間, エステサロン経験_結果, 該当項目,カウンセリング, カウンセリング_その他, 他のコース, 他のコース_その他, 現在のサイズ, 理想のサイズ, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

##来店年月日修正
for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[3]
    try :
        date_list = date.split('-')
    except Exception as e :
        continue
    year = int(date_list[0])
    if year > 2021 :
        year = year - 30
        date = str(year) + date[4:]
        tup_list = list(tup)
        tup_list[3] = date
        tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)

'''
'''
###facial
df = pd.read_sql_query('select * from questionnaire_facial_3', conn)
df = df.drop('id',axis = 1)
df = df.drop_duplicates()

c.execute('delete from questionnaire_facial_3')
#conn.commit()

sql_1 = 'insert into questionnaire_facial_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 年齢, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 血液型, 結婚, 家族構成, DM, 知った理由, 来店のきっかけ, クーポン, 来店の目的, 専門の相談, 専門の相談_有, 専門の相談_感想, 希望, 希望時期, 質問, 肌気になるところ, 手入れ_朝, 手入れ_朝_その他, 手入れ_夜, 手入れ_夜_その他, 化粧品メーカー, 化粧品メーカー_その他, 化粧品_結果, 手入れ_感想, 美容代, エステサロン経験, エステサロン経験_コース, エステサロン経験_サロン, エステサロン経験_費用, エステサロン経験_期間, エステサロン経験_結果, AYAを選んだ理由, 契約, 契約内容) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[0]
    try :
        date_list = date.split('-')
    except Exception as e :
        continue
    year = int(date_list[0])
    if year > 2021 :
        year = year - 30
        date = str(year) + date[4:]
        tup_list = list(tup)
        tup_list[0] = date
        tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)

'''

###hairremoval
df = pd.read_sql_query('select * from questionnaire_hairremoval_3', conn)
df = df.drop('id',axis = 1)
df = df.drop_duplicates()

c.execute('delete from questionnaire_hairremoval_3')
#conn.commit()

sql_1 = 'insert into questionnaire_hairremoval_3 '
sql_2 = '(来店年月日, 氏名, フリガナ, 生年月日, 既婚・未婚, 郵便番号, 住所, TEL_自宅, TEL_携帯, メールアドレス, 職業, 趣味, 知った理由, 知った理由_その他, DM, 電話対応, 第一印象, 脱毛経験, 脱毛経験_サロン名, 脱毛経験_期間, 脱毛経験_時期, 脱毛経験_方法, 脱毛経験_方法_その他, 脱毛経験_部位, 脱毛経験_部位_その他, 脱毛経験_料金, 自己処理部位, 自己処理部位_その他, 自己処理方法, 自己処理方法_その他, 希望脱毛箇所, 希望脱毛箇所_その他, 日焼け, 日焼け_いつ, 日焼け_理由, 期待・要望, 身近でムダ毛, 身近でムダ毛_人数, 興味のあること, 興味のあること_その他, 取り入れてほしいこと, 取り入れてほしいこと_その他, 妊娠の予定, 病気治療, 病気治療_病名, アレルギー, アレルギー_有, アレルギー_その他, ペースメーカー, 薬・サプリメント, お肌のタイプ, お肌のタイプ_その他, 生理, 生理周期, 生理痛, 生理時の服用薬, 生理時の服用薬_有, 最終月経, 授乳中, 医師からの注意) '
sql_3 = 'values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'

##生年月日修正_諸々修正
for row in df.iterrows():
    tup = tuple(row[1])
    date = tup[0]
    if(type(date) == str) :
        if ('.' in date) :
            date_list = date.split('.')
            date = datetime.datetime(int(date_list[0]) - 12 + 2000, int(date_list[1]), int(date_list[2]))
            tup_list = list(tup)
            tup_list[0] = date
            tup = tuple(tup_list)
    c.execute(sql_1 + sql_2 + sql_3, tup)


conn.commit()
conn.close()