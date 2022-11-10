import pandas as pd
import numpy as np

# excelファイルを読み込み
df = pd.read_excel('S:\個人作業用\渡邊\ワールドジャパン\classify_hairremoval.xlsx', sheet_name='有効データ_hairremoval', engine='openpyxl')
path = 'S:\個人作業用\渡邊\ワールドジャパン\classify_hairremoval_custom.xlsx'
sheet = '有効データ_hairremoval'

# 特定のカラムにおける文字列を数値に変換
for index, row in df.iterrows():
  if row[4] == '会社員' :
    df.loc[index, '職業'] = 1
  elif row[4] == '自営業' :
    df.loc[index, '職業'] = 2
  elif row[4] == 'パート・アルバイト' :
    df.loc[index, '職業'] = 3
  elif row[4] == '主婦' :
    df.loc[index, '職業'] = 4
  elif row[4] == '学生' :
      df.loc[index, '職業'] = 5
  elif row[4] == '無職' :
      df.loc[index, '職業'] = 6
  else :
      df.loc[index, '職業'] = 7
  if row[5] == '既婚' :
      df.loc[index, '結婚'] = 1
  elif row[5] == '未婚' :
      df.loc[index, '結婚'] = 0
  if row[6] == 'HP' :
      df.loc[index, '知った理由'] = 1
  elif row[6] == 'WEB' :
      df.loc[index, '知った理由'] = 2
  elif row[6] == 'HPB' :
      df.loc[index, '知った理由'] = 3
  elif row[6] == 'チラシ' :
      df.loc[index, '知った理由'] = 4
  elif row[6] == '紹介' :
      df.loc[index, '知った理由'] = 5
  else :
      df.loc[index, '知った理由'] = 6
  if row[7] == 'ある' :
      df.loc[index, '脱毛経験'] = 1
  elif row[7] == 'ない' :
      df.loc[index, '脱毛経験'] = 0
  cnt = 0
  if type(row[8]) == str :
    if len(row[8]) != 0 :
      if '顔' in row[8] or 'フェイス' in row[8] :
        cnt += 1
      if '眉' in row[8] :
        cnt += 1
      if '背中' in row[8] :
        cnt += 1
      if '胸' in row[8] :
        cnt += 1
      if 'ワキ' in row[8] :
        cnt += 1
      if '腕' in row[8] or '両腕' in row[8] :
        cnt += 1
      if 'ひじ' in row[8] or 'ヒジ' in row[8] :
        cnt += 1
      if '手' in row[8] :
        cnt += 1
      if '指' in row[8] :
        cnt += 1
      if '下腹' in row[8] :
        cnt += 1
      if 'Vライン' in row[8] or 'V' in row[8] or 'ビキニライン' in row[8] :
        cnt += 1
      if '足' in row[8] or '脚' in row[8] or '両足' in row[8] :
        cnt += 1
      if 'すね' in row[8] :
        cnt += 1
      if 'ひざ' in row[8] or 'ヒザ' in row[8] :
        cnt += 1
      if 'ふくらはぎ' in row[8] :
        cnt += 1
      if 'その他' in row[8] :
        cnt += 1
      df.loc[index, '自己処理部位'] = cnt
    else :
      df.loc[index, '自己処理部位'] = 0
  else :
    df.loc[index, '自己処理部位'] = 0
  cnt = 0
  if type(row[9]) == str :
    if len(row[9]) != 0 :
      if '剃る' in row[9] :
        cnt += 1
      if '抜く' in row[9] or '毛抜き' in row[9]:
        cnt += 1
      if 'カミソリ' in row[9] :
        cnt += 1
      if '脱毛器' in row[9] or '電気脱毛器' in row[9] or '電気脱毛機' in row[9] :
        cnt += 1
      if '脱毛クリーム' in row[9] :
        cnt += 1
      if 'シェイバー' in row[9] or '電気シェイバー' in row[9] :
        cnt += 1
      if '脱色' in row[9] :
        cnt += 1
      if 'ジェル' in row[9] :
        cnt += 1
      if 'その他' in row[9] :
        cnt += 1
      df.loc[index, '自己処理方法'] = cnt
    else :
      df.loc[index, '自己処理方法'] = 0
  else :
    df.loc[index, '自己処理方法'] = 0
  cnt = 0
  if type(row[10]) == str :
    if len(row[10]) != 0 :
      if '顔' in row[10] or 'フェイス' in row[10] :
        cnt += 1
      if '眉' in row[10] :
        cnt += 1
      if 'アゴ' in row[10] :
        cnt += 1
      if '鼻下' in row[10] :
        cnt += 1
      if '背中' in row[10] :
        cnt += 1
      if '胸' in row[10] :
        cnt += 1
      if 'ワキ' in row[10] or '脇' in row[10]:
        cnt += 1
      if '腕' in row[10] or '両腕' in row[10] :
        cnt += 1
      if 'ひじ' in row[10] or 'ヒジ' in row[10] :
        cnt += 1
      if '手' in row[10] or '手の甲' in row[10]:
        cnt += 1
      if '指' in row[10] :
        cnt += 1
      if '下腹' in row[10] or 'お腹' in row[10]:
        cnt += 1
      if 'Vライン' in row[10] or 'V' in row[10] or 'ビキニライン' in row[10]:
        cnt += 1
      if '足' in row[10] or '脚' in row[10] or '両足' in row[10] :
        cnt += 1
      if 'すね' in row[10] :
        cnt += 1
      if 'ひざ' in row[10] or 'ヒザ' in row[10] :
        cnt += 1
      if 'ふくらはぎ' in row[10] :
        cnt += 1
      if '全身' in row[10] :
        cnt += 1
      if 'その他' in row[10] :
        cnt += 1
      df.loc[index, '希望脱毛箇所'] = cnt
    else :
      df.loc[index, '希望脱毛箇所'] = 0
  else :
    df.loc[index, '希望脱毛箇所'] = 0

df = df.set_index('氏名')

# excelファイルとして出力
with pd.ExcelWriter(path) as writer :
	df.to_excel(writer, sheet_name = sheet)
