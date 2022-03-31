###フォーマットが正しくないファイルをリストアップ###

import os
import openpyxl as excel
import shutil

#学習用データの保存ディレクトリ
base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/2フォーマット 2/【A】'

#ディレクトリ内のファイルとその数を取得
files = os.listdir(base_dir)
len_files = str(len(files))

#プログレスバーの表示に使用
ary = []
j = 0
bar_size = 20

#各ファイルをチェック
for nam in files:
    j = j + 1
    #なぜか取得したファイル名の最初に~$が入っていることがあるため、それを除去
    #取得したファイル名とベースディレクトリを結合
    if '~$' in nam :
        path = base_dir + '/' + nam[2:]
        #print(path)
    else :
        path = base_dir + '/' + nam
        #print(path)

    #ファイルの拡張子をチェック
    name,ext = os.path.splitext(path)
    if ext != '.xlsx' :
        continue
    
    #プログレスバーの表示
    rate = round(j / int(len_files) * bar_size) - 1
    pro_bar =('=' * rate) + (' ' * (bar_size - rate))
    print('\r' + '[' + pro_bar + ']' + ' ' + str(round((j + 1) / int(len_files) * 100, 2)) + '% ' + nam + '             ', end='')

    #ファイルを開く
    bk = excel.load_workbook(path)

    val = 0
    ary = []

    #ファイル内の各シートを読込み
    for sheet in bk :
        i = 1
        if 'ボディ' in sheet.title :
            #項目数をカウント
            while val != '契約内容' :
                val = sheet.cell(row = i, column = 2).value
                i = i + 1
            #正しくない項目数の場合は例外ファイルリストに追加
            if i != 131 :
                ary.append(path)
        elif 'バスト' in sheet.title :
            #項目数をカウント
            while val != '契約内容' :
                val = sheet.cell(row = i, column = 2).value
                i = i + 1
            #正しくない項目数の場合は例外ファイルリストに追加
            if i != 95 :
                ary.append(path)
        elif 'フェイシャル' in sheet.title :
            #項目数をカウント
            while val != '契約内容' :
                val = sheet.cell(row = i, column = 2).value
                i = i + 1
            #正しくない項目数の場合は例外ファイルリストに追加
            if i != 73 :
                ary.append(path)
        elif '脱毛' in sheet.title :
            #項目数をカウント
            while val != '医師から注意を受けている事や体質的に気になる事' :
                val = sheet.cell(row = i, column = 2).value
                i = i + 1
            #正しくない項目数の場合は例外ファイルリストに追加
            if i != 62 :
                ary.append(path)

#プログレスバーの表示
rate = bar_size
pro_bar =('=' * rate) + (' ' * (bar_size - rate))
print('\r' + '[' + pro_bar + ']' + ' ' + str(round(j / int(len_files) * 100, 2)) + '%', end='')

ary = list(set(ary))

#例外ファイルを要修正フォルダに移動させる
for one in ary :
    shutil.move(one, one.replace('/【A】/', '/【A】/【要修正】/'))

#例外ファイルリストをテキストファイルに出力
f = open('例外.txt', 'w', encoding = 'UTF-8')
for one in ary :
    f.write(one[0] + ':' + one[1])
    f.write('\n')
f.close()

print('\nComplete!',end = '')