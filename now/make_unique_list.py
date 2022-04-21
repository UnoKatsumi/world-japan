###フォーマットが正しくないファイルをリストアップ###

import os
import openpyxl as excel
import shutil
import collections

#ファイル名のみの配列
ary = []
#フルパスの配列
ary2 = []

#学習用データの保存ディレクトリ
dare = ['2フォーマット(潮永さん）', '2フォーマット(渡部さん)', '2フォーマット(柏葉さん)', '慶野さん', '佐藤さん']

for name in dare:
    base_dir = 'S:/個人作業用/宇野/ワールドジャパン/学習用データ(テスト用)/' + name + '/【A】'
    #ディレクトリ内の全てを取得
    files_dirs = os.listdir(base_dir)
    #ファイルのみを取得
    files = [f for f in files_dirs if os.path.isfile(os.path.join(base_dir, f))]
    #Thumbs.dbを削除
    files = [s for s in files if s != 'Thumbs.db']
    #ファイル数を取得
    len_files = str(len(files))
    print(name + ' : ' + str(len_files))
    for path in files:
        if path not in ary:
            ary.append(path)
            ary2.append(base_dir + '/' + path)
print(str(len(ary2)))

for one in ary2 :
    shutil.copy(one, 'S:/個人作業用/宇野/ワールドジャパン/学習用データ(テスト用)/重複なしデータ/【A】/')
