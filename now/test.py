# -*- coding: utf-8 -*-

###記入はしてあるが、名前などの基本情報が別シートにあるファイルをチェック、ファイル書き出し###

import pandas as pd
import os
import openpyxl as excel

pd.set_option('display.max_rows', None)

#dareに学習用データの格納ディレクトリを入れることでどのディレクトリのデータの変更リストかわかりやすくなります
dare = '2フォーマット 2'
base_dir = 'S:/個人作業用/アルバイト/ワールドジャパン/学習用データ/' + dare + '/【A】'

#データフレームの読み込みに用いる列
use_cols = [1,2,3]

#ディレクトリ内のファイルを取得
files = os.listdir(base_dir)

i = 2

ary_name = ['【ボディ】','【バスト】','【フェイシャル】','【脱毛】']
ary = [[] for n in range(4)]	#後から用いる空配列を定義

name,ext = 'uno','katsumi'

print(name)
print(name,ext)