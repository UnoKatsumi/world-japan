import pandas as pd

# excelファイルを開く
input_book = pd.ExcelFile('S:\個人作業用\渡邊\ワールドジャパン\classify_facial_custom.xlsx')
input_sheet_name = input_book.sheet_names
input_sheet_df = input_book.parse(input_sheet_name[0])

# 相関分析
df_corr = input_sheet_df.corr()

# 結果出力
print(df_corr)