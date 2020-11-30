name = input('ティッカーシンボルを入力: ')


import openpyxl

# バランスシート
out_bsf_workbook = openpyxl.Workbook()
out_bsf_sheet = out_bsf_workbook.active

in_bsf_workbook = openpyxl.load_workbook('../2_Format/BalanceSheet_Format.xlsx', data_only=True)
in_bsf_sheet = in_bsf_workbook['BS']

for i in range(1, 7):
    for j in range(1, 100):
        copy = in_bsf_sheet.cell(row = i, column =j+2).value
        out_bsf_sheet.cell(row = i, column =j , value = copy)

out_bsf_filename = 'bsf_new_result.xlsx'
out_bsf_workbook.save(out_bsf_filename)

# インカムステートメント
out_is_workbook = openpyxl.Workbook()
out_is_sheet = out_is_workbook.active

in_is_workbook = openpyxl.load_workbook('../2_Format/IncomeStatement_Format.xlsx', data_only=True)
in_is_sheet = in_is_workbook['IS']

for i in range(1, 8):
    for j in range(1, 100):
        copy = in_is_sheet.cell(row = i, column =j+2).value
        out_is_sheet.cell(row = i, column =j , value = copy)

out_is_filename = 'is_new_result.xlsx'
out_is_workbook.save(out_is_filename)


import pandas as pd

# 両者をマージ
bsf_result = pd.read_excel(out_bsf_filename)
is_result = pd.read_excel(out_is_filename)

name_result = pd.merge(bsf_result, is_result, on='Breakdown')

name_result_pass = '../1_Storage/{0}_new_join.xlsx'.format(name)
name_result.to_excel(name_result_pass, index=False)


import os

# 不要なファイルを削除
os.remove(out_bsf_filename)
os.remove(out_is_filename)