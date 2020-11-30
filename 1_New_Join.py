import openpyxl

# バランスシート

out_bsf_workbook = openpyxl.Workbook()
out_bsf_sheet = out_bsf_workbook.active

in_bsf_workbook = openpyxl.load_workbook('BalanceSheet_Format.xlsx', data_only=True)
in_bsf_sheet = in_bsf_workbook['BS']

for i in range(1, 7):
    for j in range(1, 50):
        copy = in_bsf_sheet.cell(row = i, column =j+2).value
        out_bsf_sheet.cell(row = i, column =j , value = copy)

out_bsf_filename = 'bsf_result.xlsx'
out_bsf_workbook.save(out_bsf_filename)

# インカムステートメント

out_is_workbook = openpyxl.Workbook()
out_is_sheet = out_is_workbook.active

in_is_workbook = openpyxl.load_workbook('IncomeStatement_Format.xlsx', data_only=True)
in_is_sheet = in_is_workbook['IS']

for i in range(1, 8):
    for j in range(1, 50):
        copy = in_is_sheet.cell(row = i, column =j+2).value
        out_is_sheet.cell(row = i, column =j , value = copy)

out_is_filename = 'is_result.xlsx'
out_is_workbook.save(out_is_filename)



import pandas as pd

# 両者をマージ

bsf_result = pd.read_excel("bsf_result.xlsx")
is_result = pd.read_excel("is_result.xlsx")

name_result = pd.merge(bsf_result, is_result, on='Breakdown')

name_result.to_excel('../0_Storage/name_result.xlsx', index=False)



import os

# 不要なファイルを削除

os.remove(out_bsf_filename)
os.remove(out_is_filename)