name = input('ティッカーシンボルを入力: ')


import openpyxl
import pandas as pd

in_workbook_pass = '../1_Storage/{0}_new_join.xlsx'.format(name)
tmp_result = pd.read_excel(in_workbook_pass)

result = tmp_result[['Breakdown', 'Total Revenue', 'Net Income Common Stockholders', 'Total Assets', 'Common Stock Equity', 'Ordinary Shares Number']]

result_pass = '../Result/{0}_result.xlsx'.format(name)
result.to_excel(result_pass, index=False)