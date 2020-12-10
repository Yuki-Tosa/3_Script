name = input('ティッカーシンボルを入力: ')


import openpyxl
import pandas as pd

in_workbook_pass = '../1_Storage/{0}_new_join.xlsx'.format(name)
tmp_result = pd.read_excel(in_workbook_pass)

result = tmp_result[['Breakdown', 'Total Revenue', 'Net Income Common Stockholders', 'Total Assets', 'Common Stock Equity', 'Ordinary Shares Number']]

result['Total Revenue TTM'] = result['Total Revenue'].rolling(4).sum()
result['Net Income Common Stockholders TTM'] = result['Net Income Common Stockholders'].rolling(4).sum()
result['Total Assets ATTM'] = result['Total Assets'].rolling(4).sum() / 4
result['Common Stock Equity ATTM'] = result['Common Stock Equity'].rolling(4).sum() / 4

result_copy = result.copy()
result_fillna = result_copy.fillna('-')

result_pass = '../Result/{0}_result.xlsx'.format(name)
result_fillna.to_excel(result_pass, index=False)