# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from viewlib import *

# code cab be assigend
code = '600036'

try:
    wb = load_workbook(fn, keep_vba=True)
except Exception as e:
    print(str(e))
    os._exit(0)
ws = wb['Sheet1']
rd = RawData()
rd.reset(rd)

# get name, price
ws.cell(row_viewname, col_code).value = get_sina_id(code)
name, price = get_name_price(code)
ws.cell(row_viewname, col_share_name).value = name
ws.cell(row_viewname, col_price).value = price
year_title(ws, rd, code, row_year, col_start)
get_ROE(ws, rd, code, row_ROE, col_start)
get_fcff(ws, rd, code, row_fcff, col_start)         # 企业自由现金流量

try:
    wb.save(fn)
except Exception as e:
    print(str(e))
    os._exit(0)
    
print('\n finished')
