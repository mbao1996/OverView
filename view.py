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
# get ROE
end = False
year = int(get_today()[0:4]) - 1
col_year = 2
ws.cell(row_ROE, col_title).value = 'ROE'
while( not end ):
    df = rd.req_tushare_query(rd, code, str(year)+'1231')
    if( len(df) != 0 ):
        print(year,'---', df.iloc[0]['roe'])
        ws.cell(row_year, col_year).value = year
        ws.cell(row_ROE, col_year).value = df.iloc[0]['roe']
    else:
        print(year)
        end = True
    col_year += 1
    year -= 1    
# 企业自由现金流量
end = False
year = int(get_today()[0:4]) - 1
col_year = 2
ws.cell(row_fcff, col_title).value = u'自由现金流'
while( not end ):
    print(year)
    df = rd.req_tushare_query(rd, code, str(year)+'1231')
    if( len(df) != 0 ):
        if( is_number(df.iloc[0]['fcff']) ):
            ws.cell(row_fcff, col_year).value = round(df.iloc[0]['fcff'] / 10000)
    else:
        end = True
    col_year += 1
    year -= 1    

try:
    wb.save(fn)
except Exception as e:
    print(str(e))
    os._exit(0)
    
print('\n finished')
