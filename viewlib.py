# -*- coding: utf-8 -*-
import os
import time
import datetime
from urllib.request import Request
from urllib.request import urlopen
import tushare as ts
import pandas as pd

work_catalog = "c:\PythonWork"
TOKEN = 'c27f964551786735a0cebbc26a743d0e18b06e9181f2166632964e37'
url_quotation_before = "http://hq.sinajs.cn/list="
fn = work_catalog + '\overview.xlsm'
balancesheet_fields = 'total_share, end_date'
income_fields = 'total_revenue, operate_profit, n_income, end_date'

row_viewname = 1
row_year = row_viewname + 1
row_ROE = row_year + 1
row_fcff = row_ROE + 1          # 自由现金流
row_total_share = row_fcff + 1
row_total_revenue = row_total_share + 1          # 营业总收入
row_operate_profit = row_total_revenue + 1      # 营业利润
row_net_income = row_operate_profit + 1         # 净利润

# ======= for sheet 'Sheet1' =============
# ------- for row_viewname --------
col_code = 1
col_share_name = col_code + 1
col_price = col_share_name + 1
# ------- for row_viewname --------
col_title = 1
col_start = col_title + 1

def is_number(variate):
    flag = False
    if isinstance(variate,int):
        flag = True
    elif isinstance(variate,float):
        flag = True
    else:
        flag = False
    return(flag)
def get_t_s_id(sID):
    stkID = ''
    if( sID[0:2] == '00' ):
        stkID = sID+'.SZ'
    elif( sID[0:2] == '30' ):
        stkID = sID+'.SZ'
    elif( sID[0:2] == '60' ):
        stkID = sID+'.SH'
    else:
        print('in get_t_s_name the id :', sID)
    return(stkID)  
def get_sina_id(sID):
    stkID = ''
    if( sID[0:2] == '00' ):
        stkID = 'sz'+sID
    elif( sID[0:2] == '30' ):
        stkID = 'sz'+sID
    elif( sID[0:2] == '60' ):
        stkID = 'sh'+sID
    else:
        print('in get_sina_name the id :', sID)
    return(stkID)  
def get_name_price(s_code):                   # get name and current price
    url = url_quotation_before + get_sina_id(s_code)
    req = Request(url)
    req.add_header('User-Agent','Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.87 Safari/537.36') 
    try_done = False
    while try_done == False :
        try_done = True
        try:
            quots = urlopen(req).read()
        except Exception as e:
            print('---2--- : ' + str(e))
            print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            try_done = False
            time.sleep(300)
    quot = quots.decode('gbk')
    quot_msg = quot.split(',')
    if( len(quot_msg) > 3 ):
        name = quot_msg[0].split(u'="')[1]
        price = float(quot_msg[3])
    else:
        name = ''
        price = 0.0
    return(name, price)
def get_today():
    today=datetime.date.today()
    formatted_today=today.strftime('%Y%m%d')
    return( formatted_today )
def year_title(ws, rd, code, row, col):
    end = False
    year = int(get_today()[0:4]) - 1
    years_cnt = 0
    while( not end ):
        df = rd.req_tushare_query(rd, code, str(year)+'1231')
        if( len(df) != 0 ):
            ws.cell(row, col+years_cnt).value = year
        else:
            end = True
        year -= 1
        years_cnt += 1
    ws.cell(row, 1).value = years_cnt - 1
def get_ROE(ws, rd, code, row, col):
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = 'ROE'
    for i in range(y_num):
        df = rd.req_tushare_query(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            ws.cell(row, col+i).value = df.iloc[0]['roe']
def get_fcff(ws, rd, code, row, col):           # 企业自由现金流量
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = u'自由现金流'
    for i in range(y_num):
        df = rd.req_tushare_query(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            if(is_number(df.iloc[0]['fcff'])):
                ws.cell(row, col+i).value = round(df.iloc[0]['fcff'] / 100000000)
def total_share(ws, rd, code, row, col):           # 期末总股本
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = u'期末总股本'
    for i in range(y_num):
        df = rd.req_balancesheet(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            if(is_number(df.iloc[0]['total_share'])):
                ws.cell(row, col+i).value = round(df.iloc[0]['total_share'] / 100000000, 2)
def total_revenue(ws, rd, code, row, col):           # 营业总收入
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = u'营业总收入'
    for i in range(y_num):
        df = rd.req_income(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            if(is_number(df.iloc[0]['total_revenue'])):
                ws.cell(row, col+i).value = round(df.iloc[0]['total_revenue'] / 100000000, 2)
def operate_profit(ws, rd, code, row, col):           # 营业利润
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = u'营业利润'
    for i in range(y_num):
        df = rd.req_income(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            if(is_number(df.iloc[0]['operate_profit'])):
                ws.cell(row, col+i).value = round(df.iloc[0]['operate_profit'] / 100000000, 2)
def net_income(ws, rd, code, row, col):           # 净利润
    year = int(get_today()[0:4]) - 1
    y_num = ws.cell(row_year, 1).value
    ws.cell(row, col_title).value = u'净利润'
    for i in range(y_num):
        df = rd.req_income(rd, code, str(year-i)+'1231')
        if( len(df) != 0 ):
            if(is_number(df.iloc[0]['n_income'])):
                ws.cell(row, col+i).value = round(df.iloc[0]['n_income'] / 100000000, 2)

class delay_ctl():
    cnt = 0
    time_interval = 0
    freq_interval = 0
    tm_bf = []
    def init(self, cls, time_interval, freq_interval):
        cls.cnt = 0
        cls.time_interval = time_interval
        cls.freq_interval =freq_interval
        tm = datetime.datetime.now() - datetime.timedelta(seconds = time_interval * 2)
        for i in range(freq_interval):
            cls.tm_bf.append(tm)
    def ctl(self, cls):
        tm = datetime.datetime.now()
        tm_diff = cls.time_interval - (tm-cls.tm_bf[cls.cnt]).seconds
        if( tm_diff > 0 ):
            print('--- sleep: ', tm_diff, ' seconds. ---')
            time.sleep(tm_diff + 0.1)
            tm = datetime.datetime.now()
        cls.tm_bf[cls.cnt] = tm
        cls.cnt += 1
        if( cls.cnt == cls.freq_interval ):
            cls.cnt = 0
    def prt(self, cls):
        for i in range( cls.freq_interval ):
            print(cls.tm_bf[i])
class RawData():
    ts.set_token(TOKEN)
    pro = ts.pro_api()
    df_query = pd.DataFrame()
    df_stock_basic = pd.DataFrame()
    df_dividend = pd.DataFrame()
    df_balancesheet = pd.DataFrame()
    df_income = pd.DataFrame()
    df_forecast = pd.DataFrame()
    df_express = pd.DataFrame()
    dc = delay_ctl()
    dc.init(dc, 60, 80)
    def reset(self, cls):
        cls.df_query = pd.DataFrame()
        cls.df_dividend = pd.DataFrame()
        cls.df_balancesheet = pd.DataFrame()
        cls.df_income = pd.DataFrame()
        cls.df_forecast = pd.DataFrame()
        cls.df_express = pd.DataFrame()
    def req_tushare(self, cls, mode, para):
        if( mode == 'query'):
            df = cls.pro.query('fina_indicator', ts_code=para[0], period=para[1])
        elif( mode == 'stock_basic' ):
            df = cls.pro.stock_basic(exchange='', list_status=para[0], fields=para[1])
        elif( mode == 'dividend' ):
            df = cls.pro.dividend(ts_code=get_t_s_id(para[0]), fields=para[1])
        elif( mode == 'balancesheet' ):
            df = cls.pro.balancesheet(ts_code=para[0], period=para[1], fields=para[2])
        elif( mode == 'income' ):
            df = cls.pro.income(ts_code=para[0], period=para[1], fields=para[2])
        elif( mode == 'forecast' ):
            df = cls.pro.forecast(ts_code=get_t_s_id(para[0]), start_date=para[1], end_date=para[2], fields=para[3])
        elif( mode == 'express' ):
            df = cls.pro.express(ts_code=get_t_s_id(para[0]), start_date=para[1], end_date=para[2], fields=para[3])
        else:
            df = None
            print('mode:', mode, ' not exist.')
        # sleep
        cls.dc.ctl(cls.dc)
        return(df)
    def req_tushare_query(self, cls, code, period):
        get = False
        if(cls.df_query.shape[0] != 0):
            for i in range(cls.df_query.shape[0]):
                if( cls.df_query.iloc[i]['end_date'] == period ):
                    df = cls.df_query.iloc[[i]]
                    get = True
                    break
        if( get == False ):
            mode = 'query'
            para = []
            para.append(get_t_s_id(code))
            para.append(period)
            df = self.req_tushare(cls, mode, para)
            if( df.shape[0] != 0 ):
                cls.df_query = cls.df_query.append(df, ignore_index=True)
        return(df)
    def req_balancesheet(self, cls, code, period):
        get = False
        if(cls.df_balancesheet.shape[0] != 0):
            for i in range(cls.df_balancesheet.shape[0]):
                if( cls.df_balancesheet.iloc[i]['end_date'] == period ):
                    df = cls.df_balancesheet.iloc[[i]]
                    get = True
                    break
        if( get == False ):
            mode = 'balancesheet'
            para = []
            para.append(get_t_s_id(code))
            para.append(period)
            para.append(balancesheet_fields)
            df = self.req_tushare(cls, mode, para)
            if( df.shape[0] != 0 ):
                cls.df_balancesheet = cls.df_balancesheet.append(df, ignore_index=True)
        return(df)
    def req_income(self, cls, code, period):
        get = False
        if(cls.df_balancesheet.shape[0] != 0):
            for i in range(cls.df_income.shape[0]):
                if( cls.df_income.iloc[i]['end_date'] == period ):
                    df = cls.df_income.iloc[[i]]
                    get = True
                    break
        if( get == False ):
            mode = 'income'
            para = []
            para.append(get_t_s_id(code))
            para.append(period)
            para.append(income_fields)
            df = self.req_tushare(cls, mode, para)
            if( df.shape[0] != 0 ):
                cls.df_income = cls.df_income.append(df, ignore_index=True)
        return(df)
    def req_stock_basic(self, cls):
        if(cls.df_stock_basic.shape[0] == 0):
            mode = 'stock_basic'
            para = []
            para.append('L')
            para.append('symbol,area,industry,list_date')
            cls.df_stock_basic = self.req_tushare(cls, mode, para)
        return(cls.df_stock_basic)
    def req_dividend(self, cls, code):
        if(cls.df_dividend.shape[0] == 0):
            mode = 'dividend'
            para = []
            para.append(code)
            para.append('cash_div_tax,div_proc,end_date,record_date,ex_date,stk_bo_rate,stk_co_rate,stk_div')
            cls.df_dividend = self.req_tushare(cls, mode, para)
        return(cls.df_dividend)