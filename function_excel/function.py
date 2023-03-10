# Copyright 2023 NPhucBinh @ GitHub
# See LICENSE for details.
import xlwings as xw
import datetime as dt
import pandas as pd
import datetime as dt
from dateutil import parser
from datetime import date
import requests
import time
from io import BytesIO
import RStockvn as rpv



def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

@xw.func
def hello(name):
    return f"Hello {name}!"

@xw.func(async_mode='threading')
def report_finance_cp68(symbol,reporty,timely):
    data=rpv.report_finance_cp68(symbol,reporty,timely)
    return data





@xw.func(async_mode='threading')
def report_finance_cf(symbol,report,year,typer):
    data= rpv.report_finance_cf(symbol,report,year,typer)
    return data

@xw.func(async_mode='threading')
def event_price_cp68(symbol):
    data=rpv.event_price_cp68(symbol)
    return data

@xw.func(async_mode='threading')
def info_company(symbol):
    data=rpv.info_company(symbol)
    return data

@xw.func(async_mode='threading')
def giaodich_noibo(symbol):
    data=rpv.trade_internal(symbol)
    return data

@xw.func(async_mode='threading')
def exchange_currency(current,cover_current,from_date,to_date):
    current=str(current)
    cover_current=str(cover_current)
    from_date = pd.to_datetime(from_date,infer_datetime_format=True)
    to_date = pd.to_datetime(to_date,infer_datetime_format=True)
    data=rpv.exchange_currency(current,cover_current,str(from_date.strftime('%Y-%m-%d')),str(to_date.strftime('%Y-%m-%d')))
    return data

@xw.func(async_mode='threading')
def baocaonhanh(mcp,loai,time):
    data=rpv.baocaonhanh(mcp,loai,time)
    return data

@xw.func(async_mode='threading')
def laisuat_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.laisuat_vietstock(fromdate,todate)
    return data

@xw.func(async_mode='threading')
def getCPI_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.getCPI_vietstock(fromdate,todate)
    return data

@xw.func(async_mode='threading')
def solieu_sanxuat_congnghiep(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_sanxuat_congnghiep(fromdate,todate)
    return data

@xw.func(async_mode='threading')
def solieu_banle_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_banle_vietstock(fromdate,todate)
    return data

@xw.func(async_mode='threading')
def solieu_XNK_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_XNK_vietstock(fromdate,todate)
    return data

@xw.func(async_mode='threading')
def solieu_FDI_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_FDI_vietstock(fromdate,todate)
    return data   

@xw.func(async_mode='threading')
def tygia_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.tygia_vietstock(fromdate,todate)
    return data 


@xw.func(async_mode='threading')
def solieu_tindung_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_tindung_vietstock(fromdate,todate)
    return data 

@xw.func(async_mode='threading')
def solieu_danso_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_danso_vietstock(fromdate,todate)
    return data 

@xw.func(async_mode='threading')
def solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ):
    fromyear=int(fromyear)
    fromQ=int(fromQ)
    toyear=int(toyear)
    toQ=int(toQ)
    data=rpv.solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ)
    return data 
if __name__ == "__main__":
    xw.Book("func.xlsm").set_mock_caller()
    main()
