import xlwings as xw
import datetime as dt
import pandas as pd
from datetime import date
import requests
import time
import RStockvn as rpv
from selenium.webdriver.common.by import By
from selenium import webdriver
import gdown
from datetime import datetime, timedelta

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

@xw.func()
def key_id(code):
    day=rpv.key_id(str(code))
    return f'Đã kích hoạt {day} ngày'

def giao_dich_tu_doanh(symbol,fromdate,todate):
    fromdate = pd.to_datetime(fromdate)
    fdate = fromdate.strftime('%d/%m/%Y')
    todate = pd.to_datetime(todate)
    tdate = todate.strftime('%d/%m/%Y')
    df,data=rpv.get_proprietary_history_cafef(symbol.upper(),fdate,tdate)
    return data


@xw.func()
def list_company():
    df=rpv.list_company()
    return df

@xw.func()
def update_company():
    data=rpv.update_company()
    return data


@xw.func()
def lai_suat_cafef():
    data=rpv.lai_suat_cafef()
    return data



@xw.func()
def giao_dich_noi_bo(symbol,fromdate,todate):
    fromdate = pd.to_datetime(fromdate)
    fdate = fromdate.strftime('%d/%m/%Y')
    todate = pd.to_datetime(todate)
    tdate = todate.strftime('%d/%m/%Y')
    data=rpv.get_insider_transaction_history_cafef(symbol.upper(),fdate,tdate)
    return data

@xw.func()
def giao_dich_khoi_ngoai(symbol,fromdate,todate):
    fromdate = pd.to_datetime(fromdate)
    fdate = fromdate.strftime('%d/%m/%Y')
    todate = pd.to_datetime(todate)
    tdate = todate.strftime('%d/%m/%Y')
    data=rpv.get_foreign_transaction_history_cafef(symbol.upper(),fdate,tdate)
    return data


@xw.func()
def report_finance_vnd(symbol,types,year_f,timely):
    symbol,types, timely=symbol.upper(), types.upper(), timely.upper()
    year_f=int(year_f)
    data=rpv.report_finance_vnd(symbol,types,year_f,timely)
    return data

@xw.func()
def report_finance_cf(symbol,report,year,timely):
    symbol, report, year, timely = str(symbol), str(report), int(year), str(timely)
    data=rpv.report_finance_cf(symbol,report,year,timely)
    return data


@xw.func(async_mode='threading')
def exchange_currency(current,cover_current,from_date,to_date):
    current=str(current)
    cover_current=str(cover_current)
    from_date = pd.to_datetime(from_date,infer_datetime_format=True)
    to_date = pd.to_datetime(to_date,infer_datetime_format=True)
    data=rpv.exchange_currency(current,cover_current,str(from_date.strftime('%Y-%m-%d')),str(to_date.strftime('%Y-%m-%d')))
    return data


@xw.func()
def laisuat_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.laisuat_vietstock(fromdate,todate)
    return data

@xw.func()
def getCPI_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.getCPI_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_sanxuat_congnghiep(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_sanxuat_congnghiep(fromdate,todate)
    return data

@xw.func()
def solieu_banle_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_banle_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_XNK_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_XNK_vietstock(fromdate,todate)
    return data

@xw.func()
def solieu_FDI_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_FDI_vietstock(fromdate,todate)
    return data   

@xw.func()
def tygia_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.tygia_vietstock(fromdate,todate)
    return data 


@xw.func()
def solieu_tindung_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_tindung_vietstock(fromdate,todate)
    return data 

@xw.func()
def solieu_danso_vietstock(fromdate,todate):
    fromdate=str(fromdate)
    todate=str(todate)
    data=rpv.solieu_danso_vietstock(fromdate,todate)
    return data 

@xw.func()
def solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ):
    fromyear=int(fromyear)
    fromQ=int(fromQ)
    toyear=int(toyear)
    toQ=int(toQ)
    data=rpv.solieu_GDP_vietstock(fromyear,fromQ,toyear,toQ)
    return data 


@xw.func(async_mode='threading')
def get_price_history_cafef(symbol,fromdate,todate):
    fromdate = pd.to_datetime(fromdate)
    fdate = fromdate.strftime('%d/%m/%Y')
    todate = pd.to_datetime(todate)
    tdate = todate.strftime('%d/%m/%Y')
    data=rpv.get_price_history_cafef(symbol.upper(),fdate,tdate)
    return data


if __name__ == "__main__":
    xw.Book("func.xlsm").set_mock_caller()
    main()