import xlwings as xw
import datetime as dt
import pandas as pd
from datetime import date
import requests
import time
from user_agent import random_user
import RStockvn as rpv
from selenium.webdriver.common.by import By
from selenium import webdriver
import gdown
from datetime import datetime
from datetime import timedelta
from bs4 import BeautifulSoup
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"




@xw.func()
def momentum_ck(symbol):
    data=rpv.momentum_ck(symbol)
    return data


@xw.func()
def CW_info(symbol):
    head1={"User-Agent":'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0'}
    pload1={}
    url2=f'https://finance.vietstock.vn/chung-khoan-phai-sinh/{symbol}/cw-tong-quan.htm'
    r=requests.get(url2,headers=head1,data=pload1)
    soup=BeautifulSoup(r.text,'html.parser')
    ds=soup.find(class_="table table-hover")
    df=pd.read_html(ds.prettify())[0]
    df.rename(columns={0:'CW',1:f'{symbol[1:].upper()}'},inplace=True)
    return df



@xw.func()
def Date_auto(time):
    today = dt.datetime.now()
    today_date = dt.datetime(today.year, today.month, today.day)
    return today_date




@xw.func()
def thoi_gian_CW(time):
    today = dt.datetime.now()
    # Chuyển chuỗi thành đối tượng datetime
    fd = pd.to_datetime(time, dayfirst=True)
    # Chuyển today và fd về cùng một dạng (00:00:00)
    today_date = dt.datetime(today.year, today.month, today.day)
    fd_date = dt.datetime(fd.year, fd.month, fd.day)
    diff_days = fd_date - today_date
    return diff_days.days



def tinh_ngay(number):
    todate = dt.datetime.now()
    date=(number/4*1.9)
    fromdate = todate - timedelta(days=(number+date-1))
    fdate = fromdate.strftime('%Y-%m-%d')
    tdate = todate.strftime('%Y-%m-%d')
    return fdate, tdate



def TB_bien_dong_ck(symbol,number):
    fdate, edate= tinh_ngay(number)
    df=get_price_historical_vnd(symbol,fdate,edate)
    df['per_change'] = ((df['close'] - df['close'].shift(-1)) / df['close'].shift(-1)).astype(float)
    return df

@xw.func()
def tinh_phan_tram(symbol,number):
    df=TB_bien_dong_ck(symbol,number)
    df=df['per_change']
    return df.mean()




@xw.func()
def get_price_historical_vnd(symbol,fromdate,todate):
    fromdate, todate = pd.to_datetime(fromdate, dayfirst=True), pd.to_datetime(todate, dayfirst=True)
    fdate, tdate=fromdate.strftime('%Y-%m-%d'), todate.strftime('%Y-%m-%d')
    url=f'https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:{symbol.upper()}~date:gte:{fdate}~date:lte:{tdate}&size=100000&page=1' 
    head={"User-Agent":random_user()}
    payload={}
    r=requests.get(url,headers=head,data=payload)
    df=pd.DataFrame(r.json()['data'])
    data=df[['code','date','open','high','low','close','nmVolume','nmValue','ptVolume', 'ptValue']]
    data = data.copy()
    data.rename(columns={'nmVolume':'KLGD Khớp lệnh','nmValue':'GTGD Khớp lệnh','ptVolume':'KLGD Thỏa thuận','ptValue':'GTGD Thỏa thuận'}, inplace=True)
    return data




@xw.func()
def chuyen_ngay(todate):
    tdate=pd.to_datetime(todate, dayfirst=True)
    fromdate=tdate - timedelta(days=6)
    fromdate=fromdate.strftime('%d/%m/%Y')
    return fromdate




@xw.func()
def key_id(code):
    day=rpv.key_id(str(code))
    return day

@xw.func()
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


@xw.func()
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