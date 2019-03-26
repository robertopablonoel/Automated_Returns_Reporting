#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 23 09:22:02 2017

@author: robertonoel
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup as bs
#import pandas_datareader as pdr
from datetime import datetime, timedelta
import time
import datetime as dt
import re
import os
import pandas as pd
#import requests



#Define Google User
def define_user():
    scope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(os.getcwd() + '\\' + 'Voya-2e4969426678.json', scope)
    client = gspread.authorize(creds)
    file  = client.open(input('What file would you like to edit?\n >>> '))
    sheet_list = list(file.worksheets())
    sheet_list_len = len(sheet_list)
    return file, sheet_list_len

        
#Get tickers
def get_tickers(sheet):
    global scope
    global creds
    global client
    tickers = []
    try:
        index = sheet.row_values(1).index("Tickers")
    except:
        index = 0
    tickers = sheet.col_values(index + 1)
    tickers = [x for x in tickers if x != '']
    print(tickers[1:])
    return tickers[1:]
'''
#Get historical data for one stock
def get_historical_data(ticker, date):
    data = pdr.get_data_yahoo(symbols=ticker, start=datetime(date[0], date[1], date[2]), end=datetime(date[0], date[1], date[2]))
    data_frame = data['Adj Close']
    data_list = list(data_frame)
    return data_list[0]
'''


#get earnings release dates
def get_release_dates(tickers):
    release_dates = []
    for y in range(0, len(tickers)):
        page = requests.get("https://finviz.com/quote.ashx?t=%s" % tickers[y])
        soup = bs(page.content, 'html.parser')
        b = soup.find_all('b')
        b = list(b)
        if len(b) >= 68:
            date = str(b[68])
        elif len(b) < 68:
            date = '-'    
        release_date = date.lstrip('<b>').rstrip('</b>')
        release_dates.append(release_date)
    return release_dates



#write to google sheets
def write_to_sheets(sheet_range, write_data, sheet):
    counter = 0
    cell_list = sheet.range(sheet_range)
    for cell in cell_list:
        print(counter)
        cell.value = write_data[counter]
        counter += 1
    sheet.update_cells(cell_list)

def split_crumb_store(v):
    return v.split(':')[2].strip('"')


def find_crumb_store(lines):
    # Looking for
    # ,"CrumbStore":{"crumb":"9q.A4D1c.b9
    for l in lines:
        if re.findall(r'CrumbStore', l):
            return l
    print("Did not find CrumbStore")


def get_cookie_value(r):
    return {'B': r.cookies['B']}


def get_page_data(symbol):
    url = "https://finance.yahoo.com/quote/%s/?p=%s" % (symbol, symbol)
    r = requests.get(url)
    cookie = get_cookie_value(r)

    # Code to replace possible \u002F value
    # ,"CrumbStore":{"crumb":"FWP\u002F5EFll3U"
    # FWP\u002F5EFll3U
    lines = r.content.decode('unicode-escape').strip(). replace('}', '\n')
    return cookie, lines.split('\n')


def get_cookie_crumb(symbol):
    cookie, lines = get_page_data(symbol)
    crumb = split_crumb_store(find_crumb_store(lines))
    return cookie, crumb


def get_data(symbol, start_date, end_date, cookie, crumb):
    filename = os.getcwd() + '\\' + '%s.csv' % (symbol)
    url = "https://query1.finance.yahoo.com/v7/finance/download/%s?period1=%s&period2=%s&interval=1d&events=div&crumb=%s" % (symbol, start_date, end_date, crumb)
    response = requests.get(url, cookies=cookie)
    with open (filename, 'wb') as handle:
        for block in response.iter_content(1024):
            handle.write(block)


def get_now_epoch():
    # @see https://www.linuxquestions.org/questions/programming-9/python-datetime-to-epoch-4175520007/#post5244109
    return int(time.time())


def download_quotes(symbol):
    start_date = 0
    end_date = get_now_epoch()
    cookie, crumb = get_cookie_crumb(symbol)
    get_data(symbol, start_date, end_date, cookie, crumb)


def make_df(ticker):
            download_quotes(ticker)
            div_df = pd.read_csv(os.getcwd() + '\\' + '%s.csv' % (ticker)).sort_values(by='Date')[::-1].set_index('Date')
            return(div_df)
            
            
    

def get_sum(dates, ticker):
    
    #insert dates into list
    new_dates = [date[0:4] + '-' + date[4:6] + '-' + date[6:] for date in dates]
    try:
        div_data = make_df(ticker)
    except:
        return 'N/A'
    
    

    all_dates = new_dates + div_data.index.values.tolist()
    formatted_dates = list(set([dt.datetime.strptime(ts, "%Y-%m-%d") for ts in all_dates]))
    formatted_dates.sort(reverse = True)
    sorted_dates = [dt.datetime.strftime(ts, "%Y-%m-%d") for ts in formatted_dates]
    
    #print(sorted_dates)
    sig_dates = []
    for item in sorted_dates:
        if sorted_dates.index(item) < sorted_dates.index(new_dates[0]) and sorted_dates.index(item) >= sorted_dates.index(new_dates[1]):
            sig_dates.append(item)
        else:
            continue
    #print(sig_dates)
    try:
        divs_in_period = div_data.loc[sig_dates,:]
    except:
        divs_in_period = div_data.loc[sig_dates[1:-2],:]
    print(divs_in_period.sum().values)
    return(divs_in_period.sum().values[0])
    


def get_div_list(dates, tickers):
    div_list = []
    for ticker in tickers:
        if ticker != "CASH":
            div = get_sum(dates, ticker)
        else:
            div = 0
        div_list.append(div)
        if ticker != "CASH":
            os.remove(os.getcwd() + '\\' + '%s.csv' % (ticker))
    return(div_list)

def get_hist_data(symbol, start_date, end_date, cookie, crumb):
    filename = os.getcwd() + '\\%s.csv' % (symbol)
    url = "https://query1.finance.yahoo.com/v7/finance/download/%s?period1=%s&period2=%s&interval=1d&events=history&crumb=%s" % (symbol, start_date, end_date, crumb)    
    response = requests.get(url, cookies=cookie)
    with open (filename, 'wb') as handle:
        for block in response.iter_content(1024):
            handle.write(block)

def download_price_quotes(symbol, date):
    new_date = date[0:4] + '-' + date[4:6] + '-' + date[6:]
    pattern = '%Y-%m-%d' 
    start_date = int(time.mktime(time.strptime(new_date, pattern))) - (24*3600)
    end_date = int(time.mktime(time.strptime(new_date, pattern))) + (24*3600)
    cookie, crumb = get_cookie_crumb(symbol)
    get_hist_data(symbol, start_date, end_date, cookie, crumb)


def make_price_df(ticker, date):
            download_price_quotes(ticker, date)
            price_df = pd.read_csv(os.getcwd() + '\\%s.csv' % (ticker)).sort_values(by='Date')[::-1].set_index('Date')
            return(price_df)

def get_price(ticker, date):
    price_data = make_price_df(ticker, date)
    ref_date = date[0:4] + '-' + date[4:6] + '-' + date[6:]
    try:
        df_line = price_data.loc[ref_date,:]
        adj_close = df_line['Adj Close']
    except:
        adj_close = 'N/A'
    return(adj_close)
    
#get historical prices for all stocks
def get_all_historical_data(tickers, date):
    historical_prices = []
    for ticker in tickers:
        try:
            if ticker != "CASH":
                historical_price = get_price(ticker, date)
            else:
                historical_price = 1
        except:
            historical_price = 'N/A'
        print(historical_price)
        print(ticker)
        historical_prices.append(historical_price)
        if ticker != "CASH":
            os.remove(os.getcwd() + '\\' + '%s.csv' % (ticker))
    return historical_prices
            
        
    
def main():
    n = 1
    yesterday = datetime.now() - timedelta(days=n)
    while yesterday.strftime('%A') == 'Sunday' or yesterday.strftime('%A') == 'Saturday':
        yesterday = datetime.now() - timedelta(days= n)
        n += 1
    user = define_user()
    file = user[0]
    sheets = user[1]
    for x in range(0, sheets):
        sheet = file.get_worksheet(x)
        tickers = get_tickers(sheet)
        print(tickers)
        functions = ['DivX', 'Stock Prices', 'Earnings Release Dates', 'Historical Price', 'Price Earnings']
        functions_2 = sheet.row_values(1)
        
        for funct in functions_2:
            funct_split = funct.split(' (')
            print(funct)
            try:
                data_request = functions.index(funct_split[0])
            except:
                data_request = ''
            if data_request == 0:
                date_string = funct_split[1].rstrip(')')
                dates = date_string.split(':')
                print(dates)
                data = get_div_list(dates, tickers)
            elif data_request == 1:
                continue
            elif data_request == 2:
                data = get_release_dates(tickers)
            elif data_request == 3:
                #date_string = str(input("What date do you want data from? (input as YYYYMMDD)"))
                if funct_split[1].rstrip(')') == 'yesterday':
                    date_string = yesterday.strftime("%Y%m%d")
                else:   
                    date_string = funct_split[1].rstrip(')')
                #date = [int(date_string[0:4]), int(date_string[4:6].lstrip('0')), int(date_string[6:].lstrip('0'))]
                date = date_string
                print(date)
                #counter = 0
                data = get_all_historical_data(tickers, date)
            elif data_request == 4:
                continue
            else:
                continue
            print(data)
            funct_index = functions_2.index(funct)
            columns_raw = sheet.col_values(1)
            columns = [x for x in columns_raw if x != '']
            letter = str(chr(funct_index + 65))
            print(letter, str(len(columns)))
            write_to_sheets(letter + '2' + ':' + letter + str(len(columns)), data, sheet)

main()






