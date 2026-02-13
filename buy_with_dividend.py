import pandas as pd
import datetime
import time
import yfinance as yf
import traceback
import utility_functions as uf
import os
from dotenv import load_dotenv

load_dotenv('.env')
receiver1 = os.getenv('RECEIVER1')

try:
## 取得所有股票代號
    stock_list = pd.read_excel('TWSE.xlsx')
    all_stock = stock_list['代號'].values  

    count = 0
    dividend_store = []
    stock_store = []

    for i in all_stock:
        start = time.time()
        count += 1
        stock = yf.Ticker(f'{i}.TW')
        try:
            if stock.info['dividendYield'] != None:
                d_y = stock.info['dividendYield']
                if d_y != None and d_y >= 5:
                    stock_store.append(i)
                    dividend_store.append(d_y)
                else:
                    d_y = None
                end = time.time()
                print(f'Dealing : {count} | All : {len(all_stock)} | Stock : {i} | D_y : {d_y} | Cost time : {end - start}s')
        except:
            print(f'Error Stock ! Dealing : {count} | All : {len(all_stock)} | Stock : {i}')
            
    data = pd.DataFrame()
    data['代號'] = stock_store
    data['殖利率'] = dividend_store
    data.to_excel('Dividend_list.xlsx')
except : 
    today = datetime.date.today()
    arr_list = [receiver1]
    title = f'高配息名單篩選異常回報！'
    body = f'異常情況：\n{traceback.format_exc()}'
    mode = ''
    file_path = None
    file_name = None
    uf.sendmail(arr_list, title, body, mode, file_path, file_name)
