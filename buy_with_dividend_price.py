import pandas as pd
import numpy as np
import yfinance as yf
import datetime
import utility_functions as uf
import traceback
import os
from dotenv import load_dotenv

load_dotenv('.env')
receiver1 = os.getenv('RECEIVER1')
receiver2 = os.getenv('RECEIVER2')

try:
    today = datetime.date.today()
    if_trade = uf.is_open(today)
    if if_trade == 'Close':
        arr_list = [receiver1, receiver2]
        title = '高配息低股價'
        body = f'{today} 今日未開盤！'
        mode = ''
        file_path = ['./sugar.jpg']
        file_name = ['sugar.jpg']
        uf.sendmail(arr_list, title, body, mode, file_path, file_name)
    else:
        data = pd.read_excel('Dividend_list.xlsx')
        target_stock = data['代號'].tolist()

        # 整理時間
        today = datetime.date.today()
        start_date = today + datetime.timedelta(days=-365)
        end_date = today + datetime.timedelta(days=-1)

        start_date = start_date.strftime('%Y-%m-%d')
        end_date = end_date.strftime('%Y-%m-%d')


        target_store = []
        target_name_store = []
        highest_store = []
        now_price_store = []

        # 對於目標股票取出一年內最高價、最近一次的收盤價
        for target in target_stock:
            try:
                stock = yf.Ticker(f'{target}.TW')
                stock_info = stock.history(start=start_date, end=end_date)
                highest = np.max(stock_info['High'].values)
                now_price = stock_info['Close'].values[-1]

                if now_price < highest*0.5:
                    target_store.append(target)
                    highest_store.append(highest)
                    now_price_store.append(now_price)
            except:
                print(f'Error stock ! Stock :{target}')

        ## 得到股票名稱
        df = pd.read_excel('2.2TWSE.xlsx')
        for st in target_store:
            stock = df[df['代號'] == st]
            target_name_store.append(stock['股票名稱'].values[0])


        data = pd.DataFrame()
        data['日期'] = len(target_store)*[today]
        data['代號'] = target_store
        data['股票名稱'] = target_name_store
        data['一年內最高價'] = highest_store
        data['目前價格'] = now_price_store
        empty_df = data.to_html(index=False)

        arr_list = [receiver1, receiver2]
        title = f'{today} 高配息低股價 - 每日比對'
        body = f'''
        <html>
        <font face="微軟正黑體"></font>
        <body>
            <h4>
                偵測股票高配息且股價低
            </h4>
            {empty_df}
            <h5>
                投資理財有賺有賠，請僅慎評估風險。
            </h5>
        </body>
        </html>
        '''
        mode = 'html'
        file_path = None
        file_name = None
        uf.sendmail(arr_list, title, body, mode, file_path, file_name)
except:
    arr_list = [receiver1]
    title = '高配息低股價 異常回報'
    body = f'異常情況：\n{traceback.format_exc()}'
    mode = ''
    file_path = ''
    file_name = ''
    uf.sendmail(arr_list, title, body, mode, file_path, file_name)
