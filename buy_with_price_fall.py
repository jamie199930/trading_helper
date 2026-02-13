import pandas as pd
import datetime
import time
import utility_functions as uf
import yfinance as yf
import traceback
import os
from dotenv import load_dotenv

load_dotenv('.env')
receiver1 = os.getenv('RECEIVER1')
receiver2 = os.getenv('RECEIVER2')

data = pd.read_excel('TWSE.xlsx')
target_stocks = data['代號'].tolist()

try:
    today = datetime.date.today()
    if_trade = uf.is_open(today)
    if if_trade == 'Close':
        receiver1 = os.getenv('RECEIVER1')
        arr_list = [receiver1, receiver2]
        title = '連兩日暴跌中的股票'
        body = f'{today} 今日未開盤！'
        mode = ''
        file_path = ['./sugar.jpg']
        file_name = ['sugar.jpg']
        uf.sendmail(arr_list, title, body, mode, file_path, file_name)

    else:

        start_date = today + datetime.timedelta(days=-20)
        end_date = today + datetime.timedelta(days=1)

        start_date = start_date.strftime('%Y-%m-%d')
        end_date = end_date.strftime('%Y-%m-%d')

        stock_store = []
        today_store = [] # 儲存今日價格
        today_fall = []  # 儲存今日跌幅
        count = 0

        ## 篩選出連續兩天跌幅超5%股票
        for i in target_stocks:
            count += 1
            # time.sleep(1)
            stock = yf.Ticker(f'{i}.TW')
            stock_info = stock.history(start = start_date, end = end_date)
            
            if len(stock_info) >= 5:
                today_price = stock_info['Close'].values[-1]
                yes_price = stock_info['Close'].values[-2]
                be_yes_price = stock_info['Close'].values[-3]

                ## 計算漲跌幅
                fall_today = ((today_price - yes_price)/yes_price)*100
                fall_yes = ((yes_price - be_yes_price)/be_yes_price)*100

                ## 如果兩天的漲跌幅都小於-5就是目標
                if fall_today <= -5 and fall_yes <= -5:
                    stock_store.append(i)
                    today_store.append(today_price)
                    today_fall.append(fall_today)
            print(f'Dealing stock : {i} | All stock : {len(target_stocks)} | Now : {count}')


            control = 0
        main_df = pd.DataFrame()
        for i in stock_store:
            # time.sleep(1)
            if control == 0:
                main_df = uf.get_yahoo_news2(i)
                main_df['stock'] = len(main_df)*[i]
                control += 1
            else:  # 已經有寫入的新聞
                merge_df = uf.get_yahoo_news2(i)
                merge_df['stock'] = len(merge_df)*[i]
                main_df = pd.concat([main_df, merge_df])
        main_df.to_excel('fall_stock_news.xlsx', index=False)

        empty_df = pd.DataFrame()
        empty_df['股票代號'] = stock_store 
        empty_df['今日價格'] = today_store
        empty_df['今日跌幅'] = today_fall

        empty_df = empty_df.to_html(index=False)
        arr_list = [receiver1, receiver2]
        title = f'{today} 連兩日跌幅5%的股票'
        body = f'''
        <html>
        <font face="微軟正黑體"></font>
        <body>
            <h4>
                連兩日跌幅5%的股票
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
    today = datetime.date.today()
    arr_list = [receiver1]
    title = f'{today} 連兩日暴跌中的股票'
    body = f'異常情況：\n{traceback.format_exc()}'
    mode = ''
    file_path = None
    file_name = None
    uf.sendmail(arr_list, title, body, mode, file_path, file_name) 
