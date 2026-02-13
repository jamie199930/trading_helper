import pandas as pd
import requests
import datetime
import requests
import pandas as pd
from bs4 import BeautifulSoup

import sys
sys.path.append('C:/Users/jamie/Desktop/Trading strategy')
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from AES_Encryption.encrype_process import *

'''
獲取證交所三大法人買賣超資訊

r_date : 字串格式日期，為需要查詢三大法人買賣超日報的日期
'''

def twse_data(r_date:str, type:str):
    url = f'https://www.twse.com.tw/rwd/zh/fund/T86?date={r_date}&selectType={type}&response=json&_=1750831631952'
    response = requests.get(url)
    data_json = json.loads(response.text)

    data_store = pd.DataFrame(data_json['data'], columns = data_json['fields'])
    return data_store

'''
檢測今日有無休市

target_date:要傳入datetime日期格式,為需要判定是否開盤的日期
is_open會回傳Close/Open代表休市/開市
'''

def is_open(date:datetime.date):

    # weekend
    day = date.weekday()
    if day == 5 or day == 6:
        return "Close"

    ## holidays
    hd = pd.read_excel('./holiday.xlsx')
    hd_date = pd.to_datetime(hd['日期']).tolist()
    str_date = date.strftime('%Y%m%d')
    for i in hd_date:
        i = i.strftime('%Y%m%d')
        if (i == str_date):
            return 'Close'
    return 'Open'

'''
寄件函數

receiver : list, 收信者
title : str, 信件標題
body : str, 信件本文
mode : str, 支援text與html兩種寄信模式
file_path : list, 寄出的檔案位置
file_name : list, 寄出的檔案名稱及附檔名
'''

def sendmail(arr_list:list, title:str, body:str, mode:str, file_path:list, file_name:list):
    key_path = 'C:/Users/jamie/Desktop/key/'
    config_path = 'C:/Users/jamie/Desktop/config/'
    user_id, password = check_encrype('gmail', key_path, config_path)
    message = MIMEMultipart()

    message['From'] = user_id
    message['To'] = ','.join(arr_list)
    message['Subject'] = title
    if mode == 'html':
        message.attach(MIMEText(body, mode))
    else:
        message.attach(MIMEText(body))

    if file_path==None:
        pass
    else:
        for x in range(len(file_path)):
                with open(file_path[x], 'rb') as opened:
                      openedfile = opened.read()
                appfile = MIMEApplication(openedfile)
                appfile.add_header('content-disposition', 'attachment', filename=file_name[x])
                message.attach(appfile)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    server.login(user_id, password)
    text = message.as_string()
    server.sendmail(user_id, arr_list, text)
    print('The mail has been sent !')
    server.quit()

def get_yahoo_news(stock:str):
    url = f'https://tw.stock.yahoo.com/quote/{stock}/news'
    headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:139.0) Gecko/20100101 Firefox/139.0'
    }

    response = requests.get(url, headers = headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    # 取得新聞連結
    tags = soup.find_all('a')
    news_link_store = []
    for tag in tags:
        news_link = tag['href']
        if news_link[-4:] == 'html' and 'news' in tag['href'] :
           news_link_store.append(news_link)

    # 取得新聞標題、新聞日期
    title_store, date_store = [], []
    for link in news_link_store:
        headers = {
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:139.0) Gecko/20100101 Firefox/139.0'
        }
        response = requests.get(link)
        soup = BeautifulSoup(response.text, 'html.parser')
        title = soup.find('h1').text
        newstime = soup.find('time')
        newstime = newstime.text.split(' ')[0]
        newstime = newstime.replace('年', '/')
        newstime = newstime.replace('月', '/')
        newstime = newstime.replace('日', '')
        title_store.append(title)
        date_store.append(newstime)

    # 建立dataframe    
    yahoonews_df = pd.DataFrame()
    yahoonews_df['url'] = news_link_store
    yahoonews_df['Title'] = title_store
    yahoonews_df['Date'] = date_store

    return yahoonews_df
