import pandas as pd
import json
import requests
from bs4 import BeautifulSoup

'''
證交所上市股票代碼
'''

url = 'https://isin.twse.com.tw/isin/C_public.jsp?strMode=2'
headers = {
    'User-Agent':
	'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:139.0) Gecko/20100101 Firefox/139.0'
    }

response = requests.get(url, headers = headers)

data = pd.read_html(response.text)
data = data[0]
data.columns = data.iloc[0] 
data = data.drop([0], axis = 0) 
data['代號'] = data['有價證券代號及名稱'].apply(lambda x: x.split()[0])
data['股票名稱'] = data['有價證券代號及名稱'].apply(lambda x: x.split()[-1])
data['上市日'] = pd.to_datetime(data['上市日'], errors = 'coerce')
data = data.dropna(subset = ['上市日']) 
data = data.drop(['有價證券代號及名稱', '國際證券辨識號碼(ISIN Code)', 'CFICode', '備註'], axis = 1)
data = data[['代號', '股票名稱', '上市日', '市場別', '產業別']]
data = data.dropna(subset = ['產業別'])
data = data[data['代號'].str.isdigit()]

data.to_excel('./TWSE.xlsx')