import requests
import pandas as pd
from bs4 import BeautifulSoup

def get_yahoo_news(stock:str, headers):
    url = f'https://tw.stock.yahoo.com/quote/{stock}/news'

    response = requests.get(url, headers = headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Append news url
    news_links = []
    tags = soup.find_all('a')
    for tag in tags:
        news_link = tag['href']
        if news_link[-4:] == 'html' and 'news' in tag['href'] :
           news_links.append(news_link)

    return news_links

def news_content(link, headers):
    response = requests.get(link, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    title = soup.find('h1').text
    newstime = soup.find('time')
    newstime = newstime.text.split(' ')[0]
    newstime = newstime.replace('年', '/')
    newstime = newstime.replace('月', '/')
    newstime = newstime.replace('日', '')
    return title, newstime

def main(stock, headers):
    news_links = get_yahoo_news(stock, headers)
    titles, publishdates = [], []
    
    for news_url in news_links:
        title, newstime = news_content(news_url, headers)
        titles.append(title)
        publishdates.append(newstime)
    return titles, publishdates, news_links 

if __name__ == '__main__':
    stock = input('Stock code:')
    headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:139.0) Gecko/20100101 Firefox/139.0'
    }
    titles, publishdates, news_links = main(stock, headers)
    # Making a dataframe    
    df = pd.DataFrame()
    df['Title'] = titles
    df['PublishDate'] = publishdates
    df['url'] = news_links
    df.to_excel(f'./{stock}_news.xlsx')
    print(f'The news of {stock} has been stored to {stock}_news.xlsx!')