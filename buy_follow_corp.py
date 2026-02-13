import sys
import utility_functions as uf
import datetime
import os
import pandas as pd
import traceback
import os
from dotenv import load_dotenv

load_dotenv('.env')
receiver1 = os.getenv('RECEIVER1')
receiver2 = os.getenv('RECEIVER2')

try :

    today = datetime.date.today()
    if uf.is_open(today) == 'Close':
        arr_list = [receiver1, receiver2]
        title = '三大法人買賣超'
        body = f'{today} 今日未開盤！'
        mode=''
        file_path = ['./sugar.jpg']
        file_name = ['sugar.jpg']
        uf.sendmail(arr_list, title, body, mode, file_path, file_name)

    else:
        result = []
        control = 0 
        for i in range(0, 10):
            if control <= 2:        
                date_target = today + datetime.timedelta(days=-int(i)-1)
                if_trade = uf.is_open(date_target)
                if if_trade == "Close":
                    continue
                else:
                    convert_date = date_target.strftime('%Y%m%d')
                    print(convert_date)
                    data = uf.twse_data(convert_date, type='TIB')
                    data.to_excel(f'{convert_date}_twse.xlsx')
                    data = pd.read_excel(f'{convert_date}_twse.xlsx', thousands=',')
                    os.remove(f'{convert_date}_twse.xlsx')   

                    d_s = data[data['三大法人買賣超股數'] > 0]
                    d_s = d_s[:50]
                    if control == 0:
                        result = set(d_s['證券代號'].tolist())
                    else:
                        result = result.intersection(d_s['證券代號'].tolist())

                    control += 1  
                    
            else:
                break    

        arr_list = [receiver1, receiver2]
        title = '三大法人買賣超'
        body = f'目標股票 {result} 連續三日法人買超'
        mode = ''
        file_path = ''
        file_name = ''
        uf.sendmail(arr_list, title, body, mode, file_path, file_name)

except SystemError:
    print("It's ok !")
except:
    arr_list = [receiver1]
    title = '三大法人買賣超異常回報！'
    body = f'異常情況：\n{traceback.format_exc()}'
    mode = ''
    file_path = ''
    file_name = ''
    uf.sendmail(arr_list, title, body, mode, file_path, file_name)

