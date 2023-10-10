import datetime
import requests as re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
import time
import sys
import openpyxl
from openpyxl.styles import PatternFill
#BeautifulSoup のバグの修正
sys.path.append("lib.bs4")

#株価の価格や価格推移の情報を格納するためのリスト
price_list = []
price_change_list = []
company_name_list = []

FILEPATH = "output.csv"

INTERVAL_TIME = 10

#リストにスクレイピングしたい企業の証券コードを入力
COMPANY_CODE_LIST = ["1376","1377", "1333"]


for company_code in COMPANY_CODE_LIST:   

    #webドライバーの起動
    driver = webdriver.Chrome()
    
    #URLを作成
    STOCK_URL = 'https://shikiho.toyokeizai.net/stocks/' + company_code
    driver.get(STOCK_URL)
    url = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(url, "html.parser")
    company_name = soup.find("div", class_="head__main__left__title").text.strip()
    price = soup.find("div", class_="stock-index__price__current").text.strip()
    price_change = soup.find("div", class_="stock-index__price__change").text.strip()

    #リストに情報を格納
    price_list.append(price)
    price_change_list.append(price_change)
    company_name_list.append(company_name)
    
    #スクレイピングの時間を空ける
    print(company_name +":" + price + ":" + price_change)
    time.sleep(INTERVAL_TIME)


df =pd.DataFrame(data = {'会社名':company_name_list,
                            '株価':price_list,
                            '株価推移':price_change_list
                            })
df = df.set_index('会社名')

#日程を取得
date = datetime.datetime.now()

df.to_excel(f'株価リスト_{date.month}月{date.day}日.xlsx', sheet_name='10月')
driver.quit()

#株価リストを編集
book = openpyxl.load_workbook(f'株価リスト_{date.month}月{date.day}日.xlsx')
ws = book['10月']

red_fill = PatternFill(patternType="solid", fgColor="ff7f50")
blue_fill = PatternFill(patternType="solid", fgColor="6495ed")
for rows in ws.iter_rows(min_row=1, min_col=3, max_row=len(company_name_list)+1, max_col=3):
    for cell in rows:
        if "+" in cell.value:
            #セルの背景色を赤くする
            ws[cell.coordinate].fill = red_fill
        elif "-" in cell.value:
            #セルの背景色を青くする
            ws[cell.coordinate].fill = blue_fill
            
book.save(f'株価リスト_{date.month}月{date.day}日.xlsx')