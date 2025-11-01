
from bs4 import BeautifulSoup
import requests
import datetime as dt
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import gspread
from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException  
from selenium.common.exceptions import TimeoutException
from urllib.error import HTTPError
from requests.exceptions import ConnectionError
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

driver = webdriver.Chrome(options=chrome_options)


spread_api = gspread.service_account(filename="/app/spreadapi-main.json")


spread_sheet = spread_api.open("BTC and Dollars")



spread_api_redundant = gspread.service_account(filename="/app/spreadapi-redundant/spreadapi-redundant.json")

spread_sheet_redunant = spread_api_redundant.open("BTC and Dollars Redundant")

#coin_price
'''
try:      
    driver.get('https://shop.btcegyptgold.com/shop/gold/coins.html?gm=8gm')
    coin_price = driver.find_element(By.XPATH,"//span[@class='price']").text
    print('gold selenium')

except:
    try:
        coin_response = requests.get('https://shop.btcegyptgold.com/shop/gold/coins.html?gm=8gm')

        coin_soup = BeautifulSoup(coin_response.content)

        coin_price = coin_soup.find('span',"price").text
        print('coin bs4')
    except:
        coin_price = 'Closed or Unreachable'
        print('coin closed')
        
print(coin_price) 
'''

#Dollar_price        



try:
    driver.get('https://www.nbe.com.eg/NBE/E/#/EN/ExchangeRatesAndCurrencyConverter')
    us = []
    search = driver.find_elements(By.XPATH,"//td[@class='marker']")
    for i in search:
        us.append(i.text)
    spliting = us[3].split('\n')
    Dollar_price = (spliting[0].split(' '))[1]
    Dollar_price
    print("NBE")
    
    
except:
    try:

        driver.get('https://www.google.com')
        search = driver.find_element(By.XPATH,"//input[@class='gLFyf']")
        search.send_keys('dollar to egp')
        search.send_keys(Keys.ENTER)
        Dollar_price = driver.find_element(By.XPATH,"//span[@class='DFlfde SwHCTb']").text
    
        print('Google')
    except:
        Dollar_price = 'Closed or Unreachable'
        print('dollar Closed')
        
print(Dollar_price)   

#Gold Prices   

try:
    kerat_21_response = requests.get('https://market.isagha.com/prices').content

    kerat_21_soup = BeautifulSoup(kerat_21_response)
    
    kerat_21_span = kerat_21_soup.find_all('div', class_ = 'value')

    kerat = []
    for i in kerat_21_span:
        kerat.append(i.text)

    kerat_24_buy = kerat[0]
    kerat_24_sell = kerat[1]

    kerat_21_buy = kerat[6]
    kerat_21_sell = kerat[7]

    kerat_18_buy = kerat[9]
    kerat_18_sell = kerat[10]
    ounce_dollar = kerat[24].split()[0]
    coin_price = (float(kerat_21_buy)+54)*8
    Dollar_to_egp = float(kerat_24_buy) / (float(ounce_dollar)/31.1)
    Dollar_to_egp=round(Dollar_to_egp,2)
    coin_price = round(coin_price)
    print('Gold BS4')
    
except:
    try: 
        driver.get('https://market.isagha.com/prices')
        search = driver.find_elements(By.XPATH,"//div[@class='value']")
        kerat_price = []
        for i in search:
            kerat_price.append(i.text)
        kerat_24_buy = kerat_price[0]
        kerat_24_sell = kerat_price[1]
        kerat_21_buy = kerat_price[6]
        kerat_21_sell = kerat_price[7]
        kerat_18_buy = kerat_price[9]
        kerat_18_sell = kerat_price[10]
        ounce_dollar = kerat_price[24].split()[0]
        coin_price = (float(kerat_21_buy)+54)*8
        Dollar_to_egp = float(kerat_24_buy) / (float(ounce_dollar)/31.1)
        Dollar_to_egp=round(Dollar_to_egp,2)
        coin_price = round(coin_price)
        print('Selenium')
        
    except:
        kerat_18_sell = 'Closed or Unreachable'
        kerat_21_sell = 'Closed or Unreachable'
        kerat_24_sell = 'Closed or Unreachable'
        kerat_18_buy = 'Closed or Unreachable'
        kerat_21_buy = 'Closed or Unreachable'
        kerat_24_buy = 'Closed or Unreachable'
        coin_price ='Closed or Unreachable'
        Dollar_to_egp='Closed or Unreachable'
        ounce_dollar='Closed or Unreachable'
        print('coin closed')
        
try:        
    driver.get('https://sarf-today.com/currency/us_dollar/market')
    price_list = driver.find_element(By.XPATH,"//div[@class='col-md-8 cur-info-container']").text
    blackmarket = price_list.split('\n')

    avgblackmarket = ((float(blackmarket[3])+float(blackmarket[5]))/2)
except:
    
    avgblackmarket = 'Closed or Unreachable'
    
current_time = dt.datetime.now()
    

data = [current_time.strftime("%Y-%m-%d"), current_time.strftime("%H:%M:%S"), str(coin_price) + ' EGP' , Dollar_price,kerat_18_buy,kerat_21_buy, kerat_24_buy, kerat_18_sell,kerat_21_sell,kerat_24_sell,'Laptop',ounce_dollar,Dollar_to_egp,avgblackmarket]

wks1 = spread_sheet.worksheet('Sheet1')
wks1_redundant = spread_sheet_redunant.worksheet('Sheet1')

wks1.insert_row(values = data , index = 2, value_input_option= 'raw')
wks1_redundant.insert_row(values = data , index = 2, value_input_option= 'raw')

print(data)
wks2= spread_sheet.worksheet('Sheet2')  
wks2.update('A2:N2', [data])
wks2_redundant = spread_sheet_redunant.worksheet('Sheet2')
wks2_redundant.update('A2:N2', [data])
