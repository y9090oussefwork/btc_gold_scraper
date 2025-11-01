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
from selenium.webdriver.chrome.options import Options
import json
import base64

# -------------------------------
# إعداد gspread من Environment Variable
# -------------------------------
encoded_json = os.getenv("SPREAD_API_JSON_B64")
if not encoded_json:
    raise ValueError("الـ Environment Variable SPREAD_API_JSON_B64 غير موجود!")

json_bytes = base64.b64decode(encoded_json)
json_dict = json.loads(json_bytes)

spread_api = gspread.service_account_from_dict(json_dict)
spread_sheet = spread_api.open("BTC and Dollars")

# إذا عندك نسخة Redundant
spread_sheet_redunant = spread_api.open("BTC and Dollars Redundant")

# -------------------------------
# إعداد Selenium
# -------------------------------
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

# -------------------------------
# Scraping Prices
# -------------------------------

# Dollar Price
try:
    driver.get('https://www.nbe.com.eg/NBE/E/#/EN/ExchangeRatesAndCurrencyConverter')
    us = [i.text for i in driver.find_elements(By.XPATH,"//td[@class='marker']")]
    spliting = us[3].split('\n')
    Dollar_price = (spliting[0].split(' '))[1]
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
        print('Dollar Closed')

print(Dollar_price)

# Gold Prices
try:
    kerat_21_response = requests.get('https://market.isagha.com/prices').content
    kerat_21_soup = BeautifulSoup(kerat_21_response)
    kerat_21_span = kerat_21_soup.find_all('div', class_='value')

    kerat = [i.text for i in kerat_21_span]

    kerat_24_buy = kerat[0]
    kerat_24_sell = kerat[1]
    kerat_21_buy = kerat[6]
    kerat_21_sell = kerat[7]
    kerat_18_buy = kerat[9]
    kerat_18_sell = kerat[10]
    ounce_dollar = kerat[24].split()[0]

    coin_price = round((float(kerat_21_buy)+54)*8)
    Dollar_to_egp = round(float(kerat_24_buy) / (float(ounce_dollar)/31.1), 2)

    print('Gold BS4')

except:
    kerat_18_sell = kerat_21_sell = kerat_24_sell = 'Closed or Unreachable'
    kerat_18_buy = kerat_21_buy = kerat_24_buy = 'Closed or Unreachable'
    coin_price = Dollar_to_egp = ounce_dollar = 'Closed or Unreachable'
    print('Gold Closed')

# Black Market Dollar
try:
    driver.get('https://sarf-today.com/currency/us_dollar/market')
    price_list = driver.find_element(By.XPATH,"//div[@class='col-md-8 cur-info-container']").text
    blackmarket = price_list.split('\n')
    avgblackmarket = ((float(blackmarket[3])+float(blackmarket[5]))/2)
except:
    avgblackmarket = 'Closed or Unreachable'

# -------------------------------
# تحضير البيانات للـ Google Sheet
# -------------------------------
current_time = dt.datetime.now()
data = [
    current_time.strftime("%Y-%m-%d"),
    current_time.strftime("%H:%M:%S"),
    str(coin_price) + ' EGP',
    Dollar_price,
    kerat_18_buy,
    kerat_21_buy,
    kerat_24_buy,
    kerat_18_sell,
    kerat_21_sell,
    kerat_24_sell,
    'Laptop',
    ounce_dollar,
    Dollar_to_egp,
    avgblackmarket
]

# تحديث الـ Sheets
wks1 = spread_sheet.worksheet('Sheet1')
wks1.insert_row(values=data, index=2, value_input_option='RAW')

wks2 = spread_sheet.worksheet('Sheet2')
wks2.update('A2:N2', [data])

# لو عندك نسخة Redundant
wks1_redundant = spread_sheet_redunant.worksheet('Sheet1')
wks1_redundant.insert_row(values=data, index=2, value_input_option='RAW')

wks2_redundant = spread_sheet_redunant.worksheet('Sheet2')
wks2_redundant.update('A2:N2', [data])

print("تم تحديث Google Sheets بنجاح:", data)
