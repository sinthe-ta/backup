from selenium import webdriver
import chromedriver_binary
import time
from time import sleep
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
import json
import textwrap
import csv
import datetime
import pptx
import urllib.error
import urllib.request
from pptx.util import Pt
import shutil
import math
import sys

URL = 'https://www.googleapis.com/youtube/v3/'
API_KEY = 'AIzaSyBXlTVtjMZI3eeJZaByEHPlMmWJvBCbvXU'



def get_page_info(url):
    text = ""                              
    options = Options()                    
    options.add_argument('--incognito')    
    #options.add_argument('--headless')      
    driver = webdriver.Chrome(options=options)
    driver.get(url)                         
    driver.implicitly_wait(10)         
    text = driver.page_source
    # 最下部までスクロールして更新されたらtextを変更、変化なしの場合スクロール終了break
    while 1:
        driver.find_element(By.TAG_NAME,'body').send_keys(Keys.END)
        sleep(3)
        text2=driver.page_source
        if text!=text2:
            text=text2
        else:
            break
    driver.quit()                           
    return text     

def get_title(url):
    id_list = []
    text = get_page_info(url)    
    bs = BeautifulSoup(text, features='lxml')  
    title =  bs.select('.titlebody')
    
    return title

def do():
    # ページ
    url = 'http://blog.livedoor.jp/dqnplus/archives/2020629.html'
    get_title(url)



if __name__ == '__main__':
    do()