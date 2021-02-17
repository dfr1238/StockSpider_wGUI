import scrapy
import csv
import pandas as pd
from datetime import datetime

class stockPriceSpider(scrapy.Spider):
    name='stockPriceSpider' #爬蟲名稱
    Year='' #紀錄西元年份
    ROC_Year='' #記錄民國年份
    Month=''
    Day=''
    Date=''
    Co_id=''
    twse_SDate=''
    tpex_SDate=''
    start_urls=[]
    allows_domains=['www.tpex.org.tw','www.twse.com.tw']

    def __init__(self, **kwargs):
        self.Day=str(datetime.today().day)
        self.Month=str(datetime.today().month)
        self.Year=str(datetime.today().year)
        self.ROC_Year=str(datetime.today().year-1911)
        self.Date=datetime.today().strftime("%Y-%m-%d")
        self.twse_SDate=(f'{self.Year}{self.Month}{self.Day}')
        self.tpex_SDate=(f'{self.ROC_Year}/{self.Month}/{self.Day}')
        print(f"西元{self.Year}年{self.Month}月{self.Day}日\n民國{self.ROC_Year}年{self.Month}月{self.Day}日")
        print(f"{self.twse_SDate}\n{self.tpex_SDate}")
        super().__init__( **kwargs)