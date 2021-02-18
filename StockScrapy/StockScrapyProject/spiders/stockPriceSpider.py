import scrapy
import csv
import pandas as pd
from datetime import datetime
from ..items import StockPrice_items

class stockPriceSpider(scrapy.Spider):
    items=StockPrice_items()
    Type='CO_PRICE'
    name='stockPriceSpider' #爬蟲名稱
    Year='' #紀錄西元年份
    ROC_Year='' #記錄民國年份
    Month=''
    Day=''
    Date=''
    Co_ids=[]
    CSV_File_PATH=''
    noExist=[]
    twse_SDate=''
    tpex_SDate=''
    is_TWSE_open=bool
    is_TPEX_open=bool
    se_price_urls=[]
    allows_domains=['www.tpex.org.tw','www.twse.com.tw']

    def load_CSV(self):
        with open(self.CSV_File_PATH,newline='',encoding="utf-8")as csvfile_Lc:# 讀入CSV檔
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                    if( (len(row['代號']) ==4) and row['代號'].isnumeric()): #檢查股號是否為純號碼以及是否為4位數
                        self.Co_ids.append(row['代號'])
        return None
    
    def __init__(self,CSV_File_PATH, **kwargs):
        self.CSV_File_PATH=CSV_File_PATH
        self.Day=str(datetime.today().day)
        self.Month=str(f'{datetime.today().month:02d}')
        self.Year=str(datetime.today().year)
        self.ROC_Year=str(datetime.today().year-1911)
        self.Date=datetime.today().strftime("%Y-%m-%d")
        self.twse_SDate=(f'{self.Year}{self.Month}{self.Day}')
        self.tpex_SDate=(f'{self.ROC_Year}/{self.Month}/{self.Day}')
        print(f"西元{self.Year}年{self.Month}月{self.Day}日\n民國{self.ROC_Year}年{self.Month}月{self.Day}日")
        print(f"{self.twse_SDate}\n{self.tpex_SDate}")
        print(self.se_price_urls)
        self.load_CSV()
        super().__init__( **kwargs)
    def write_items(self,response):
        for co_id in self.Co_ids:
            twse_get=response.xpath(f'//td[contains(text(),"{co_id}")]//text()').get()
            tpex_get=response.xpath(f'//td[contains(text(),"{co_id}")]//text()').get()
            if(twse_get != 'None'):
                twse_price=float(response.xpath(f'//td[contains(text(),"{co_id}")]/following-sibling::td[6]//text()').get())
                twse_co_name=response.xpath(f'//td[contains(text(),"{co_id}")]/following-sibling::td[1]//text()').get()
                self.items['CO_ID']=str(co_id)
                self.items['CO_NAME']=str(twse_co_name)
                self.items['Price']=twse_price
            if(tpex_get != 'None'):
                tpex_price=float(response.xpath(f'//td[contains(text(),"{co_id}")]/following-sibling::td[2]//text()').get())
                tpex_co_name=response.xpath(f'//td[contains(text(),"{co_id}")]/following-sibling::td[1]//text()').get()
                self.items['CO_ID']=str(co_id)
                self.items['CO_NAME']=str(tpex_co_name)
                self.items['Price']=tpex_price
            if(twse_get != 'None' or tpex_get != 'None'):
                self.items['SYear']=str(self.Year)
                self.items['SDate']=str(self.Date)
            else:
                self.noExist.append(co_id)

    def parse_TWSE(self,response):
        return_text = str(response.xpath("//div[contains(text(),'很抱歉，沒有符合條件的資料!')]//text()").get())
        print('TWSE='+return_text)
        print(type(return_text))
        if(return_text != 'None'):
            self.is_TWSE_open=False
            pass
        else:
            self.is_TWSE_open=True
            self.write_items(response)
        print('TWSE', 'is Open' if self.is_TWSE_open else 'is Close')
    
    def parse_TPEX(self,response):
        return_text = str(response.xpath("//td[contains(text(),'共0筆')]//text()").get())
        print('TPEX='+return_text)
        print(type(return_text))

        if(return_text != 'None'):
            self.is_TPEX_open=False
            pass
        else:
            self.is_TPEX_open=True
            self.write_items(response)
        print('TPEX', 'is Open' if self.is_TPEX_open else 'is Close')

    def start_requests(self):
        urls=(
            (self.parse_TPEX,f'https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&o=htm&d={self.tpex_SDate}&s=0,asc,0'),
            (self.parse_TWSE,f'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={self.twse_SDate}&type=ALLBUT0999')
        )
        for cb,url in urls:
            yield(scrapy.Request(url,callback=cb))
        return super().start_requests()