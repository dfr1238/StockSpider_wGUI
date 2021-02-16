import scrapy
import csv
import urllib.parse as urlParse
from urllib.parse import parse_qs
from decimal import *
import pandas as pd

from ..items import StockscrapyprojectItem

class StockSpider(scrapy.Spider):
    name = "StockSpider" #爬蟲名稱。
    start_urls=[]
    noExist=[]
    import_csv='..\無股測試.csv'
    total=0
    ready_crawl=0
    exist=0
    current=1
    allowed_domains =['mops.twse.com.tw'] # 允許網域

    def print_info(self): #列出爬蟲資訊
        print(f"CSV總筆數:{self.total},匯入有效筆數:{self.ready_crawl},目前筆數:{self.current},確認存在股數:{self.exist},確認未存在股號數:{len(self.noExist)}")
        print("\n未存在股號列表：")
        for printdata in range(len(self.noExist)):
            print(self.noExist[printdata])

    def __init__(self,Year='',Season='', **kwargs):
        self.Year=Year #帶入參數年份 -a Year
        self.Season=Season #帶入參數季度 -a Season
        with open(self.import_csv,newline='',encoding="utf-8") as csvfile_Lc: #讀入CSV檔案
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                self.total+=1
                if(row['代號'].isnumeric()): #檢查股號是否為純號碼
                    self.ready_crawl+=1
                    Co_id=row['代號']
                    self.start_urls.append(f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={Co_id}&SYEAR={Year}&SSEASON={Season}&REPORT_ID=C') #帶入網址序列
        super().__init__(**kwargs)  # python3

    def is_Number(self,s): #檢查字串是否為數目
        try:
            float(s)
            return True
        except ValueError:
            return False

    def parse(self, response):
        items = StockscrapyprojectItem()# 匯入資料集。
        parsed = urlParse.urlparse(response.request.url)
        company_id=parse_qs(parsed.query)['CO_ID']
        if(response.xpath("/html/body/h4//text()").get() is None): #檢查是否存在檔案不存在之字串
            self.exist+=1
            items['ID'] = self.exist
        else:
            self.noExist.append(str(company_id))

        items['CO_ID'] = company_id
        items['Syear'] = self.Year
        items['SSeason'] = self.Season
        #主要爬蟲區
        for datas in response.xpath('body/div[2]/div[3]'):
            tables1_ID=['1100','1110','1120','1136','1139','25XX','3110']
            tables1_ItemsName=['A1','A2','A3','A4','A5','A6','A7']
            tables2_ID=['4000','6900','7000','9850']
            tables2_ItemsName=['B1','B2','B3','B4']
            print(tables1_ItemsName[0])

            for tableID in range(0,len(tables1_ID)): #表一獲取資料
                data = datas.xpath(f"//td[contains(text(),'{tables1_ID[tableID]}')]/following-sibling::td[2]//text()").getall()
                if(len(data)):
                    data[0] = data[0].replace(',','')
                    if(self.is_Number(data[0])):
                        items[tables1_ItemsName[tableID]] = float(data[0])
                    else:
                        data[1] = data[1].replace(',','')
                        items[tables1_ItemsName[tableID]] = float(data[1])
                else:
                    items[tables1_ItemsName[tableID]] = None
            
            for tableID in range(0,len(tables2_ID)): #表二獲取資料
                data = datas.xpath(f"//td[contains(text(),'{tables2_ID[tableID]}')]/following-sibling::td[2]//text()").getall()
                if(len(data)):
                    data[0] = data[0].replace(',','')
                    if(self.is_Number(data[0])):
                        items[tables2_ItemsName[tableID]] = float(data[0])
                    else:
                        data[1] = data[1].replace(',','')
                        items[tables2_ItemsName[tableID]] = float(data[1])
                else:
                    items[tables2_ItemsName[tableID]] = None
            
            yield(items)
        self.print_info()
        self.current+=1