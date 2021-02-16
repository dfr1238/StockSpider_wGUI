import scrapy
import csv
from scrapy import signals
import urllib.parse as urlParse
from urllib.parse import parse_qs
import pandas as pd 
from ..items import StockscrapyprojectItem
from pydispatch import dispatcher
from datetime import datetime

class StockSpider(scrapy.Spider):
    Year=''
    Season=''
    name = "StockSpider" #爬蟲名稱。
    start_urls=[]
    noExist=[]
    wait_url_A=0
    import_csv='..\上市_urlA.csv'
    total=0
    ready_crawl=0
    exist=0
    current=1
    allowed_domains =['mops.twse.com.tw'] # 允許網域

    def load_csv(self):
         with open(self.import_csv,newline='',encoding="utf-8") as csvfile_Lc: #讀入CSV檔案
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                self.total+=1
                if(row['代號'].isnumeric() & (len(row['代號']) !=6)): #檢查股號是否為純號碼以及是否不是6位數
                    self.ready_crawl+=1
                    Co_id=row['代號']
                    self.start_urls.append(f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={Co_id}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=C') #帶入網址序列

    def print_info(self): #列出爬蟲資訊
        if (self.current > self.ready_crawl):
            print("超額運算中，開始處理A類RID") 
        else:
            print("正常運算當中")
        print(f"CSV總筆數:{self.total},匯入有效筆數:{self.ready_crawl},目前筆數:{self.current},確認存在股數:{self.exist},確認未存在股號數:{len(self.noExist)},待導入A類查尋筆數:{self.wait_url_A}")
        print("\n未存在股號列表：")
        for printdata in range(len(self.noExist)):
            print(self.noExist[printdata])

    def output_EmptyList_csv(self):
        now = datetime.now()
        dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
        if(not len(self.noExist)):
            self.noExist = list(filter(None, self.noExist))
            dict ={'代號' : self.noExist}
            df = pd.DataFrame(dict)
            filename=f'..\{dt_string}-未存在股號.csv'
            df.to_csv(filename, index=False)
            print(f'已匯出未存在的股號至{filename}')
        else:
            print('無缺漏股號。')

    def spider_closed(self, spider): #爬蟲關閉時的動作
        self.output_EmptyList_csv()
    
    def __init__(self,Year='',Season='',CSV='', **kwargs):
        dispatcher.connect(self.spider_closed, signals.spider_closed) #設置爬蟲關閉時的動作
        self.import_csv=CSV #匯入CSV之路徑 -a CSV 'Path'
        self.Year=Year #帶入參數年份 -a Year
        self.Season=Season #帶入參數季度 -a Season
        self.load_csv()
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
        company_Id=parse_qs(parsed.query)['CO_ID'] #獲取網址股號
        report_ID=parse_qs(parsed.query)['REPORT_ID'] #獲取回報ID
        if(response.xpath("/html/body/h4//text()").get() is None): #檢查是否存在檔案不存在之字串
            self.exist+=1
            items['ID'] = self.exist
            if(str(report_ID[0])=='A'): #如果是回報A則減少待導入尋找筆數
                self.wait_url_A-=1
        else:
            if(str(report_ID[0])=='A'): #如果是回報A則記錄為無資料股
                print("該股A與C類皆無資料，記錄至未存在表中。")
                self.noExist.append(str(company_Id[0]))
                self.wait_url_A-=1
                pass
            else: #試圖用回報A連結重新爬取
                print("該股類型C無資料，轉入類型A查資料。")
                self.wait_url_A+=1
                self.start_urls.append(f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={company_Id[0]}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=A')
                pass

        items['CO_ID'] = str(company_Id[0])
        co_name = str(response.xpath('/html/body/div[2]/div[1]/div[2]/span[1]//text()').get())
        items['CO_NAME'] = co_name
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
                        items[tables1_ItemsName[tableID]] = -(float(data[1]))
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
                        items[tables2_ItemsName[tableID]] = -(float(data[1]))
                else:
                    items[tables2_ItemsName[tableID]] = None
            
            yield(items)
        self.print_info()
        self.current+=1