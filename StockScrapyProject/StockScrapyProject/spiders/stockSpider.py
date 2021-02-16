import scrapy
import csv
import urllib.parse as urlParse
from urllib.parse import parse_qs
from decimal import *
import pandas as pd

from scrapy import item
from ..items import StockscrapyprojectItem

class StockSpider(scrapy.Spider):
    name = "StockSpider" #爬蟲名稱。
    start_urls=[]
    allowed_domains =['mops.twse.com.tw'] # 允許網域
    def __init__(self,Year='',Season='', **kwargs):
        self.Year=Year #帶入參數年份 -a Year
        self.Season=Season #帶入參數季度 -a Season
        with open('..\上市_100筆test.csv',newline='',encoding="utf-8") as csvfile_Lc: #讀入CSV檔案
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                if(row['代號'].isnumeric()): #檢查股號是否為純號碼
                    Co_id=row['代號']
                    self.start_urls.append(f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={Co_id}&SYEAR={Year}&SSEASON={Season}&REPORT_ID=C') #帶入網址序列
        super().__init__(**kwargs)  # python3

    def is_Number(self,s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def parse(self, response):
        items = StockscrapyprojectItem()# 匯入資料集。
        parsed = urlParse.urlparse(response.request.url)
        items['CO_ID'] = parse_qs(parsed.query)['CO_ID']
        items['Syear'] = self.Year
        items['SSeason'] = self.Season
        ##表一：資產負債表
        for datas in response.xpath('body/div[2]/div[3]'):
            tables1_ID=['1100','1110','1120','1136','1139','25XX','3110']
            tables1_ItemsName=['A1','A2','A3','A4','A5','A6','A7']
            tables2_ID=['4000','6900','7000','9850']
            tables2_ItemsName=['B1','B2','B3','B4']

            for tableID in len(tables1_ID):
                data = datas.xpath(f"//td[contains(text(),{tables1_ID[tableID]})]/following-sibling::td[2]//text()").getall()
                if(len(data)):
                    data[0] = data[0].replace(',','')
                    if(self.is_Number(data[0])):
                        items[{tables1_ItemsName[tableID]}] = float(data[0])
                    else:
                        data[1] = data[1].replace(',','')
                        items[{tables1_ItemsName[tableID]}] = float(data[1])
                else:
                    items[{tables1_ItemsName[tableID]}] = None
                    
            ##表一：資產負債表
            data_A1 = datas.xpath("//td[contains(text(),'1100')]/following-sibling::td[2]//text()").getall() #A1:現金及約當現金
            
            if(len(data_A1)):
                data_A1[0] = data_A1[0].replace(',','')
                if(self.is_Number(data_A1[0])):
                    items['A1'] = float(data_A1[0])
                else:
                    data_A1[1] = data_A1[1].replace(',','')
                    items['A1'] = float(data_A1[1])
            else:
                items['A1'] = None
            
            data_A2 = datas.xpath("//td[contains(text(),'1110')]/following-sibling::td[2]//text()").getall() #A2:透過損益按公允價值衡量之金融資產－流動

            if(len(data_A2)):
                data_A2[0] = data_A2[0].replace(',','')
                if(self.is_Number(data_A2[0])):
                    items['A2'] = float(data_A2[0])
                else:
                    data_A2[1] = data_A2[1].replace(',','')
                    items['A2'] = float(data_A2[1])
            else:
                items['A2'] = None
            
            data_A3 = datas.xpath("//td[contains(text(),'1120')]/following-sibling::td[2]//text()").getall() #A3:透過其他綜合損益按公允價值衡量之金融資產－流動
            
            if(len(data_A3)):
                data_A3[0] = data_A3[0].replace(',','')
                if(self.is_Number(data_A3[0])):
                    items['A3'] = float(data_A3[0])
                else:
                    data_A3[1] = data_A3[1].replace(',','')
                    items['A3'] = float(data_A3[1])
            else:
                items['A3'] = None
            
            data_A4 = datas.xpath("//td[contains(text(),'1136')]/following-sibling::td[2]//text()").getall() #A4:按攤銷後成本衡量之金融資產－流動
            
            if(len(data_A4)):
                data_A4[0] = data_A4[0].replace(',','')
                if(self.is_Number(data_A4[0])):
                    items['A4'] = float(data_A4[0])
                else:
                    items['A4'] = float(data_A4[1])
            else:
                items['A4'] = None
            
            data_A5 = datas.xpath("//td[contains(text(),'1139')]/following-sibling::td[2]//text()").getall() #A5:避險之金融資產－流動
            
            if(len(data_A5)):
                data_A5[0] = data_A5[0].replace(',','')
                if(self.is_Number(data_A5[0])):
                    items['A5'] = float(data_A5[0])
                else:
                    items['A5'] = float(data_A5[1])
            else:
                items['A5'] = None

            data_A6 = datas.xpath("//td[contains(text(),'25XX')]/following-sibling::td[2]//text()").getall() #A6:非流動負債合計
            
            if(len(data_A6)):
                data_A6[0] = data_A6[0].replace(',','')
                if(self.is_Number(data_A6[0])):
                    items['A6'] = float(data_A6[0])
                else:
                    items['A6'] = float(data_A6[1])
            else:
                items['A6'] = None

            data_A7 = datas.xpath("//td[contains(text(),'3110')]/following-sibling::td[2]//text()").getall() #A7:普通股股本
            
            if(len(data_A7)):
                data_A7[0] = data_A7[0].replace(',','')
                if(self.is_Number(data_A7[0])):
                    items['A7'] = float(data_A7[0])
                else:
                    items['A7'] = float(data_A7[1])
            else:
                items['A7'] = None
            #表二：綜合損益表
            data_B1 = datas.xpath("//td[contains(text(),'4000')]/following-sibling::td[2]//text()").getall() #B1:營業收入合計
            
            if(len(data_B1)):
                data_B1[0] = data_B1[0].replace(',','')
                if(self.is_Number(data_B1[0])):
                    items['B1'] = float(data_B1[0])
                else:
                    items['B1'] = float(data_B1[1])
            else:
                items['B1'] = None

            data_B2 = datas.xpath("//td[contains(text(),'6900')]/following-sibling::td[2]//text()").getall() #B2:營業利益（損失）
            
            if(len(data_B2)):
                data_B2[0] = data_B2[0].replace(',','')
                if(self.is_Number(data_B2[0])):
                    items['B2'] = float(data_B2[0])
                else:
                    items['B2'] = float(data_B2[1])
            else:
                items['B2'] = None

            data_B3 = datas.xpath("//td[contains(text(),'7000')]/following-sibling::td[2]//text()").getall() #B3:營業外收入及支出合計
            
            if(len(data_B3)):
                data_B3[0] = data_B3[0].replace(',','')
                if(self.is_Number(data_B3[0])):
                    items['B3'] = float(data_B3[0])
                else:
                    items['B3'] = float(data_B3[1])
            else:
                items['B3'] = None

            data_B4 = datas.xpath("//td[contains(text(),'9850')]/following-sibling::td[2]//text()").getall() #B4:稀釋每股盈餘合計
            
            if(len(data_B4)):
                data_B4[0] = data_B4[0].replace(',','')
                if(self.is_Number(data_B4[0])):
                    items['B4'] = float(data_B4[0])
                else:
                    items['B4'] = float(data_B4[1])
            else:
                items['B4'] = None
            yield(items)