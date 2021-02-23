import csv
import urllib.parse as urlParse
from datetime import datetime

import winsound
from pandas import DataFrame
import PySimpleGUI as sg
import scrapy
from pydispatch import dispatcher
from scrapy import signals
import time
from ..items import StockPrice_items


class stockPriceSpider(scrapy.Spider):
    Type = '股價資料'
    SubType = ''
    name = 'StockPriceSpider'  # 爬蟲名稱
    Year = ''  # 紀錄西元年份
    ROC_Year = ''  # 記錄民國年份
    Month = ''
    Day = ''
    info = ''
    Date = ''
    Co_ids = []

    # 判別是否為首次運行
    possible_Co_ids_TWSE = []
    possible_Co_ids_TPEX = []
    TWSE_First_Run = True
    TPEX_First_Run = True

    # 讀入參數
    CSV_File_PATH = ''
    noExist = []
    twse_SDate = ''
    tpex_SDate = ''
    # State
    is_TWSE_open = False
    is_TPEX_open = False
    # Info
    se_status = ''
    se_urls = []
    start_urls = []
    allows_domains = ['www.tpex.org.tw', 'www.twse.com.tw']

    def run_tpex(self):
        return scrapy.Request(self.se_urls[0], callback=self.tpex_mining_Data_Parse)

    def run_twse(self):
        return scrapy.Request(self.se_urls[1], callback=self.twse_mining_Data_Parse)

    def start_requests(self):
        for url in self.se_urls:
            yield scrapy.Request(url, callback=self.check_se_parse, priority=5, dont_filter=True)

        #yield scrapy.Request(self.se_urls[1], callback=self.twse_mining_Data_Parse, dont_filter=True)
        return super().start_requests()

    def is_number(self, string):
        try:
            float(string)
            return True
        except ValueError:
            return False

    def output_EmptyList_csv(self):  # 列出未存在股號
        now = datetime.now()
        dt_string = now.strftime("%Y-%m-%d %H-%M-%S")
        if(len(self.noExist)):
            self.noExist = list(filter(None, self.noExist))
            dict = {'代號': self.noExist}
            df = DataFrame(dict)
            filename = f'.\{dt_string}-股價資料-未存在股號.csv'
            df.to_csv(filename, index=False)
            print(f'已匯出未存在的股號至{filename}')
        else:
            print('無缺漏股號。')

    def load_CSV(self):
        with open(self.CSV_File_PATH, newline='', encoding="utf-8")as csvfile_Lc:  # 讀入CSV檔
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                if((len(row['代號']) == 4) and row['代號'].isnumeric()):  # 檢查股號是否為純號碼以及是否為4位數
                    self.Co_ids.append(row['代號'])

    def __init__(self, CSV_File_PATH, **kwargs):
        dispatcher.connect(self.spider_closed,
                           signals.spider_closed)  # 設置爬蟲關閉時的動作
        self.CSV_File_PATH = CSV_File_PATH
        self.Day = str(datetime.today().day-1)
        self.Month = str(f'{datetime.today().month:02d}')
        self.Year = str(datetime.today().year)
        self.ROC_Year = str(datetime.today().year-1911)
        self.Date = datetime.today().strftime("%Y-%m-%d")
        self.twse_SDate = (f'{self.Year}{self.Month}{self.Day}')
        self.tpex_SDate = (f'{self.ROC_Year}/{self.Month}/{self.Day}')
        print(
            f"西元{self.Year}年{self.Month}月{self.Day}日\n民國{self.ROC_Year}年{self.Month}月{self.Day}日")
        print(f"{self.twse_SDate}\n{self.tpex_SDate}")
        self.load_CSV()
        self.se_urls.append(
            f'https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&o=htm&d={self.tpex_SDate}&s=0,asc,0')
        self.se_urls.append(
            f'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={self.twse_SDate}&type=ALLBUT0999')
        super().__init__(**kwargs)

    def spider_closed(self, spider):  # 爬蟲關閉時的動作
        self.output_EmptyList_csv()
        if(self.is_TPEX_open and self.is_TWSE_open):
            winsound.PlaySound("SystemAsterisk",winsound.SND_ALIAS)
        else:
            winsound.PlaySound("SystemQuestion",winsound.SND_ALIAS)
        sg.popup(self.se_status)

    def check_se_parse(self, response):
        title = '正在線上檢查今日證券交換所是否開盤...'
        sg.SystemTray.notify(title, '')
        domain = urlParse.urlparse(response.url).hostname
        print('線上檢查今日證券交換所是否開盤...')
        print(f'網域：{domain}')
        return_text_twse = ''
        return_text_tpex = ''
        if(domain == 'www.twse.com.tw'):
            return_text_twse = str(response.xpath(
                "//div[contains(text(),'很抱歉，沒有符合條件的資料!')]//text()").get())
            print(return_text_twse)
            self.is_TWSE_open = True if (return_text_twse == 'None') else False
        if(domain == 'www.tpex.org.tw'):
            return_text_tpex = str(response.xpath(
                "//td[contains(text(),'共0筆')]//text()").get())
            print(return_text_tpex)
            self.is_TPEX_open = True if (return_text_tpex == 'None') else False

        if(self.is_TPEX_open and self.is_TWSE_open):
            self.se_status = 'TWSE與TPEX皆收盤'
        elif(self.is_TPEX_open and not(self.is_TWSE_open)):
            self.se_status = 'TPEX已收盤，TWSE未收盤'
        elif(not(self.is_TPEX_open) and self.is_TWSE_open):
            self.se_status = 'TPEX未收盤，TWSE已收盤'
        else:
            self.se_status = 'TWSE與TPEX未收盤'
        sg.SystemTray.notify(self.se_status, '')
        if(self.is_TPEX_open and self.is_TPEX_open):
            yield scrapy.Request(self.se_urls[1], callback=self.twse_mining_Data_Parse, dont_filter=True)

    def tpex_mining_Data_Parse(self, response):
        local_Co_ids = []
        if(self.TWSE_First_Run):
            local_Co_ids = self.Co_ids
        else:
            local_Co_ids = self.possible_Co_ids_TPEX
        for data in response.xpath('body'):
            domain = urlParse.urlparse(response.url).hostname
            print('First RUN:', self.TPEX_First_Run)
            print('爬取開始')
            print(f'網域：{domain}')
            print(local_Co_ids)

            for co_id in local_Co_ids:
                items = StockPrice_items()
                print('First RUN:', self.TWSE_First_Run)
                print(co_id)
                tpex_get = ''
                tpex_get = str(response.xpath(
                    f'//td[text()="{co_id}"]//text()').get())
                tpex_co_name = str(data.xpath(
                    f'//td[text()="{co_id}"]/following-sibling::td[1]//text()').get())
                print(tpex_get)
                if(tpex_get == 'None' and (not tpex_co_name.isnumeric())):
                    if(self.TWSE_First_Run):
                        print(f'股號 {co_id} 不存在於交易所，可能為TWSE的股號，丟入至暫存中...')
                        self.possible_Co_ids_TWSE.append(co_id)
                        continue
                    else:
                        print(f'股號 {co_id} 不存在兩邊交易所，丟入到未存在股號中...')
                        self.noExist.append(co_id)
                        continue
                else:
                    print('TPEX GET ITEMS')
                    tpex_price = str(data.xpath(
                        f'//td[text()="{co_id}"]/following-sibling::td[2]//text()').get())
                    tpex_price = tpex_price.replace(',', '')
                    if(self.is_number(tpex_price)):
                        tpex_price = float(tpex_price)
                    else:
                        tpex_price = None
                    items['CO_ID'] = str(co_id)
                    items['CO_SHORT_NAME'] = str(tpex_co_name)
                    items['Price'] = tpex_price
                    items['SUB_DATA_TYPE'] = 'TPEX'
                    items['SYear'] = str(self.Year)
                    items['SDate'] = str(self.Date)
                    items['DATA_TYPE'] = self.Type
                    yield(items)
        self.TPEX_First_Run = False

    def twse_mining_Data_Parse(self, response):
        if(not(self.is_TPEX_open and self.is_TWSE_open)):
            print(self.se_status)
            pass
        else:
            local_Co_ids = []
            if(self.TPEX_First_Run):
                local_Co_ids = self.Co_ids
            else:
                local_Co_ids = self.possible_Co_ids_TWSE
            for data in response.xpath('body'):
                domain = urlParse.urlparse(response.url).hostname
                print('First RUN:', self.TWSE_First_Run)
                print('爬取開始')
                print(f'網域：{domain}')

                for co_id in local_Co_ids:
                    items = StockPrice_items()
                    print('First RUN:', self.TWSE_First_Run)
                    print(co_id)
                    twse_get = ''
                    twse_get = str(response.xpath(
                        f'//td[text()="{co_id}"]//text()').get())
                    twse_co_name = str(data.xpath(
                        f'//td[text()="{co_id}"]/following-sibling::td[1]//text()').get())
                    print(twse_get)
                    if(twse_get == 'None' and (not twse_co_name.isnumeric())):
                        if(self.TPEX_First_Run):
                            print(f'股號 {co_id} 不存在於交易所，可能為TPEX的股號，丟入至暫存中...')
                            self.possible_Co_ids_TPEX.append(co_id)
                            continue
                        else:
                            print(f'股號 {co_id} 不存在兩邊交易所，丟入到未存在股號中...')
                            self.noExist.append(co_id)
                            continue
                    else:
                        print('TWSE GET ITEMS')
                        twse_price = str(data.xpath(
                            f'//td[text()="{co_id}"]/following-sibling::td[6]//text()').get())
                        twse_price = twse_price.replace(',', '')
                        print(twse_price)
                        if (self.is_number(twse_price)):
                            twse_price = float(twse_price)
                        else:
                            twse_price = None
                        items['CO_ID'] = str(co_id)
                        items['CO_SHORT_NAME'] = str(twse_co_name)
                        items['Price'] = twse_price
                        items['SUB_DATA_TYPE'] = 'TWSE'
                        items['SYear'] = str(self.Year)
                        items['SDate'] = str(self.Date)
                        items['DATA_TYPE'] = self.Type
                        yield(items)
            self.TWSE_First_Run = False
            yield scrapy.Request(self.se_urls[0], callback=self.tpex_mining_Data_Parse, dont_filter=True)
