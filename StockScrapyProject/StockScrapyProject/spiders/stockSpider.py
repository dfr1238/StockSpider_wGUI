import csv
import logging
import urllib.parse as urlParse
from datetime import datetime
from urllib.parse import parse_qs

import pandas as pd
import PySimpleGUI as sg
import scrapy
from pydispatch import dispatcher
from scrapy import signals

from ..items import StockSpider_items


class StockSpider(scrapy.Spider):
    Type = '財務報告'
    SubType = ''
    Year = ''
    info = ''
    Season = ''
    Mode = ''
    name = 'StockSpider'  # 爬蟲名稱。
    start_urls = []
    noExist = []
    wait_url_A = 0
    import_csv = '..\上市_urlA.csv'
    total = 0
    ready_crawl = 0
    exist = 0
    current = 1
    allowed_domains = ['mops.twse.com.tw']  # 允許網域

    def auto_Mode(self):  # 自動模式
        with open(self.import_csv, newline='', encoding="utf-8") as csvfile_Lc:  # 讀入CSV檔案
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                self.total += 1
                if((len(row['代號']) == 4) and row['代號'].isnumeric()):  # 檢查股號是否為純號碼以及是否為4位數
                    self.ready_crawl += 1
                    Co_id = row['代號']
                    self.start_urls.append(
                        f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={Co_id}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=C')  # 帶入網址序列

    def print_info(self):  # 列出爬蟲資訊
        title = ''
        if (self.current > self.ready_crawl):
            title = "超額運算中，開始處理A類RID"
            logging.info("超額運算中，開始處理A類RID")
        else:
            title = "正常運算當中"
            logging.info("正常運算當中")
        self.info = (
            f"CSV總筆數:{self.total},匯入有效筆數:{self.ready_crawl}\n目前筆數:{self.current},確認存在股數:{self.exist}\n確認未存在股號數:{len(self.noExist)},待導入A類查尋筆數:{self.wait_url_A}")
        if(self.current % 15 == 0 or self.current == 1):
            sg.SystemTray.notify(title, self.info)
        logging.info(self.info)
        print("\n未存在股號列表：")
        for printdata in range(len(self.noExist)):
            print(self.noExist[printdata])

    def manual_Mode(self, CO_ID):  # 手動模式
        if((len(CO_ID) == 4) and CO_ID.isnumeric()):
            self.exist += 1
            self.start_urls.append(
                f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={CO_ID}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=C')  # 帶入網址序列
            self.start_urls.append(
                f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={CO_ID}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=A')  # 帶入網址序列
        else:
            logging.error('請輸入正確的四位數純數字股號')
            pass

    def output_EmptyList_csv(self):  # 列出未存在股號
        now = datetime.now()
        dt_string = now.strftime("%Y-%m-%d %H-%M-%S")
        if(len(self.noExist)):
            self.noExist = list(filter(None, self.noExist))
            dict = {'代號': self.noExist}
            df = pd.DataFrame(dict)
            filename = f'.\{dt_string}-財務報告-未存在股號.csv'
            df.to_csv(filename, index=False)
            sg.SystemTray.notify(f'已匯出未存在的股號至\n{filename}')
        else:
            logging.info('無缺漏股號。')

    def spider_closed(self, spider):  # 爬蟲關閉時的動作
        self.output_EmptyList_csv()
        sg.popup(self.info)

    def __init__(self, Year='', Season='', CSV='', Mode='', CO_ID='', **kwargs):  # 初始化動作
        dispatcher.connect(self.spider_closed,
                           signals.spider_closed)  # 設置爬蟲關閉時的動作
        self.Year = Year  # 帶入參數年份 -a Year 數字字串
        self.Season = Season  # 帶入參數季度 -a Season 數字字串
        self.Mode = Mode  # 帶入爬蟲模式 -a Mode 文字字串，Auto與Manual模式
        if(Mode == 'Auto' or Mode == 'A'):
            sg.SystemTray.notify('財務報告爬蟲－初始化', '以批次模式進行中...')
            logging.info(f'目前輸入的參數，年份：{Year}、季度{Season}、模式：{Mode}、CSV路徑:{CSV}')
            self.import_csv = CSV  # 匯入CSV之路徑 -a CSV 'Path'
            self.auto_Mode()
        elif (Mode == 'Manual' or Mode == 'M'):
            sg.SystemTray.notify('財務報告爬蟲－初始化', '以單筆模式進行中...')
            logging.info(f'目前輸入的參數，年份：{Year}、季度{Season}、模式：{Mode}、股號:{CO_ID}')
            self.manual_Mode(CO_ID)
        else:
            logging.error(
                "請輸入正確的抓取參數-a Mode=[參數]\n參數\n 自動模式：Auto或A，加上-a CSV=[檔案路徑]\n手動輸入股號模式：Manual或M，加上-CO_ID[股號]")
        super().__init__(**kwargs)  # python3

    def is_Number(self, s):  # 檢查字串是否為數目
        try:
            float(s)
            return True
        except ValueError:
            return False

    def get_From_Table(self, items, response, tables_ID, tables_ItemsName):  # 從表格當中獲取資料
        for datas in response.xpath('body/div[2]/div[3]'):
            for tableID in range(0, len(tables_ID)):  # 表一獲取資料
                data = datas.xpath(
                    f"//td[contains(text(),'{tables_ID[tableID]}')]/following-sibling::td[2]//text()").getall()
                if(len(data)):
                    data[0] = data[0].replace(',', '')
                    if(self.is_Number(data[0])):
                        items[tables_ItemsName[tableID]] = float(data[0])
                    else:
                        data[1] = data[1].replace(',', '')
                        items[tables_ItemsName[tableID]] = -(float(data[1]))
                else:
                    items[tables_ItemsName[tableID]] = None

    def parse(self, response):  # 擷取開始
        _page_exist = True
        items = StockSpider_items()  # 匯入資料集。
        parsed = urlParse.urlparse(response.request.url)
        company_Id = parse_qs(parsed.query)['CO_ID']  # 獲取網址股號
        report_ID = parse_qs(parsed.query)['REPORT_ID']  # 獲取回報ID

        if(response.xpath("/html/body/h4//text()").get() is None):  # 檢查是否存在檔案不存在之字串
            self.exist += 1
            if(str(report_ID[0]) == 'A'):  # 如果是回報A則減少待導入尋找筆數
                self.wait_url_A -= 1
        else:
            if(str(report_ID[0]) == 'A'):  # 如果是回報A則記錄為無資料股
                _page_exist = False
                logging.error("該股A與C類皆無資料，記錄至未存在表中。")
                self.noExist.append(str(company_Id[0]))
                self.wait_url_A -= 1
                pass
            else:  # 試圖用回報A連結重新爬取
                _page_exist = False
                logging.info("該股類型C無資料，轉入類型A查資料。")
                self.wait_url_A += 1
                self.start_urls.append(
                    f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={company_Id[0]}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=A')
                pass
        if(_page_exist):
            items['DATA_TYPE'] = self.Type
            items['SUB_DATA_TYPE'] = '個別財務報告' if (
                report_ID == 'A') else '合併財務報告'
            items['CO_ID'] = str(company_Id[0])
            co_name = str(response.xpath(
                '/html/body/div[2]/div[1]/div[2]/span[1]//text()').get())
            items['CO_FULL_NAME'] = co_name
            items['Syear'] = self.Year
            items['SSeason'] = self.Season
            # 主要爬蟲區
            tables1_ID = ['1100', '1110', '1120',
                          '1136', '1139', '25XX', '3110']
            tables1_ItemsName = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7']
            tables2_ID = ['4000', '6900', '7000', '9850']
            tables2_ItemsName = ['B1', 'B2', 'B3', 'B4']
            self.get_From_Table(items, response, tables1_ID, tables1_ItemsName)
            self.get_From_Table(items, response, tables2_ID, tables2_ItemsName)
            yield(items)
        self.print_info()
        self.current += 1
