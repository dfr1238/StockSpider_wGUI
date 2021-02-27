import csv
import logging
import urllib.parse as urlParse
import winsound
from datetime import datetime
from urllib.parse import parse_qs

import PySimpleGUI as sg
import scrapy
from pandas import DataFrame
from pydispatch import dispatcher
from scrapy import signals
from scrapy.exceptions import CloseSpider

from ..items import StockSpider_items


class StockSpider(scrapy.Spider):
    Type = '財務報告'
    SubType = ''
    Year = ''
    info = ''
    Season = ''
    Mode = ''
    wait_url_A_MAX=0
    name = 'StockSpider'  # 爬蟲名稱。
    start_urls = []
    noExist = []
    wait_url_A = 0
    cant_reach = []
    import_csv = '..\上市_urlA.csv'
    total = 0
    ready_crawl = 0
    exist = 0
    current = 1
    not_mamual_cancel = True
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
            f"CSV總筆數:{self.total},匯入有效筆數:{self.ready_crawl}\n目前筆數:{self.current},確認存在股數:{self.exist}\n確認未存在股號數:{len(self.noExist)},待導入A類查尋筆數:{self.wait_url_A},爬取失敗的筆數:{len(self.cant_reach)}\n未存在的列表\n{self.noExist}\n爬取失敗的列表\n{self.cant_reach}")
        self.not_manual_cancel = sg.one_line_progress_meter('目前爬取進度',self.current,self.ready_crawl,'Stock','運行時請勿點擊視窗，顯示沒有回應請勿關閉，為正常現象。\nElapsed Time 為已運行時間\nTime Remaining 為剩餘時間\nEstimated Total Time 為估計完成時間',no_titlebar=False,orientation='h')
        if(not self.not_manual_cancel and self.current < self.exist+self.wait_url_A):
            Button = sg.popup_yes_no('是否取消？','取消爬取')
            if(Button=='Yes'):
                sg.popup('已手動取消！')
                raise CloseSpider("使用者取消！")
        
        logging.info(self.info)
        print("\n未存在股號列表：")
        for printdata in range(len(self.noExist)):
            print(self.noExist[printdata])

    def manual_Mode(self, CO_ID):  # 手動模式
        if((len(CO_ID) == 4) and CO_ID.isnumeric()):
            self.ready_crawl = 1
            self.start_urls.append(
                f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={CO_ID}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=C')  # 帶入網址序列
        else:
            logging.error('請輸入正確的四位數純數字股號')
            pass

    def output_EmptyList_csv(self):  # 列出未存在股號
        now = datetime.now()
        dt_string = now.strftime("%Y-%m-%d %H-%M-%S")
        if(len(self.noExist)):
            self.noExist = list(filter(None, self.noExist))
            dict = {'代號': self.noExist}
            df = DataFrame(dict)
            filename = f'.\{dt_string} -財務報告- {self.Year} 年第 {self.Season} 季中未存在之股號.csv'
            df.to_csv(filename, index=False)
            print(f'已匯出未存在的股號至{filename}')
            sg.SystemTray.notify('財務報告爬蟲',f'爬取完成！\n已匯出未存在的股號至\n{filename}')
        else:
            logging.info('無缺漏股號。')

        if(len(self.cant_reach)):
            self.cant_reach = list(filter(None,self.cant_reach))
            dict = {'代號': self.cant_reach}
            df = DataFrame(dict)
            filename = f'.\{dt_string} -財務報告- {self.Year} 年第 {self.Season} 季中爬取失敗之股號.csv'
            df.to_csv(filename, index=False)
            print(f'已匯出未存在的股號至{filename}')
            sg.SystemTray.notify('財務報告爬蟲',f'爬取完成！\n已匯出爬取失敗的股號至\n{filename}')
        else:
            logging.info('無爬取失敗股號。')

    def spider_closed(self, spider):  # 爬蟲關閉時的動作
        self.output_EmptyList_csv()
        winsound.PlaySound("SystemQuestion",winsound.SND_ALIAS)
        sg.popup(f'已完成 {self.Year} - 第 {self.Season} 季度\n{self.info}')

    def __init__(self, Year='', Season='', CSV='', Mode='', CO_ID='', **kwargs):  # 初始化動作
        dispatcher.connect(self.spider_closed,
                           signals.spider_closed)  # 設置爬蟲關閉時的動作
        self.Year = Year  # 帶入參數年份 -a Year 數字字串
        self.Season = Season  # 帶入參數季度 -a Season 數字字串
        self.Mode = Mode  # 帶入爬蟲模式 -a Mode 文字字串，Auto與Manual模式
        if(Mode == 'Auto' or Mode == 'A'):
            sg.SystemTray.notify('財務報告爬蟲', '初始化 - 以批次模式進行中...', display_duration_in_ms=300, fade_in_duration=.2)
            logging.info(f'目前輸入的參數，年份：{Year}、季度{Season}、模式：{Mode}、CSV路徑:{CSV}')
            self.import_csv = CSV  # 匯入CSV之路徑 -a CSV 'Path'
            self.auto_Mode()
        elif (Mode == 'Manual' or Mode == 'M'):
            sg.SystemTray.notify('財務報告爬蟲', '初始化 - 以單筆模式進行中...', display_duration_in_ms=300, fade_in_duration=.2)
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

    def get_From_Table(self, items, response, tables_ID, tables_ItemsName,isTable1):  # 從表格當中獲取資料
        for datas in response.xpath('body/div[2]/div[3]'):
            for tableID in range(0, len(tables_ID)):  # 表一獲取資料
                if(isTable1):
                    tables1_ID_Type2 = ['11000', '1110', '1120','1136', '1139', '25XX', '31101']
                    tables1_ID_Type3 = ['111100', '112000', '1120','1136', '1139', '220000', '31100']
                    tables1_ID_Type4 = ['111100', '112000', '1120','1136', '1139', '220000', '301010']
                    data = datas.xpath(
                    f"//td[text() = '{tables_ID[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    data2 = datas.xpath(
                    f"//td[text() = '{tables1_ID_Type2[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    data3 = datas.xpath(
                    f"//td[text() = '{tables1_ID_Type3[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    data4 = datas.xpath(
                    f"//td[text() = '{tables1_ID_Type4[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    if(len(data)):
                        data[0] = data[0].replace(',', '')
                        if(self.is_Number(data[0])):
                            items[tables_ItemsName[tableID]] = float(data[0])
                            continue
                        else:
                            data[1] = data[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data[1]))
                            continue
                    elif(len(data2)):
                        data2[0] = data2[0].replace(',', '')
                        if(self.is_Number(data2[0])):
                            items[tables_ItemsName[tableID]] = float(data2[0])
                            continue
                        else:
                            data2[1] = data2[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data2[1]))
                            continue
                    elif(len(data3)):
                        data3[0] = data3[0].replace(',', '')
                        if(self.is_Number(data3[0])):
                            items[tables_ItemsName[tableID]] = float(data3[0])
                            continue
                        else:
                            data3[1] = data3[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data3[1]))
                            continue
                    elif(len(data4)):
                        data4[0] = data4[0].replace(',', '')
                        if(self.is_Number(data4[0])):
                            items[tables_ItemsName[tableID]] = float(data4[0])
                            continue
                        else:
                            data4[1] = data4[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data4[1]))
                            continue
                    else:
                        items[tables_ItemsName[tableID]] = float(0.0)
                else:
                    tables2_ID_Type2 = ['41000', '61000', '59000', '985000']
                    tables2_ID_Type3 = ['41000', '61000', '59000', '9750']
                    data = datas.xpath(
                    f"//td[text() = '{tables_ID[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    data2 = datas.xpath(
                    f"//td[text() = '{tables2_ID_Type2[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    data3 = datas.xpath(
                    f"//td[text() = '{tables2_ID_Type3[tableID]}']/following-sibling::td[2]//text()").getall() #td[2]使用代號，td[1]使用名稱定位
                    if(len(data)):
                        data[0] = data[0].replace(',', '')
                        if(self.is_Number(data[0])):
                            items[tables_ItemsName[tableID]] = float(data[0])
                            continue
                        else:
                            data[1] = data[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data[1]))
                            continue
                    elif(len(data2)):
                        data2[0] = data2[0].replace(',', '')
                        if(self.is_Number(data2[0])):
                            items[tables_ItemsName[tableID]] = float(data2[0])
                            continue
                        else:
                            data2[1] = data2[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data2[1]))
                            continue
                    elif(len(data3)):
                        data3[0] = data3[0].replace(',', '')
                        if(self.is_Number(data3[0])):
                            items[tables_ItemsName[tableID]] = float(data3[0])
                            continue
                        else:
                            data3[1] = data3[1].replace(',', '')
                            items[tables_ItemsName[tableID]] = -(float(data3[1]))
                            continue
                    
                    else:
                        items[tables_ItemsName[tableID]] = float(0.0)

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
                self.current += 1
                self.wait_url_A -= 1
                pass
            else:  # 試圖用回報A連結重新爬取
                _page_exist = False
                logging.info("該股類型C無資料，轉入類型A查資料。")
                self.wait_url_A += 1
                self.wait_url_A_MAX+=1
                self.start_urls.append(
                    f'https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID={company_Id[0]}&SYEAR={self.Year}&SSEASON={self.Season}&REPORT_ID=A')
                pass
        if(_page_exist):
            self.current += 1
            items['DATA_TYPE'] = self.Type
            items['SUB_DATA_TYPE'] = '個別財務報告' if (
                report_ID == 'A') else '合併財務報告'
            items['CO_ID'] = str(company_Id[0])
            co_name = str(response.xpath(
                '/html/body/div[2]/div[1]/div[2]/span[1]//text()').get())
            items['CO_FULL_NAME'] = co_name
            items['SYear'] = self.Year
            items['SSeason'] = self.Season
            # 主要爬蟲區
            tables1_ID = ['1100', '1110', '1120','1136', '1139', '25XX', '3110']
            #tables1_ID = ['現金及約當現金', '透過損益按公允價值衡量之金融資產－流動', '透過其他綜合損益按公允價值衡量之金融資產－流動',
            #              '按攤銷後成本衡量之金融資產－流動', '避險之金融資產－流動', '非流動負債合計', '普通股股本']
            tables1_ItemsName = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7']
            #tables2_ID = ['營業收入合計', '營業利益（損失）', '營業外收入及支出合計', '稀釋每股盈餘合計']
            tables2_ID = ['4000', '6900', '7000', '9850']
            tables2_ItemsName = ['B1', 'B2', 'B3', 'B4']
            self.get_From_Table(items, response, tables1_ID, tables1_ItemsName,True)
            self.get_From_Table(items, response, tables2_ID, tables2_ItemsName,False)
            if(co_name!=None):
                yield(items)
            else:
                self.cant_reach.append(company_Id)
                return
        self.print_info()
