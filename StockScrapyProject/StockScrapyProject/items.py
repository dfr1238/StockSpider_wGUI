# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class StockSpider_items(scrapy.Item):
    ##基本資料
    DATA_TYPE = scrapy.Field() #資料類型－財務報告
    SUB_DATA_TYPE= scrapy.Field() #子資料類別－報告類型
    CO_ID = scrapy.Field() #股號
    CO_FULL_NAME = scrapy.Field() #公司全名
    SYear = scrapy.Field() #年度
    SSeason = scrapy.Field() #季度
    ##A
    A1 = scrapy.Field() #A1:現金及約當現金
    A2 = scrapy.Field() #A2:透過損益按公允價值衡量之金融資產－流動
    A3 = scrapy.Field() #A3:透過其他綜合損益按公允價值衡量之金融資產－流動
    A4 = scrapy.Field() #A4:按攤銷後成本衡量之金融資產－流動
    A5 = scrapy.Field() #A5:避險之金融資產－流動
    A5_5 = scrapy.Field() #A5_5:合約資產 - 流動
    A6 = scrapy.Field() #A6:非流動負債合計
    A7 = scrapy.Field() #A7:普通股股本
    ##B
    B1 = scrapy.Field() #B1:營業收入合計
    B2 = scrapy.Field() #B2:營業利益（損失）
    B3 = scrapy.Field() #B3:營業外收入及支出合計
    B4 = scrapy.Field() #B4:稀釋每股盈餘合計
    pass

class StockPrice_items(scrapy.Item):
    #基本資料
    DATA_TYPE = scrapy.Field() #資料類型－股價
    SUB_DATA_TYPE= scrapy.Field() #子資料類別－交易所
    CO_ID = scrapy.Field()  #股號
    CO_SHORT_NAME = scrapy.Field()    #公司縮名
    SYear = scrapy.Field()  #年度
    SDate = scrapy.Field()  #詳細日期
    Price = scrapy.Field()  #收盤價
    pass
