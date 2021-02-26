from StockScrapyProject.StockScrapyProject.spiders.stockSpider import StockSpider
from StockScrapyProject.StockScrapyProject.spiders.stockPriceSpider import stockPriceSpider
from scrapy.crawler import CrawlerProcess, CrawlerRunner
from scrapy.utils.project import get_project_settings
import os
from twisted.internet import reactor

class Scraper:
    def __init__(self):
        settings_file_path = 'StockScrapyProject.StockScrapyProject.settings' # The path seen from root, ie. from main.py
        os.environ.setdefault('SCRAPY_SETTINGS_MODULE', settings_file_path)
        setting = get_project_settings()
        self.process = CrawlerProcess(setting)
        self.runner  = CrawlerRunner(setting)

    def change_Project_Setting(self,MONOGBURL,DBNAME,CONAME):
        setting = get_project_settings()
        setting.set('MONGO_URI',MONOGBURL)
        setting.set('MONGO_DATABASE',DBNAME)
        setting.set('MONGO_CODATA',CONAME)
        self.process = CrawlerProcess(setting)
        self.runner  = CrawlerRunner(setting)

    def set_PriceSpider(self,CSV=''):
        self.process.crawl(stockPriceSpider,CSV_File_PATH=CSV)
        #d = self.runner.crawl(stockPriceSpider,CSV_File_PATH=CSV)
        #d.addBoth(lambda _: reactor.stop())
    
    def run_PriceSpider(self):
        self.process.start()
        #reactor.run()

    def set_StockSpider(self,Year='',Season='',CSV='',Mode='',CO_ID='', **kwargs):
        d = self.runner.crawl(StockSpider,Year=Year,Season=Season,CSV=CSV,Mode=Mode,CO_ID=CO_ID)
        d.addBoth(lambda _: reactor.stop())
    
    def run_StockSpider(self):
        reactor.run()