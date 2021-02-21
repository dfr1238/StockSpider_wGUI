from StockScrapyProject.StockScrapyProject.spiders.stockSpider import StockSpider
from StockScrapyProject.StockScrapyProject.spiders.stockPriceSpider import stockPriceSpider
from scrapy.crawler import CrawlerProcess, CrawlerRunner
from scrapy.utils.project import get_project_settings
from twisted.internet import reactor
import os


class Scraper:
    def __init__(self):
        settings_file_path = 'StockScrapyProject.StockScrapyProject.settings' # The path seen from root, ie. from main.py
        os.environ.setdefault('SCRAPY_SETTINGS_MODULE', settings_file_path)
        self.process = CrawlerProcess(get_project_settings())
        self.spider = StockSpider # The spider you want to crawl
        self.runner  = CrawlerRunner(get_project_settings())

    def set_PriceSpider(self,CSV=''):
        self.process.crawl(stockPriceSpider,CSV_File_PATH=CSV)
    
    def run_PriceSpider(self):
        self.process.start()

    def set_StockSpider(self,Year='',Season='',CSV='',Mode='',CO_ID='', **kwargs):
        self.process.crawl(StockSpider,Year=Year,Season=Season,CSV=CSV,Mode=Mode,CO_ID=CO_ID)
    
    def run_StockSpider(self):
        self.process.start()