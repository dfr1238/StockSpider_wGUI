from StockScrapyProject.StockScrapyProject.spiders.stockSpider import StockSpider
from StockScrapyProject.StockScrapyProject.spiders.stockPriceSpider import stockPriceSpider
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings
import os


class Scraper:
    def __init__(self):
        settings_file_path = 'StockScrapyProject.StockScrapyProject.settings' # The path seen from root, ie. from main.py
        os.environ.setdefault('SCRAPY_SETTINGS_MODULE', settings_file_path)
        self.process = CrawlerProcess(get_project_settings())
        self.spider = StockSpider # The spider you want to crawl

    def run_StockSpider(self,Year='',Season='',CSV='',Mode='',CO_ID='', **kwargs):
        self.process.crawl(self.spider,Year=Year,Season=Season,CSV=CSV,Mode=Mode,CO_ID=CO_ID)
        self.process.start()  # the script will block here until the crawling is finishe