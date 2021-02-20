from StockScrapyProject.StockScrapyProject.spiders.stockSpider import StockSpider
from StockScrapyProject.StockScrapyProject.spiders.stockPriceSpider import stockPriceSpider
from scrapy.crawler import CrawlerProcess, CrawlerRunner
from scrapy.utils.project import get_project_settings
from scrapy.utils.log import configure_logging
from twisted.internet import reactor
from twisted.internet import defer
import os


class Scraper:
    def __init__(self):
        settings_file_path = 'StockScrapyProject.StockScrapyProject.settings' # The path seen from root, ie. from main.py
        os.environ.setdefault('SCRAPY_SETTINGS_MODULE', settings_file_path)
        self.process = CrawlerProcess(get_project_settings())
        self.spider = StockSpider # The spider you want to crawl
        self.runner  = CrawlerRunner(get_project_settings())
    dfs = set()
    def set_StockSpider(self,Year='',Season='',CSV='',Mode='',CO_ID='', **kwargs):
        runner  = CrawlerRunner(get_project_settings())
        d = runner.crawl(self.spider,Year=Year,Season=Season,CSV=CSV,Mode=Mode,CO_ID=CO_ID)
        self.dfs.add(d)
        defer.DeferredList(self.dfs).addBoth(lambda _: reactor.stop())
        #d.addBoth(lambda _: reactor.stop())
    
    def run_StockSpider(self):
        reactor.run() # the script will block here until the crawling is finished