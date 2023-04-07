import scrapy
import pandas as pd

class SuperstoreSpider(scrapy.Spider):
    name = "superstore"
    allowed_domains = ["walmart.ca"]
    start_urls = ["http://walmart.ca/"]

    def __init__(self, name=None, **kwargs):
        super().__init__(name, **kwargs)
        self.start_urls = [self.url]
        self.downloaded_items = 0

    def parse(self, response):
        pass
