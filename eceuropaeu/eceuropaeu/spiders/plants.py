import scrapy


class PlantsSpider(scrapy.Spider):
    name = 'plants'
    allowed_domains = ['ec-europa.eu']
    start_urls = ['http://ec-europa.eu/']

    def parse(self, response):
        pass
