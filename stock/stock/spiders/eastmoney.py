import scrapy
from scrapy_selenium import SeleniumRequest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import xlwings as xw
import os

class EastmoneySpider(scrapy.Spider):
    name = 'eastmoney'
    allowed_domains = ['so.eastmoney.com', 'quote.eastmoney.com']
    # 获取本地数据文件
    wb = xw.Book(os.path.abspath('../../股票数据一键获取.xlsm'))
    sheet = wb.sheets['Sheet1']
    # 绑定指定数据列名
    stockNameCol = 'A'
    stockCodeCol = 'B'
    stockBoardCol = 'C'
    stockOpenPriceCol = 'D'
    stockNowPriceCol = 'E'
    stockZdfCol = 'F'

    def start_requests(self):
        # 股票数据一键获取
        stockNames = []
        for i in range(2, 102):
            stockZdf = self.sheet.range(self.stockZdfCol + str(i)).value
            if stockZdf:
                # 断点续爬
                continue
            stockName = self.sheet.range(self.stockNameCol + str(i)).value
            if stockName is not None:
                if stockName not in stockNames:
                    stockNames.append(stockName)
                    url = 'https://so.eastmoney.com/web/s?keyword=' + stockName
                    yield SeleniumRequest(
                        url=url,
                        meta={'row': str(i)},
                        callback=self.parse_info,
                        wait_time=5,
                        wait_until=EC.presence_of_element_located((By.CLASS_NAME, 'exstock'))
                    )
                else:
                    # 同名股票重复添加
                    self.sheet.range(self.stockZdfCol + str(i)).value = '已存在'
            else:
                break

    def parse_info(self, response):
        url = response.css('.exstock_exinfo_links a::attr(href)').get() # 行情页
        yield SeleniumRequest(
            url=url,
            meta={'row': response.meta['row']},
            callback=self.parse_result,
            wait_time=3,
            wait_until=EC.text_to_be_present_in_element((By.ID, 'km2'), '%')
        )

    # 股票数据抽取
    def parse_result(self, response):
        # 股票代码
        stockCode = response.css('#code::text').get()
        self.sheet.range(self.stockCodeCol + response.meta['row']).value = stockCode
        # 上市板块
        stockBoard = response.css('#jys-box>a>b::text').get()[0:1]
        self.sheet.range(self.stockBoardCol + response.meta['row']).value = stockBoard
        # 开盘价
        stockOpenPrice = response.css('#gt1::text').get()
        if stockOpenPrice is None:
            stockOpenPrice = '-'
        self.sheet.range(self.stockOpenPriceCol + response.meta['row']).value = stockOpenPrice
        # 最新价
        stockNowPrice = response.css('#price9::text').get()
        if stockNowPrice is None:
            stockNowPrice = '-'
        self.sheet.range(self.stockNowPriceCol + response.meta['row']).value = stockNowPrice
        # 涨跌幅
        stockZdf = response.css('#km2::text').get()
        if stockZdf is None:
            stockZdf = '-'
        self.sheet.range(self.stockZdfCol + response.meta['row']).value = stockZdf
    