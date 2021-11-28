from scrapy import cmdline
def run():
    # TODO: 通过 Excel 启动 Scrapy
    cmdline.execute('scrapy crawl eastmoney'.split())