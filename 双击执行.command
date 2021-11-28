#!/bin/bash
cd `dirname $0`/stock/stock
scrapy crawl eastmoney
# read