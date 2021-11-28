# “股票数据一键获取”开发笔记

1. 完成 Scrapy 业务逻辑。

   - 新建 Scrapy 项目。

     ```shell
     scrapy startproject stock
     cd stock
     scrapy genspider eastmoney so.eastmoney.com
     ```

2. 完成 Selinenm 业务逻辑。

3. 完成 xlwings 业务逻辑。

   - 通过 Python 启动爬虫。

     ```python
     from scrapy import cmdline
     def start():
         cmdline.execute("scrapy crawl eastmoney".split())
     ```
   
   - 通过 xlwings 连接 Excel 和 Python 代码。
   
     ```vb
     Sub ClickToGetZdf()
     	RunPython "import crawl; crawl.start()"
     End Sub
     ```

