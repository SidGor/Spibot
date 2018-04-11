# -*- coding: utf-8 -*-
"""
Created on Sun Apr  8 16:55:09 2018

@author: sida
"""

import scrapy
import pandas as pd
import datetime



class BJgzw_UDSpider(scrapy.Spider):
    name = 'BJGZW'

    start_urls = ['http://www.bjgzw.gov.cn/SyAction.do?method=js&js_bt=%BB%EC%B8%C4&js_lm=&js_zw=']
    
    headers =  {'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                        'Accept-Encoding':'gzip, deflate, sdch',
                        'Accept-Language':'en-US,en;q=0.8',
                        'Cache-Control':'max-age=0',
                        'Connection':'keep-alive',
                        'Cookie':'_gscbrs_1429685822=1; _gscu_1429685822=21680099q711yw13; JSESSIONID=A903D93BA7D4DE75C8470380BCDCE588',
                        'Host':'www.bjgzw.gov.cn',
                        'Referer':'http://www.bjgzw.gov.cn/QtCommonAction.do?method=cxxx&type=0000004050&id=fb220baaa23200&fanhuiFlag=1',
                            'Upgrade-Insecure-Requests':'1',
                        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36'
                      }
    
    def parse(self, response):
        ws = pd.read_excel('./NatReform/Crawled_Data/BJGZW.xlsx') 
        title = response.css('tbody div a::text').extract()[0].strip()
        link = 'www.bjgzw.gov.cn' + response.css('tbody div a::attr(href)').extract()[0]
        concept = 'Testing.Co[600000.SH]'
        latest_link = ws.iloc[-1].link
        latest_ID = ws.iloc[-1].SpiderSpecID + 1
        
        if(link != latest_link):
            new_rec = pd.DataFrame([[title,link,concept,datetime.datetime.now(), latest_ID]], columns= \
                                     ['title','link','concept','time','SpiderSpecID'])
            
            ws = ws.append(new_rec)
            ws.to_excel('./NatReform/Crawled_Data/BJGZW.xlsx')
            
        yield {
                    'title': title.replace(","," "),
                    'link': 'www.bjgzw.gov.cn' + response.css('tbody div a::attr(href)').extract()[0],
                    'concept': 'Testing.Co[600000.SH]',}
            
        
            
        