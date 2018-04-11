# -*- coding: utf-8 -*-
"""
Created on Wed Apr 11 14:31:05 2018

@author: sida
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Apr  8 16:55:09 2018

@author: sida
"""

import scrapy
import pandas as pd
import datetime



class SinaLive_Spider(scrapy.Spider):
    name = 'SinaLive'

    start_urls = ['http://live.sina.com.cn/zt/f/v/finance/globalnews1']
    
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
        l_title_ix = 0
        ws = pd.read_excel('./NatReform/Crawled_Data/SinaLive.xlsx') 
        title = response.css('p.bd_i_txt_c::text')[l_title_ix].extract()
        link = 'http://live.sina.com.cn/zt/f/v/finance/globalnews1'
 
        concept = '新浪财经24小时直播'
        # max_len = len(response.css('p.bd_i_txt_c::text'))
        latest_title = ws.iloc[-1].title
        latest_ID = ws.iloc[-1].SpiderSpecID + 1
        
        if(title != latest_title):
            
            new_rec = pd.DataFrame([[title,link,concept,datetime.datetime.now(), latest_ID]], columns= \
                                     ['title','link','concept','time','SpiderSpecID'])
            
            ws = ws.append(new_rec)
            ws.to_excel('./NatReform/Crawled_Data/SinaLive.xlsx')
            
        
# =============================================================================
# A module for tracing back missed logs
#         while(title != latest_title):
#             
#             new_rec = pd.DataFrame([[title,link,concept,datetime.datetime.now(), latest_ID]], columns= \
#                                      ['title','link','concept','time','SpiderSpecID'])
#             
#             ws = ws.append(new_rec)
#             ws.to_excel('./NatReform/Crawled_Data/SinaLive.xlsx')
#             l_title_ix = l_title_ix + 1
#             latest_ID = latest_ID + 1
#             if l_title_ix < max_len:
#                 title = response.css('p.bd_i_txt_c::text')[l_title_ix].extract()
#             else: 
#                 title = latest_title
# =============================================================================
                
        yield {
                    'title': title,
                    'link': link,
                    'concept': concept}
            
        
            
        