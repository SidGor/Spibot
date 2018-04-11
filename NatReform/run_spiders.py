# -*- coding: utf-8 -*-
"""
Created on Wed Apr 11 14:49:12 2018

@author: sida
"""

from subprocess import call
import pandas as pd
import time

while True:
    
    spider_list = pd.read_excel("..\spider_records.xlsx")['SpiderName']

    for spiders in spider_list:
        call("scrapy crawl " + spiders)
        
    time.sleep(5)