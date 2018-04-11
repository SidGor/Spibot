# =============================================================================
# from openpyxl import load_workbook
# =============================================================================


# Find current available spiders and their last updates

# =============================================================================
# spider_record = load_workbook('./spider_records.xlsx')
# spider_record_ws = spider_record[spider_record.sheetnames[0]]
# 
# =============================================================================
#making test table 
# =============================================================================
# spider_record_ws.append({1:'test', 2:'test', 3:'test'})
#  
# =============================================================================

import pandas as pd
import queue
import itchat
import time
import random
from decimal import Decimal 
import datetime

itchat.login()

MessageQueue = queue.Queue()

while True:

    spider_record = pd.read_excel('./spider_records.xlsx')  
    # =============================================================================
    # Testing the append fuction in pandas.
    #
    # test_record = {'SpiderName':['a'],'LastUpdateTitle':['a'],'LastUpdateLink':['a']}
    # test_record1 = pd.DataFrame([['b','b','b']], columns = ['SpiderName','LastUpdateTitle','LastUpdateLink'])
    # spider_record.append(test_record1)
    # =============================================================================
    spider_list = spider_record.iloc[:,0]
    
    
    # The reader is in charge of reaiding news feeds into Message Queue:
    
    for spiders in spider_list:
        print('reading: '+spiders)
        spider_doc = pd.read_excel('./NatReform/NatReform/Crawled_Data/'+ spiders + '.xlsx')
        if spider_doc.shape[0] == 0:
            print(spiders + " has no current record to send.")
        else:
            spider_reg_rec = spider_record.loc[spider_record['SpiderName'] == spiders].iloc[-1]
            last_ID = spider_reg_rec.SpiderSpecID
            if last_ID != spider_doc.iloc[-1].SpiderSpecID:
                
                for ID in spider_doc.iloc[last_ID + 1:].SpiderSpecID:
                    spider_doc_temp = spider_doc.loc[spider_doc['SpiderSpecID'] == ID].iloc[-1]   
                    quote = spiders + ' (' + spider_doc_temp.concept + ')《' + \
                            spider_doc_temp.title + '》 链接: ' + spider_doc_temp.link + \
                            '  于' + spider_doc_temp.time.__str__() + ' 捕获.'
                            
                    new_rec = pd.DataFrame([[spiders, spider_doc_temp.title, \
                                             spider_doc_temp.link,spider_doc_temp.SpiderSpecID]], columns \
                                            = ['SpiderName','LastUpdateTitle','LastUpdateLink','SpiderSpecID'])
                    spider_record.iloc[spider_reg_rec.name] = new_rec.iloc[-1]
                
                    spider_record.to_excel('./spider_records.xlsx')  #overwrite old data
                    MessageQueue.put(quote) # Insert Record to a queue in global enviornment
                    print(spiders+' updated to latest record')
    # Test Environment for the MessageQueue
                
    # =============================================================================
    # MessageQueue.put("Spibot Testing: 你应该这么做，")
    # MessageQueue.put("Spibot Testing: 我也应该死。 ")
    # MessageQueue.put("Spibot Testing: 曾经有一份真诚的爱情放在我面前，")
    # MessageQueue.put("Spibot Testing: 我没有珍惜，")
    # MessageQueue.put("Spibot Testing: 等到失去的时候才后悔莫及，")
    # MessageQueue.put("Spibot Testing: 人世间最痛苦的事莫过于此。")
    # MessageQueue.put("Spibot Testing: 你的剑在我的咽喉上割下去吧！")
    # MessageQueue.put("Spibot Testing: 不用再犹豫了！")
    # MessageQueue.put("Spibot Testing: 如果上天能够给我一个再来一次的机会，")
    # MessageQueue.put("Spibot Testing: 我会对那个女孩子说三个字：我爱你。")
    # MessageQueue.put("Spibot Testing: 如果非要在这份爱上加上一个期限，")
    # MessageQueue.put("Spibot Testing: 我希望是……")
    # MessageQueue.put("Spibot Testing: 一万年…… ")
    # =============================================================================
    
    
    
    # A sender in charge of messages
    while (not MessageQueue.empty()):
        receiver = itchat.search_chatrooms(name = 'Spibot1群')[0].UserName
        itchat.send(msg = MessageQueue.get(), toUserName = receiver)
        time.sleep(round(random.uniform(3,5),2))
    
    
    if ((datetime.datetime.now().hour > 21) & (datetime.datetime.now().minute > 0)):
        itchat.send(msg = "Spibot Message:房子买了吗，老婆娶了吗，孩子有了吗？没有就赶紧回家相亲吧，Spibot在这美好的夜晚向您祝福。", toUserName = receiver)
        break
    
    time.sleep(5)
    
    