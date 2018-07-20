#encoding=utf-8

import requests
import re
import json
import time
import xlwt

style=xlwt.XFStyle
font=xlwt.Font()
font.name='SimSun'
style.font=font

#create a table
w=xlwt.Workbook(encoding='utf-8')

ws=w.add_sheet('sheet 1', cell_overwrite_ok=True)

row=1

ws.write(0,0,"评论")
ws.write(0,1,"时间")
ws.write(0,2,"购买的产品")
ws.write(0,3,"用户")
def write_json_to_xls(dat):
    global row
    for comment in dat["comments"]:
        ws.write(row,0,comment["content"])
        ws.write(row,1,comment["date"])
        ws.write(row,2,comment['auction']["sku"])
        ws.write(row,3,comment['user']["nick"])
        row+=1
        
        
for i in range(1,100):       
    url='https://rate.taobao.com/feedRateList.htm?auctionNumId=573002059428&currentPageNum='+str(i)
    try:
        print(url)
        json_req = requests.get(url)
        dat=json.loads(json_req.text.strip().strip('()'))
        write_json_to_xls(dat)
        print('WOW!')
    except Exception as e:
        print('failed！',e)
            
            
w.save('taobao.xls')