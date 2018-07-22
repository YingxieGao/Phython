#encoding=utf-8
"""
@author: Yingxie Gao
"""

import requests
import json
import xlwt
import tkinter as tk

style=xlwt.XFStyle
font=xlwt.Font()
font.name='SimSun'
style.font=font

#create a table
w=xlwt.Workbook(encoding='utf-8')
ws=w.add_sheet('sheet 1', cell_overwrite_ok=True)

row=1

ws.write(0,0,"Comment")
ws.write(0,1,"time")
ws.write(0,2,"product")
ws.write(0,3,"User")
def write_json_to_xls(dat):
    global row
    for comment in dat["comments"]:
        ws.write(row,0,comment["content"])
        ws.write(row,1,comment["date"])
        ws.write(row,2,comment['auction']["sku"])
        ws.write(row,3,comment['user']["nick"])
        row+=1

id=''
        
def  go():
    print(xls_text.get())
    root.quit()
    root.destroy()
    
root=tk.Tk()
root.title("input id")
root.geometry('300x100')
xls_text=tk.StringVar()
info=tk.Label(root, text='input')
info.pack()
xls = tk.Entry(root, textvariable=xls_text)
xls_text.set(" ")
xls.pack()
tk.Button(root,text='ok', command=go).pack()
root.mainloop()
id=xls_text.get().lstrip()

url='https://rate.taobao.com/feedRateList.htm?auctionNumId='+str(id)+'&currentPageNum=1'
print(url)
json_req = requests.get(url)
dat=json.loads(json_req.text.strip().strip('()'))
max = dat['total']
print("This product has",max,"comments")
page = 0
while row<max:
    try:
        page = page + 1
        print(url)
        json_req = requests.get(url[:-1]+str(page))       
        dat = json.loads(json_req.text.strip().strip('()'))
        write_json_to_xls(dat)
        print('We have got',row-1,'comments!')
    except Exception as e:
        print('failed！',e)
        break
    except TypeError as e:
        print('failed,', e, ".Sorry, Taobao only open 251 pages of comments")
        break
            
print("Done, check your comments.xls")                        
w.save('comments.xls')
input ("Please Enter to exit:")