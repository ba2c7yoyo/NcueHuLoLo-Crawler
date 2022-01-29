# -*- coding: utf-8 -*-
"""
Created on Sun Jun 28 18:12:22 2020

@author: ZHAOQI
"""
import csv
import requests
from bs4 import BeautifulSoup

url =  'http://webap0.ncue.edu.tw/deanv2/other/ob010' 
resp = requests.get(url)
resp.encoding = 'utf-8-sign'    #轉換編碼至UTF-8
soup=BeautifulSoup(resp.text,"lxml") #先抓全部
form_data={
"sel_cls_branch":"D",#日間部
"sel_yms_year":"109",#學年
"sel_yms_smester":"1",#學期
"sel_cls_id":"D240BN4A", #修課班別 
#生物系D240BN1A、D240BN2A、D240BN3A、D240BN4A
#精進英外文D410BV1H、大二體育DMD0BV00、大三四體育DMD0BV0A、教育學程DMA0BV00
#精進中文D420BV0E、軍訓課DM35BV00
"X-Requested-With":"XMLHttpRequest",}

filename='生物四'+'.csv' #輸出的檔案名
response_post=requests.post("http://webap0.ncue.edu.tw/deanv2/other/ob010"
                            ,data=form_data) #用post把選單資料送出
soup_post=BeautifulSoup(response_post.text, "lxml") #送出的資料放在soup_post

#%%
classnumember =["課程代碼："] #抓含有課程代碼:的字 用if
result_classnum=[]
for tnum in soup_post.find_all("td"):
    if tnum.get("data-th")in classnumember:
        print(tnum.text.strip())
        result_classnum+=[tnum.text.strip()]
#%% 
teachername =["教師姓名："] #抓抓含有教師姓名:的字 用if
result_teachername=[]
for tn in soup_post.find_all("td"):
    if tn.get("data-th")in teachername:
        print(tn.text.strip())
        result_teachername+=[tn.text.strip().replace('\n ','')] #合開老師合併
        
#%%
classname =["課程名稱："] #抓含有課程名稱:的字 用if
result_classname=[]
for td in soup_post.find_all("td"):
    if td.get("data-th")in classname:
        print(td.text.strip())
        result_classname+=[td.text.strip()]
#https://stackoverflow.com/questions/44336394/how-to-handle-td-data-without-tag-in-beautifulsoup
#https://www.slideshare.net/tw_dsconf/python-78691041 大講義
together=[]

#%%
#輸出csv檔案
with open(filename, 'w',newline='',encoding='utf-8-sig') as csvfile: ## 檔名格式，'a+'代表可覆寫        
    writer=csv.writer(csvfile,delimiter=',')
    writer.writerow(['代碼','課程名稱','老師名稱','關鍵字觸發詞'])
    for td in  range(len(result_classname)):
        together=( result_classnum[td],result_classname[td],result_teachername[td],result_teachername[td]+result_classname[td])
        writer.writerow(together)
csvfile.close()
print('csv檔案已儲存在資料夾。')
"""
#舊版:沒有老師會自動略過
result_teachername=[]
for tn in soup_post.find_all('span',style="white-space: pre-line"):
    print(tn.text.strip())
    result_teachername+=[tn.text.strip()]"""