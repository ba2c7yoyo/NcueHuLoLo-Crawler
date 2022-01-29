# -*- coding: utf-8 -*-
"""
Created on Sun Jun 28 18:12:22 2020

@author: ZHAOQI
功能:將教務系統的課爬下來-->輸出表格到google sheet
"""
import requests
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
#%%
#選定要上傳的GOOGLE SHEET位置
auth_json_path = 'ringed-spirit-281809-b1f94a7661f8.json'
gss_scopes = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(auth_json_path,gss_scopes)
gss_client = gspread.authorize(credentials)
spreadsheet_key = '1bjAxvSGNflQKFqF1KPv-PlQiqBs0qaiLt-eerPoVY6Q' #表格網址(選)
#要記得表格要先讓python-sheet@ringed-spirit-281809.iam.gserviceaccount.com共用
sheetrow = gss_client.open_by_key(spreadsheet_key)
localtime = time.asctime( time.localtime(time.time()) )
sheet = sheetrow.worksheet("生四") #分頁名稱(選) #精進中英外文 大二三四體育課
sheet.clear() #精進中英外文不刪除資料
listtitle=['代碼','課程名稱','老師名稱','關鍵字觸發詞','更新日期:'+localtime] #精進中英外文 大二三四體育課不刪除資料
sheet.append_row(listtitle)  # 標題         #精進中英外文 大二三四體育課不刪除資料
#%%
#抓取最初教務系統課表
url =  'http://webap0.ncue.edu.tw/deanv2/other/ob010' 
resp = requests.get(url)
resp.encoding = 'utf-8-sign'    #轉換編碼至UTF-8
soup=BeautifulSoup(resp.text,"lxml") #先抓全部
form_data={
"sel_cls_branch":"D",#日間部
"sel_yms_year":"109",#學年(選)
"sel_yms_smester":"1",#學期(選)
"sel_cls_id":"D240BN4A", #修課班別 (選)
#生一D240BN1A、生二D240BN2A、生三D240BN3A、生四D240BN4A
#精進英外文D410BV1H、大二體育DMD0BV00、大三四體育DMD0BV0A、教育學程DMA0BV00
#精進中文D420BV0E、軍訓課DM35BV00
#核心通識D720BV0A
"X-Requested-With":"XMLHttpRequest",}

#filename='軍訓課'+'.csv' #輸出的檔案名(選)
response_post=requests.post("http://webap0.ncue.edu.tw/deanv2/other/ob010"
                            ,data=form_data) #用post把選單資料送出
soup_post=BeautifulSoup(response_post.text, "lxml") #送出的資料放在soup_post

#%%
#抓含有課程代碼:的字 用if
classnumember =["課程代碼："] 
result_classnum=[]
for tnum in soup_post.find_all("td"):
    if tnum.get("data-th")in classnumember:
        print(tnum.text.strip())
        result_classnum+=[tnum.text.strip()]
#%% 
#抓抓含有教師姓名:的字 用if
teachername =["教師姓名："] 
result_teachername=[]
for tn in soup_post.find_all("td"):
    if tn.get("data-th")in teachername:
        print(tn.text.strip())
        result_teachername+=[tn.text.strip().replace('\n ','')] #合開老師合併
        
#%%
#抓含有課程名稱:的字 用if
classname =["課程名稱："] 
result_classname=[]
for td in soup_post.find_all("td"):
    if td.get("data-th")in classname:
        print(td.text.strip())
        result_classname+=[td.text.strip()]
#https://stackoverflow.com/questions/44336394/how-to-handle-td-data-without-tag-in-beautifulsoup
#https://www.slideshare.net/tw_dsconf/python-78691041 大講義
#%%
#GOOGLE SHEET編輯區
for td in range(len(result_classname)):
   listdata=[result_classnum[td],result_classname[td],result_teachername[td],result_teachername[td]+result_classname[td]]
   sheet.append_row(listdata)  # 資料內容


print('資料上傳google sheet成功。')
"""
#舊版:沒有老師會自動略過
result_teachername=[]
for tn in soup_post.find_all('span',style="white-space: pre-line"):
    print(tn.text.strip())
    result_teachername+=[tn.text.strip()]"""