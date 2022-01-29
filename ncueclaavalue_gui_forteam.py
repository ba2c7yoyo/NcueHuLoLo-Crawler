# -*- coding: utf-8 -*-
"""
Created on 05 21 2021

@author: ZHAOQI NCUE (For NCUEHULOLO TEAM)

"""

import tkinter as tk
import tkinter.ttk as ttk
import csv
import requests
from bs4 import BeautifulSoup
global access_code
import os
import glob
import xlwt 
access_code=0
def generate_data_step1 (access_code):
    print("正在通過generate_data_step1...")
    #驗證使用者是否未輸入以及將輸入的值轉換為爬蟲所需的參數
    global year_val,smester_val
    year_val=sel_yms_year.get() #開課學年
    if year_val=="": #驗證
        tk.messagebox.showwarning(title='提醒', message='未選擇開課學年')
        return 0
    
    smester_val=sel_yms_smester.get() #開課學期
    if smester_val=="":#驗證
        tk.messagebox.showwarning(title='提醒', message='未選擇開課學期')
        return 0 
    if access_code==102:
       classid_val="DMD0BV0A"
       print("正在進入第二次抓取課程變數中(step1_102)...")
    elif access_code==104:
       classid_val="D410BV1H"    
       print("正在進入第二次抓取課程變數中(step1_104)...")
    else:
        sel_cls_id_val=sel_cls_id.get() #課程類型
        if sel_cls_id_val=="核心通識":
            classid_val="D720BV0A"
        elif sel_cls_id_val=="教育學程":
            classid_val="DMA0BV00"
        elif sel_cls_id_val=="精進中英外文":#因為教務系統分開查詢關係
            classid_val="D420BV0E"
            if access_code==100:
                access_code=103
                print("第一次通過精進中英外文....(generate_data_step1)")
            elif access_code==200:             
                access_code=203
        elif sel_cls_id_val=="大二三四體育" :#因為教務系統分開查詢關係
            classid_val="DMD0BV00"
            if access_code==100:
                access_code=101
                print("第一次通過大二三四體育....(generate_data_step1)")
            elif access_code==200:             
                access_code=201
        elif sel_cls_id_val=="軍訓課":
            classid_val="DM35BV00"
        elif sel_cls_id_val=="生一":
            classid_val="D240BN1A"
        elif sel_cls_id_val=="生二":
            classid_val="D240BN2A"
        elif sel_cls_id_val=="生三":
            classid_val="D240BN3A"
        elif sel_cls_id_val=="生四":
            classid_val="D240BN4A"
        elif sel_cls_id_val=="":#驗證
            tk.messagebox.showwarning(title='提醒', message='未選擇課程類型')
            return 0 
    biology_val=biology.get()#生物系判別式，生物系的課程代碼都是24xxx開頭
    if biology_val==1:
       biologycode="24___" 
    else:
       biologycode="" 
    web_craeler_step2(access_code,year_val,smester_val,classid_val,biologycode)
    #最後將使用者選擇的參數丟給爬蟲處理(web_craeler_step2)
def web_craeler_step2(access_code,year_val,smester_val,classid_val,biologycode): #爬蟲
    url =  'https://webap0.ncue.edu.tw/deanv2/other/ob010' 
    resp = requests.get(url)
    resp.encoding = 'utf-8-sign'    #轉換編碼至UTF-8
    BeautifulSoup(resp.text,"lxml") #先抓全部
    form_data={
        "sel_cls_branch":"D",#日間部
        "sel_yms_year":year_val,#學年
        "sel_yms_smester":smester_val,#學期
        "sel_cls_id":classid_val, #修課班別 
        #生物系D240BN1A、D240BN2A、D240BN3A、D240BN4A
        #精進英外文D410BV1H、大二體育DMD0BV00、大三四體育DMD0BV0A、教育學程DMA0BV00
        #精進中文D420BV0E、軍訓課DM35BV00 課程代碼_生物系查詢用
        "scr_selcode":biologycode,
        "X-Requested-With":"XMLHttpRequest"}
    response_post=requests.post("https://webap0.ncue.edu.tw/deanv2/other/ob010"
                            ,data=form_data) #用post把選單資料送出
    soup_post=BeautifulSoup(response_post.text, "lxml") #送出的資料放在soup_post
    classnumember =["課程代碼："] #抓含有課程代碼:的字 用if
    result_classnum=[]
    for tnum in soup_post.find_all("td"):
        if tnum.get("data-th")in classnumember:
            print(tnum.text.strip()) #檢查點
            result_classnum+=[tnum.text.strip()]
    teachername =["教師姓名："] #抓抓含有教師姓名:的字 用if
    result_teachername=[]
    for tn in soup_post.find_all("td"):
        if tn.get("data-th")in teachername:
            #print((tn.text.strip().replace('/n',''))) #檢查點
            mod_moreteacher=tn.text.strip()
            if len( mod_moreteacher)>8:
                 mod_moreteacher="【合開】" #處理超過兩位老師開課的問題20201224       
            result_teachername+=[ mod_moreteacher.replace('\n',"").replace('\r',"")] #合開老師合併

            #if len(str(result_teachername))>8:
             #   result_teachername='合開'
        
    classname =["課程名稱："] #抓含有課程名稱:的字 用if
    result_classname=[]
    for td in soup_post.find_all("td"):
        if td.get("data-th")in classname:
           #print(type(td.text.strip())) #檢查點
           #result_classname+=[re.sub("[A-Za-z\&\,\'\s+]", "",td.text.strip())]
           a=td.text.strip()
           b=a.split(" \r")
           c=[]
           c=b[0].split(",")
           result_classname+=c  
           #把課名的一堆空格(s+、英文, &都砍掉20210102
           #print(c) #檢查點 
      
           
#https://stackoverflow.com/questions/44336394/how-to-handle-td-data-without-tag-in-beautifulsoup
#https://www.slideshare.net/tw_dsconf/python-78691041 大講義
    if access_code==200: #選擇輸出csv
        export_csv(result_classnum,result_classname,result_teachername)
    elif access_code==201: #選擇輸出csv且大二三四體育第一次爬蟲
          classid_val="DMD0BV0A"
          export_csv_language_PEclass_1(access_code,year_val,smester_val,classid_val,result_classnum,result_classname,result_teachername)
    elif access_code==202: #選擇輸出csv且大二三四體育第二次爬蟲
          export_csv_language_PEclass_2(result_classnum,result_classname,result_teachername)   
    elif access_code==203: #選擇輸出csv且精進中英外文第一次爬蟲
          classid_val="D410BV1H"#要變換成英外文代碼
          export_csv_language_PEclass_1(access_code,year_val,smester_val,classid_val,result_classnum,result_classname,result_teachername)
    elif access_code==204: #選擇輸出csv且精進中英外文第二次爬蟲
          export_csv_language_PEclass_2(result_classnum,result_classname,result_teachername)   
    else:
        tk.messagebox.showwarning(title='提醒', message='錯誤')
        return 0
def choice_csv_filenamechecking():#檔名驗證
    global file_ENTRY_val
    file_ENTRY_val=file_ENTRY.get()
    if file_ENTRY_val=="":
       tk.messagebox.showwarning(title='提醒', message='未輸入檔名')
       return 0 
    access_code=200 #若驗證通過(200)則交由爬蟲
    generate_data_step1(access_code)      
def choice_csv():
    global file_ENTRY
    tk.Label(labelframe,text="檔名(csv/xls)").grid(row=10,column=0)      
    file_ENTRY=tk.StringVar()
    tk.Entry(labelframe,textvariable=file_ENTRY).grid(row=10,column=1)
    tk.Button(labelframe,text="執行excel檔",command=choice_csv_filenamechecking).grid(row=11,column=1)
    #傳送到choice_csv_filenamechecking
def export_csv(result_classnum,result_classname,result_teachername):#輸出csv檔案
    filename_all=file_ENTRY_val+'.csv' #輸出的檔案名
    together=[]
    with open(filename_all, 'w',newline='',encoding='utf-8-sig') as csvfile: ## 檔名格式，'a+'代表可覆寫        
        writer=csv.writer(csvfile,delimiter=',')
        writer.writerow(['classnumember','classname','teachername','keywordup']) 
        for td in  range(len(result_classname)):
            together=(result_classnum[td],result_classname[td],result_teachername[td],result_teachername[td]+result_classname[td])
            writer.writerow(together)   
    csvfile.close()
    Csvtoxls(filename_all) #轉換csv to xls 20201225
    tk.messagebox.showinfo(title='提醒', message='csv/xls檔案已儲存完成!')
def export_csv_language_PEclass_1(access_code,year_val,smester_val,classid_val,result_classnum,result_classname,result_teachername):
    if access_code==201:#換到第二次爬蟲時候的課程代碼
        access_code=202
    elif access_code==203:
        access_code=204
    filename_all=file_ENTRY_val+'.csv' #輸出的檔案名
    together=[]
    with open(filename_all, 'a',newline='',encoding='utf-8-sig') as csvfile: ## 檔名格式，'a+'代表可覆寫        
        writer=csv.writer(csvfile,delimiter=',')
        writer.writerow(['classnumember','classname','teachername','keywordup']) 
        for td in  range(len(result_classname)):
            together=(result_classnum[td],result_classname[td],result_teachername[td],result_teachername[td]+result_classname[td])
            writer.writerow(together)   
    biologycode=""
    web_craeler_step2(access_code,year_val,smester_val,classid_val,biologycode)
def export_csv_language_PEclass_2(result_classnum,result_classname,result_teachername):
    filename_all=file_ENTRY_val+'.csv' #輸出的檔案名
    with open(filename_all, 'a+',newline='',encoding='utf-8-sig') as csvfile: ## 檔名格式，'a+'代表可覆寫        
        writer=csv.writer(csvfile,delimiter=',')
        for td in  range(len(result_classname)):
            together=(result_classnum[td],result_classname[td],result_teachername[td],result_teachername[td]+result_classname[td])
            writer.writerow(together)   
    csvfile.close()
    Csvtoxls(filename_all) #轉換csv to xls 20201225
    tk.messagebox.showinfo(title='提醒', message='csv/xls"精進中英外文/大二三四體育"已儲存完成!，注意該同名檔案會被接著寫下去')
def Csvtoxls(filename_all): #轉換csv to xls 20201225
    for csvfile in glob.glob(os.path.join('.', filename_all)):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('data')
        with open(csvfile, 'rt',encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, val in enumerate(row):
                    ws.write(r, c, val)
        wb.save(file_ENTRY_val+'.xls')
#%%        
#GUI視窗基礎設定
win=tk.Tk()
win.title("彰師小生物開課爬蟲器")
win.geometry("300x300")
win.wm_attributes("-topmost",1) #置頂
radioValue = tk.IntVar() #跟單選有關

title=tk.Label(text="新課表更新程式(Django適應版)",font="微軟正黑體",fg="blue")#標題
title.grid(row=0,column=0,columnspan=2)#標題的位置

#內文說明
tk.Label(text="這個程式是為了幫助新學期課表公布後，").grid(row=1,column=0,columnspan=2)
tk.Label(text="可以立即將課程列表自動更新到網站。").grid(row=2,column=0,columnspan=2)
tk.Label(text="(2021.05.21 ZHAOQI)").grid(row=4,column=0,columnspan=2)

#%% GUI選項
sel_yms_year=ttk.Combobox(win,state="readonly")
sel_yms_year['value']=("112","111","110","109","108","107","106")
tk.Label(text="開課學年").grid(row=5,column=0)
sel_yms_year.grid(row=5,column=1)

sel_yms_smester=ttk.Combobox(win,state="readonly")
sel_yms_smester['value']=("1","2")
tk.Label(text="開課學期").grid(row=6,column=0)
sel_yms_smester.grid(row=6,column=1)

sel_cls_id=ttk.Combobox(win,state="readonly")
sel_cls_id['value']=("核心通識","教育學程","精進中英外文","大二三四體育"
             ,"軍訓課","生一","生二","生三","生四")
tk.Label(text="課程類型").grid(row=7,column=0)
sel_cls_id.grid(row=7,column=1)

biology=tk.IntVar()
tk.Checkbutton(win,text='僅顯示生物系課程',variable=biology,onvalue=1,offvalue=0).grid(row=8,column=0,columnspan=2)

labelframe = tk.LabelFrame(win, text='Excel Frame')#出現一個框框(frame)
labelframe.grid(row=9,column=0,columnspan=2)

tk.Button(labelframe,text="輸出excel檔",command=choice_csv).grid(row=11,column=0)
#按下去按鈕後傳送至choice_csv



win.mainloop() #視窗永久存在