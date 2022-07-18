# Ncuehulolo_crawler
Ncuehulolo_crawler 是一款用於輸出彰師 (NCUE) 當學期開課清單的GUI程式，利用爬蟲技術以及tk GUI套件構成，使用者僅需在輸出的程式中選擇對應的學年、學期、課程類型以及檔名，即可獲得一份開課清單(.csv & .xls)。
目前程式主要在[彰師小生物](https://ncuehulolo.idv.tw/general/ "彰師小生物")被使用，用於更新當學期課表讓同學得知該課是否有對應的評價。

操作時建議使用```ncueclaavalue_gui_forteam.py```，其餘是先前迭代的檔案。

## 簡介

- 操作介面

![操作介面](https://user-images.githubusercontent.com/81002817/179436216-ce0df2f4-c29a-43c5-a278-3d19e9529b42.png)

- 輸出檔案

![輸出檔案](https://user-images.githubusercontent.com/81002817/179436668-52d62b98-2268-4452-85c3-6d3ede1d991d.png)


1. 三名老師以上開課，老師名部分一律會顯示【合開】。
1. 課程類型選項如下。
![課程類型選項](https://user-images.githubusercontent.com/81002817/179436856-f27862f6-3563-4ee4-943a-7cf5a448f8c8.png)

## 進階參數調整

- 開課學年新增選項

	在此行增加新的年份即可。

```python
sel_yms_year['value']=("112","111","110","109","108","107","106")
```
