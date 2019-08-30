from selenium import webdriver
from bs4 import BeautifulSoup
import random,numpy as np,pandas as pd,os
from selenium.common.exceptions import NoSuchElementException
from StocksFilter.bar import showBar  #匯入自訂進度條函數

Options=webdriver.ChromeOptions()
Options.add_argument("headless")
browser1 = webdriver.Chrome(chrome_options=Options)
browser1.get("http://goodinfo.tw/StockInfo/StockList.asp")
svCookies=browser1.get_cookies()
stkmenu=browser1.find_elements_by_css_selector("#txtStockListMenu a")
#股票menu字典
stkCatas={}
for i,ele in enumerate(stkmenu):
    ele=ele.text.strip()
    if(ele=="全部類股"):
        continue
    elif(ele=="00"):
        break
    print("{:<2}:{:.13}".format(i,ele)) if (i%4==0) else print("{:<2}:{:.13}".format(i,ele),end=" ")
    stkCatas[str(i)]=ele
print()
#print(stkCatas)
#選取股票類別與連結網址
snum=input("請輸入類股代號:")
#s="2"
print("已選擇:",stkCatas[snum])

#browser1.get("http://goodinfo.tw/StockInfo/StockList.asp")
browser1.implicitly_wait(random.random()*2+1)
browser1.find_element_by_partial_link_text(stkCatas[snum]).click()
browser1.delete_all_cookies()
browser1.add_cookie(svCookies[0])
stkInCata=browser1.find_elements_by_css_selector(".solid_1_padding_4_1_tbl:nth-child(1) td:nth-child(2) a")                                                 
stockDict={}
for ele in stkInCata:
    if(ele.text.strip()!="名稱"):
        shref=ele.get_property("href")
        ele=ele.text
        stockDict[shref[-4:]]=ele
print(stockDict)
totalStk=len(stockDict)
if (totalStk==0):
    print("沒有取到類別下之股票")
    browser1.quit()
    os._exit(0)
browser1.quit()

browser2 = webdriver.Chrome(chrome_options=Options)
browser2.get("http://www.cmoney.tw/finance/f00025.aspx")
savCookies2=browser2.get_cookies()

browser3 = webdriver.Chrome(chrome_options=Options)
browser3.get("http://www.cmoney.tw/finance/f00025.aspx")
savCookies3=browser3.get_cookies()

browser4 = webdriver.Chrome(chrome_options=Options)
browser4.get("http://www.cmoney.tw/finance/f00025.aspx")
savCookies4=browser4.get_cookies()
#browser1.implicitly_wait(random.random()*2+1)
#篩選出符合條件的股票
#1.自由現金流量5年皆大於0;2.現金股利近5年平均大於0;3.ROE近5年平均大於10%;
#寫入xlsx檔案的資料列
Datas,DataRow=[],[]
#累計處理數量
cumNum=0
for xid in stockDict:
    print("目前處理:"+xid)
    remark=" " #備註
    cumNum+=1
    showBar("篩選股票",totalStk,cumNum)
    try:
        browser2.get("http://cmoney.tw/finance/f00042.aspx?s="+xid+"&o=4")
        browser2.delete_all_cookies()
        browser2.add_cookie(savCookies2[0])
        for i,ele in enumerate(savCookies2):
            browser2.add_cookie(savCookies2[i])
        soup= BeautifulSoup(browser2.page_source,"html.parser")
        #網頁table第8列資料,自由現金流量
        sp=soup.select(".tb-out tr:nth-child(8)")
        FCF_ROW=sp[0].find_all("td")
        FCF5year,FCF5yearAvg=[],0
        if((len(FCF_ROW)>=6) and (FCF_ROW[0].text.strip()=="自由現金流量")): #第一欄為"自由現金流量"
            for i in range(1,6):    #取5年的現金流量皆大於0
                #取代金額逗號
                FCF_NUM=eval(FCF_ROW[i].text.replace(",",""))
                if FCF_NUM>=0:
                    FCF5year.append(FCF_NUM)  
            if(len(FCF5year)==5):
                FCF5yearAvg=np.average(FCF5year)
            else:
                continue
        else:
            print(xid,":自由現金流量格式異常")
            remark+=(xid+"自由現金流量不足5年")
    except NoSuchElementException as e:
        print(e)
        browser2.quit()
    #現金股利近5年平均大於0 Cash dividend
    try:
        #browser3.implicitly_wait(random.random()*3+1)
        browser3.get("http://www.cmoney.tw/finance/f00027.aspx?s="+xid)
        browser3.delete_all_cookies()
        for i,ele in enumerate(savCookies2):
            browser3.add_cookie(savCookies3[i])
        soup= BeautifulSoup(browser3.page_source,"html.parser")
        spTitle=soup.select("table th:nth-child(2)")
        spCont=soup.select("table tr td:nth-child(2)")
        Cash5year,Cash5yearAvg=[],0
        if ((spTitle[0].text.strip()=="現金股利") and (len(spCont)>=5)):
            for i in range(0,5):      #5年的現金股利平均
                cash_NUM=eval(spCont[i].text)
                Cash5year.append(cash_NUM)
            Cash5yearAvg=np.average(Cash5year)
            if (Cash5yearAvg<=0):
                continue
        else:
            print(xid,":現金股利格式異常")
            remark+=(xid+"現金股利年數不足5年")
    except NoSuchElementException as e:
        print(e)
        browser3.quit() 
    #ROE近5年平均大於10%
    try:
        #browser4.implicitly_wait(random.random()*3+1)
        browser4.get("http://www.cmoney.tw/finance/f00043.aspx?s="+xid+"&o=3")
        browser4.delete_all_cookies()
        for i,ele in enumerate(savCookies4):
            browser2.add_cookie(savCookies4[i])
        soup= BeautifulSoup(browser4.page_source,"html.parser")
        sp=soup.select(".tb-out tr:nth-child(6)")
        ROE_ROW=sp[0].find_all("td")
        ROE5year,ROE5yearAvg=[],0    
        if(ROE_ROW[0].text.strip()=="稅後股東權益報酬率" and len(ROE_ROW)>=6):  
            for i in range(1,6):
                #取代數字負號
                if (ROE_ROW[i].text.replace("-","").replace(".","").isdigit()): #欄位是數字
                    ROE5year.append(eval(ROE_ROW[i].text))
            ROE5yearAvg=np.average(ROE5year) 
            if((len(ROE5year)<5) or (ROE5yearAvg < 10)):
                continue   
        else:
            print(xid,":稅後股東權益報酬率格式異常")
            remark+=(xid+"稅後股東權益報酬率不足5年")
    except NoSuchElementException as e:
        print(e)  
        browser4.quit()   
    #篩股條件:
    #1.自由現金流量5年皆大於0;2.現金股利近5年平均大於0;3.ROE近5年平均大於10%;
    DataRow=[stockDict[xid],xid,FCF5yearAvg,Cash5yearAvg,ROE5yearAvg,remark]
    Datas.append(DataRow)
    print("符合股票:"+xid)
    
      
browser2.close() 
browser2.quit()
browser3.close() 
browser3.quit()
browser4.close() 
browser4.quit()
columns=["個股名稱","個股代號","自由現金流量5年平均(元)","現金股利5年平均(元)","ROE5年平均(%)","remark"]
df=pd.DataFrame(Datas,columns=columns)
fileName="stkFilter"+stkCatas[snum]+".xlsx"
if(len(Datas)>0):
    df.to_excel(fileName,sheet_name="篩選好股",index=False,encoding="utf8")
    print(os.getcwd()+fileName,"檔案已產出")
else:
    print("沒有符合條件的股票")
