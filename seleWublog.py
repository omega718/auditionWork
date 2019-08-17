from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import pandas as pd
from bs4 import BeautifulSoup as bsp
browser = webdriver.Chrome("chromedriver.exe")
browser.get("http://www.google.com/")
browser.find_element_by_name("q").send_keys(u"吳老師教學部落格")
browser.find_element_by_name("q").send_keys(Keys.ENTER)
browser.find_element_by_xpath(u"(.//*[normalize-space(text()) and normalize-space(.)='附有網站連結的網頁搜尋結果'])[1]/following::div[5]").click()
soup= bsp(browser.page_source,"html.parser")
browser.close()
browser.quit()
content1,content2,content3=[],[],[] 
for i,ele in enumerate(soup.select(".post-outer ul li a")):
        content1.append(ele.text)
        content2.append(ele.get("href"))
        content3.append('=HYPERLINK("' + ele.get("href") + '","link")')
df = pd.DataFrame({'標題':content1,'網址':content2,'連結':content3})
df.to_excel('吳老師部落格教學連結.xlsx',index=False)
print("吳老師部落格教學連結.xlsx ,檔案寫入完畢")
