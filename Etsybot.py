from selenium import webdriver
# from username import username, password
from selenium.webdriver.common.keys import Keys
import time
#from  openpyxl import *
from xlwt import *
import xlwt
from datetime import datetime

class Etsy:
    def __init__(self):
        self.browser = webdriver.Chrome()
        # self.username = username
        # self.password = password
        self.ck =  Workbook()
        self.k1 =self.ck.add_sheet('sayfa1')
        a = 'Tarih'
        b = 'Saat'
        c = 'Anahtar kelime'

        self.k1.write(0,0,a)
        self.k1.write(0,1,b)
        self.k1.write(0,2,c)

    def signIn(self):
        self.driver = self.browser.get("https://www.etsy.com")
        
        # loginpage = self.browser.find_element_by_xpath("//*[@id='gnav-header-inner']/div[4]/nav/ul/li[1]/button")
        # loginpage.click()
        # time.sleep(5)
        # usernameInput = self.browser.find_element_by_xpath("//*[@id='join_neu_email_field']")
        # passwordInput = self.browser.find_element_by_xpath("//*[@id='join_neu_password_field']")

        # usernameInput.send_keys(self.username)
        # passwordInput.send_keys(self.password)
        # passwordInput.send_keys(Keys.ENTER)
        # time.sleep(2)
        
        
        
    def getWords(self):
        self.wordList = []
        searchbtn = self.browser.find_element_by_xpath("//*[@id='global-enhancements-search-query']")
        searchbtn.click()
        time.sleep(2)
        self.datestring = datetime.now()
        words = self.browser.find_elements_by_css_selector("div[class$='-suggestion']")
        for word in words:
            self.wordList.append(word.text)
        

    def exportExcel(self):
        k1 = self.k1
    
        style1 = xlwt.easyxf(num_format_str='DD.MM.YYYY')
        style2 = xlwt.easyxf(num_format_str='h:mm:ss')
        
        for i in range(self.n-10, self.n):
            k1.write(i+1,0,self.datestring, style1 )
            k1.write(i+1,1,self.datestring, style2 )
            k1.write(i+1,2,self.wordList[i-(self.n-9)])
            



etsy =Etsy()
etsy.signIn()

etsy.n=0

while True:
        etsy.n+=10
        etsy.getWords() 
        etsy.exportExcel()
        #time.sleep(300)
        etsy.browser.refresh()
        etsy.ck.save('etsywords.xls')