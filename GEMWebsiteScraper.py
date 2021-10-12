from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager as CDM
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from selenium.common.exceptions import NoSuchElementException
import webbrowser
from PyPDF2 import PdfFileReader, PdfFileWriter
import requests
from pathlib import Path
import io
from tabula import read_pdf
from tika import parser
import time
import urllib.request
import pandas as pd
import re
import tabula as tb


driver = webdriver.Chrome(CDM().install())
driver.maximize_window()
page_traverse = 1
c_date = []
o_date = []
p_date = []
title = []
o_name = []
bid_closing_date = []



#j=1

for page_traverse in range(1,5):
    driver.get("https://gem.gov.in/cppp/{}?".format(page_traverse))
    html_source = driver.page_source
    for row_traverse in range(1,11):
        c_date.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[1]".format(row_traverse)).text)
        o_date.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[2]".format(row_traverse)).text)
        p_date.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[3]".format(row_traverse)).text) 
        title.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[4]/a".format(row_traverse)).text)
        o_name.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[5]".format(row_traverse)).text) 



        driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[4]/a".format(row_traverse)).click()



        url = driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[4]/a".format(row_traverse)).get_attribute('href')


        #### To download and close each pdf with file names as pdf1, pdf2, etc ####
        #response = urllib.request.urlopen(URL)    
        #file = open("pdf{}.pdf".format(j), 'wb')
        #file.write(response.read())
        #file.close()
        
        driver.implicitly_wait(5)
        driver.switch_to.window(driver.window_handles[1])
        driver.implicitly_wait(5)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(3)
        driver.implicitly_wait(10)
        
        bid_closing_date.append(driver.find_element(By.XPATH, "/html/body/section[2]/div/div[3]/table/tbody/tr[{}]/td[1]".format(row_traverse)).text) 

        #print(bid_closing_date[j])
        #print(j)
        #j+=1
        
final = zip(c_date, o_date, p_date, title, o_name, bid_closing_date)
#print(list(final))
wb = Workbook()
sh = wb.active

sh.append(["Bid Submission Date", "Tender Opening Date", "e-Published Date", "Title", "Organization Name", "Bid Closing Time"])
for x in list(final):
    sh.append(x)
    
wb.save("Submission.xlsx")    



