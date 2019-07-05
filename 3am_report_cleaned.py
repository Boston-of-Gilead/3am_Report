import arrow #loads time feature
import csv #loads csv reading
import os
import shutil
from shutil import copyfile
import time
import datetime
import pyautogui
import keyboard
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

#Get webdriver from https://selenium-release.storage.googleapis.com/3.9/IEDriverServer_Win32_3.9.0.zip
driver = webdriver.Ie("c:\\users\\<edited>\\desktop\\iedriverserver.exe")

#get webpage
driver.get("<edited>") # Opens webpage

#click login
login = driver.find_element_by_class_name("formButton").click()

#Clicks XXEIS Express Admin button
time.sleep(3)
xxeis = driver.find_element_by_xpath("//*[text()='XXEIS eXpress Administrator']").click()

#Select requests tab
time.sleep(4)
reqtab = driver.find_element_by_id("XXEIS_RSC_ADMIN_REQUESTS").click()

#Change 2nd dropdown to "User"
ddown2 = driver.find_element_by_xpath("//*[text()='User']").click()
time.sleep(2)

#Enter NTID in 3rd field\
time.sleep(2)
ntid = driver.find_element_by_name("SeachText")
ntid.send_keys("<edited>")#NTID

#Generate today's date in desired format of DD-Mon-YYYY
today = datetime.date.today()
today2 = today.strftime("%d-%b-%Y")

#Send formatted date to both start and end date fields
sdate = driver.find_element_by_id("StartDate")
sdate.send_keys(today2)
edate = driver.find_element_by_id("EndDate")
edate.send_keys(today2)

#click Search
searchBtn = driver.find_element_by_id("GoControl").click()

#Now we need to loop through the fourteen results: click on the Output icon id="Summarytable:YO:0" through "Summarytable:YO:13", click id="ExcelAsposeLink" on the following page, and click Back (not return link), i+1
i = 0
for i in range(0,13):
 	outputOf = driver.find_element_by_id("Summarytable:YO:" + str(i) ).click()
 	time.sleep(1)
 	pyautogui.click(3204,1016)
 	excel = driver.find_element_by_id("ExcelAsposeLink").click()
 	time.sleep(4)
 	pyautogui.click(3204,1016)
 	driver.execute_script("window.history.go(-1)")
 	time.sleep(2)
 	i = i+1
#loop works

#make today's date directory out on server:
os.mkdir("\\\\<edited> " + today2)
 
#maintenance macro  \\\\<edited>\\Macro - EiS Maintenance.xls
wb = xw.Book(r'Macro - EiS Maintenance.xls')
wb.macro('Button1_Click')()

#rename them
path = ("C:\\Users\\<edited>\\Desktop\\Downloaded") #use Alex's download location
for filename in os.listdir(path):
    if filename.startswith("ebsprd_Account Alias"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Account Alias Transactions - WC' + '.xlsx'))
    elif filename.startswith("ebsprd_Bin Location"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Bin Location' + '.xlsx'))
    elif filename.startswith("ebsprd_EIS Force"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Force Ship Report - WC' + '.xlsx'))
    elif filename.startswith("ebsprd_Inventory Listing"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Inventory Listing Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Invoice Pre-Register"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Invoice Pre-Register Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Low Qty"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Low Qty Manual Price Change Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Open Purchase"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Open Purchase Orders Listing - WC' + '.xlsx'))
    elif filename.startswith("ebsprd_Open Sales"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Open Sales Orders Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Orders Exception"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Orders Exception Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Receipt Details"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Receipt Details Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Shipped Orders"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Shipped Orders on Hold Report - WC' + '.xlsx'))
    elif filename.startswith("ebsprd_Specials Report"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Specials Report' + '.xlsx'))
    elif filename.startswith("ebsprd_Time"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Time Sensitive Mgmt - WC' + '.xlsx'))
    elif filename.startswith("ebsprd_Inventory - Onhand"):
    	os.rename(os.path.join(path,filename), os.path.join(path,'EiS Item Sub-inventory' + '.xlsx'))
