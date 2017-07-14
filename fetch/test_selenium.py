from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from os import rename, listdir
import os


import requests
import bs4
import time
import shutil
import glob

site = 'http://122.154.18.61/'
login_site = 'http://122.154.18.61/logon.aspx'
refer_site = 'http://122.154.18.61/neon-network.aspx'

#Choose Driver
#driver = webdriver.Chrome('driver/chromedriver.exe')
#driver = webdriver.PhantomJS('driver/phantomjs-2.1.1-windows/phantomjs-2.1.1-windows/bin/phantomjs.exe')

#Set Directory Path
chromeOptions = webdriver.ChromeOptions()
# prefs = {"download.default_directory" : "C:/Users/asus/Desktop/Neon"}
# chromeOptions.add_experimental_option("prefs",prefs)

chromeOptions.add_experimental_option("prefs", {
  "download.default_directory": r"root/data", #<-------------------------- set directory
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

chromedriver = "driver/chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)


expand = [247,250,252,270,272,275,277,279,285,288,290,292,294,296,300,302,305,330,333,340]


date = [
		 # ('1/01/2010' , '31/12/2010') ,
		 # ('1/01/2011' , '31/12/2011') ,
		 # ('1/01/2012' , '31/12/2012') , 
	 	#  ('1/01/2013' , '31/12/2013') , 
		 # ('1/01/2014' , '31/12/2014') , 
		 # ('1/01/2015' , '31/12/2015') , 
		 ('1/01/2016' , '31/12/2016') ,  
		 ('1/01/2017' , '31/12/2017') , 
	   ]


#Start
print("Run Driver")
driver.get(login_site)
print(driver.title)
driver.find_element_by_id("txtUsername").send_keys("guest")
driver.find_element_by_id("txtPassword").send_keys("ridview")
driver.find_element_by_id('btnLogon').click()
driver.get(refer_site)
print(driver.title)



for i in xrange(7,398):	
	count = 0
	for j in expand:
		if i == j:
			driver.find_element_by_xpath('//*[@id="ctlMenu_ctlTreeMenu_treeMenun' + str(j) + '"]/img').click()
			print ("id = ctlMenu_ctlTreeMenu_treeMenut " + str(j) + " expand ok")
	if i in expand:
		continue	
	
	driver.find_element_by_id('ctlMenu_ctlTreeMenu_treeMenut' + str(i)).click() #<------ select left_menu 

	try:
		element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'gridChannels_ctl02_hrefChannel')))
		element.click()
		driver.find_element_by_id('rdoDisplayType_1').click()
		element2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'chkLoggerChannels')))
		element2.click()
		driver.find_element_by_id('btnGetData').click()
		title = driver.find_element_by_id('ctlNodeTabs_lblPageHeading').text
		title = title.replace(" ", "_").replace(":", "_").replace(".", "_")
		#print ("name is " + name.encode("utf-8"))
		for x , y in date: # <---------- filter date
			year = x[5:]

			driver.find_element_by_id("dateStart_txt_Date").clear()
			driver.find_element_by_id("dateEnd_txt_Date").clear()
			startTime = Select(driver.find_element_by_id('lstStartTime'))
			startTime.select_by_value('0') # 0.00 - 1.00
			endTime = Select(driver.find_element_by_id('lstEndTime'))
			endTime.select_by_value('23') # 23.00 - 24.00
			driver.find_element_by_id("dateStart_txt_Date").send_keys(str(x))
			driver.find_element_by_id("dateEnd_txt_Date").send_keys(str(y))
			driver.find_element_by_id('btnGetData').click()
			driver.find_element_by_id('btnExportToExcel').click() # <-------- download csv file 

			time.sleep(5)

			for filename in os.listdir(path): # <---------- check and rename file 
				if filename == "Item_Data.csv":
					old_file = "C:/Users/asus/Desktop/Neon"+os.path.sep+"Item_Data.csv"
					new_file = "C:/Users/asus/Desktop/Neon"+os.path.sep+ title + "-" + str(year) + ".csv"
					#new_file = "C:/Users/asus/Desktop/Neon"+os.path.sep+"Item_Data_ID" + str(i) + "-" + str(year) + ".csv"
					os.rename(old_file,new_file)
					count+=1
					break	
			
	except TimeoutException:
		continue
	




#---------------------------------------------------------------------------

	#print (i.get_attribute('id')) 
	#print(menu.text.encode("utf-8"))
	# driver.find_element_by_id('ctlMenu_ctlTreeMenu_treeMenut3').click()
	# driver.find_element_by_id('gridChannels_ctl02_hrefChannel').click()
	# driver.find_element_by_id('rdoDisplayType_1').click()
	# driver.find_element_by_id('chkLoggerChannels').click()
	# driver.find_element_by_id('btnGetData').click()
	#title = driver.find_element_by_id('ctlNodeTabs_lblPageHeading').text
	# table = driver.find_element_by_id('gridData')
	#print(title.encode("utf-8"))
	#print(table.text.encode("utf-8"))
	#print(time() - st)
