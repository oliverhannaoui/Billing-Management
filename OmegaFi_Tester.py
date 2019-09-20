import os
import datetime
import shutil
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait


# Automate login and retrieve balance info
chromedriver = "/usr/local/bin/chromedriver"
driver = webdriver.Chrome(chromedriver)
driver.get("https://login.omegafi.com/users/sign_in")
driver.find_element_by_id('username').send_keys('enter personal email here')
driver.find_element_by_id('password').send_keys('enter personal password here')
driver.find_element_by_name('commit').click()
driver.find_element_by_partial_link_text('Vault').click()
driver.implicitly_wait(10)
driver.find_element_by_link_text('Billing').click()
driver.implicitly_wait(10)
driver.find_element_by_link_text('Billing Roster').click()
driver.implicitly_wait(10)
driver.find_element_by_xpath("//button[@class='btn btn-primary dropdown-toggle btn-sm']").click()
driver.implicitly_wait(10)
driver.find_element_by_xpath("//a[@target='_blank']").click()

# Saving spreadsheet file to correct path
now = datetime.datetime.now()
origin = '/Users/oliverhannaoui/Downloads'
dest = '/Users/oliverhannaoui/Desktop/SigmaChi/Quaestor'
os.chdir('/Users/oliverhannaoui/Downloads')
file_name = 'billing_roster_0'+str(now.month)+'_'+str(now.day)+'_'+str(now.year)+'.xlsx'
remove_file = 'billing_roster.xlsx'
remove_location = '/Users/oliverhannaoui/Desktop/SigmaChi/Quaestor'
path = os.path.join(remove_location, remove_file) 
os.remove(path) 
os.rename(file_name,'billing_roster.xlsx')
shutil.move(origin+'/billing_roster.xlsx',dest)

# Parsing excel sheet and creating dictionary
excel_sheet = 'billing_roster.xlsx'
df = pd.read_excel(excel_sheet,usecols = [0,5], skiprows = 4)
df_dict = df.set_index('Last Name')['Aging Status'].to_dict()
print(df_dict)