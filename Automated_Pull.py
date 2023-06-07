from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import win32com.client
import re 
import time
import glob
from datetime import datetime, timedelta , date
import pandas as pd
from pydomo import Domo

#Set to always download file and other options
options=webdriver.ChromeOptions()
prefs = {'directory'}
options.add_experimental_option('prefs',prefs)
driver = webdriver.Chrome(chrome_options=options)

#Load Lowes Portal
driver.get('website')
wait = WebDriverWait(driver,10)
#Login
Email = wait.until(EC.presence_of_element_located((By.ID, 'idToken1')))
Email.send_keys('username')

Password = driver.find_element(By.ID,'idToken2')
Password.send_keys('password')

Login = driver.find_element(By.ID,'loginButton_0')
Login.click()

#One Time Password
time.sleep(60)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
message = messages.GetLast()
body_content = message.body
digets = re.findall('\d+',body_content)
otp = digets[2]

One_time_password = driver.find_element(By.ID,'idToken2')
One_time_password.send_keys(otp)

Login2 = driver.find_element(By.ID, 'idToken3_0')
Login2.click()

#Navigate Reporting
time.sleep(15)

Dropdown_menu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[name()="svg"]')))
Dropdown_menu.click()

Saved_Reports = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@class="mgMainMenu"][7]/a')))
Saved_Reports.click()

exit = wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div/div[1]/div/div/div[1]/div[1]/div[3]/button')))
exit.click()

report1 = wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[1]/div[2]/div/div[1]/div/a')))
report1.click()

download = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div/div/div[1]/div[2]/button")))
download.click()

time.sleep(5)
driver.close()


#Get Path for most recent upload
path = r"C:\Users\Keaton\Documents\Lowes VPP"
filenames = glob.glob(path + "\*.xlsx")
numbers = []
for i in filenames:
    number = re.findall('\d+',i)
    numbers.append(number)

max_gen = 0
max_fore = 0 
for file in filenames:
    if file.__contains__('Demand Forecast'):
        try:
            if int(re.findall('\d+',file)[0]) > max_fore:
                max = int(re.findall('\d+',file)[0])
        except:
            pass
    if file.__contains__('General Sales'):
        try:
            if int(re.findall('\d+',file)[0]) > max_gen:
                max = int(re.findall('\d+',file)[0])
        except:
            pass


#Get last friday for date
day = timedelta(days=1)
curr_day = date.today()
if curr_day.weekday() == 4:
    curr_day - timedelta(days=7)
while curr_day.weekday()  != 4:
    curr_day -= day

#create general sales DataFrame
general_sales = pd.read_excel('General Sales ('+str(max_gen)+').xlsx',header= 9)
general_sales['Week ID'] = curr_day



#Append to Domo Dataset
domo = Domo('domo_ds_id','api_key',api_host='api.domo.com')
domo.ds_update(ds_id='domo_ds_id',df_up=general_sales)

