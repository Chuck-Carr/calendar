import requests
from bs4 import BeautifulSoup as bs
from selenium import webdriver
import csv
import time
import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from secrets import *

# import html5lib

#### use webdriver to log into salesforce website ####
# driver = webdriver.Chrome()
driver = webdriver.Chrome(executable_path="Chromedriver")
driver.get('https://na77.salesforce.com/a0o?fcf=00B1M000009oC5w&rolodexIndex=-1&page=1')
username = salesforce
password = sf_pass
driver.find_element_by_id("username").send_keys(username)
driver.find_element_by_id("password").send_keys(password)
driver.find_element_by_id("Login").click()
#### Pause to allow website to load ####
time.sleep(15)

## Extract data from the table and save as csv ####
html = driver.page_source
tables = pd.read_html(html, attrs={'class':'x-grid3-row-table'})
pd.set_option('display.max_columns', None)
df = pd.concat(tables, ignore_index=True)
df = df.drop([0,9,10,11,7], axis=1)
df = df.rename({1: 'Work Order', 2:'Acct', 3:'Street', 4:'City', 5:'State', 6:'Zip', 8:'Scheduled Date'}, axis=1)
df.to_csv('table.csv', index=False)

### Close the browser ####
driver.close()

# #### Use Panda to compare new and old csv files ####
f1 = pd.read_csv("table.old.csv", encoding = 'ISO-8859-1')
f2 = pd.read_csv("table.csv", encoding = 'ISO-8859-1')
#### Comparing new and old csv files to determine newly assigned work orders ####

def dataframe_difference(df1, df2):
   ### returning rows that are only in new table ###
    comparison_df = df1.merge(df2,indicator=True,how='outer')
    diff_df = comparison_df[comparison_df['_merge'] == 'left_only']
    diff_df = diff_df.drop(['_merge'], axis=1)
    diff_df.to_csv('update.csv', index=False)
    return diff_df
dataframe_difference(f2, f1)


#### Saving new csv as old.csv for future comparisons ####
f2.to_csv('table.old.csv', index=False)

diff_df = pd.read_csv('update.csv', encoding= 'ISO-8859-1')

#### If any new work orders have been assigned send email notification ####
if os.path.getsize('update.csv') > 64:
    email_address = send_email
    email_password = send_pass
    msg = EmailMessage()
    msg['Subject'] = f'You have been assigned {len(diff_df)} new work orders.'
    msg['From'] = 'NEW WORK ORDERS'
    msg['To'] = yahoo_email
    msg.set_content(f'You have been assigned {len(diff_df)} new work orders.')
    files = ['update.csv']
    for file in files:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = f.name        
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_address, email_password)
        smtp.send_message(msg)








