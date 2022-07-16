import pandas as pd
import requests
import datetime
import smtplib
import time
import sys, os, getpass
from datetime import date
from tqdm import tqdm
from progressbar import ProgressBar
pbar = ProgressBar()


#Taken from the original Google sheet copy link.  Must create a new link and
#paste the ID here if the original link ID expires or is changed. 
googleSheetId = "1eabOF2Jb8fozVN3iSXEvt6tsFc_ZWYD9B-o3zjA6dOQ"

worksheetName = "SkuLinks"


URL = 'https://docs.google.com/spreadsheets/d/{0}/gviz/tq?tqx=out:csv&sheet={1}'.format(googleSheetId, worksheetName)

df = pd.read_csv(URL)

df = df.dropna(how='all', axis='columns')

df2 = df["Product URL"]

existsArr = []

prefixURL  = r"https://wildlifewonders.com"
df2Range = len(df2)

for i in pbar(range(df2Range)):
    try:
        suffix = str(df2[i])
        suffix = suffix.strip("'")
        complete = prefixURL + suffix

        response = requests.get(complete)
        if response.status_code == 200:
            existsVar = 'Yes'
            existsArr.append(existsVar)
        else:
            existsVar = 'No'
            existsArr.append(existsVar)
    except:
        existsVar = "Failed"

df3 = pd.DataFrame(existsArr, columns=['Website Exists?'])

frames = [df, df3]
result = pd.concat(frames, axis=1)


userName = getpass.getuser()  #pulls user name that is logged into computer.  Also used in path. 
datetime = time.localtime() # pulls date of the action associated.  For file name creation. 
current_time = time.strftime("%H:%M:%S", datetime) # pulls time of the action associated.  For file name creation. 
fixed_current_time = current_time.replace(":",".") #Formats the time output to replace all : with a .

today = date.today()  # Holds todays date in default format Mon-DD-YYYY
addDate = today.strftime("%b-%d-%Y")

fileName = userName + '-' + '-' + addDate + '.' + fixed_current_time + '.xlsx'

with pd.ExcelWriter(str(fileName)) as writer:
    result.to_excel(writer, sheet_name='Valid URLS', index=False)
