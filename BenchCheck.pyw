"""
File Name: BechCheck.pyw
Author: Armando Castro
Created on : 7.13.2022

License restrictions: Restricted for use only to WildlifeWonders.com staff.

README:
This script determines if a manufacturer's products are available for
immediate shipment.  It searches for a predetermined list of product names
on a manufacturer's Ecommerce website, to determine the availability of these
products for a reseller.

The script generates an Excel file to hold product informaiton.  This file
can be imported to a Google Doc, or used as is to provide quick access to a
manufacturer's products for a reseller, without having to visit a manufactuer's
website in real time.

Usage instructions:
The script is located inside the "dist" folder. Copy or move script to the folder
you wish to have your excel documents generated in.  Doubleclick script to run silently.
An excel spreadsheet with desired product information will be generated in the
same folder the script is located in.  It is recommended to import this generated
file into a Google sheets document for better visibility and access. 

"""





import pandas as pd
from bs4 import BeautifulSoup
import requests
import datetime
import smtplib
import time
import sys, os, getpass
from datetime import date
from tqdm import tqdm


#Taken from the original Google sheet copy link.  Must create a new link and
#paste the ID here if the original link ID expires or is changed. 
googleSheetId = "1-YM9-Vhj7pWekStgABEyXO5Xbb3viXNCEBENSRuJdpk"

worksheetName = "BenchLinks"


URL = 'https://docs.google.com/spreadsheets/d/{0}/gviz/tq?tqx=out:csv&sheet={1}'.format(googleSheetId, worksheetName)

df = pd.read_csv(URL)


#Drops all columns with no value. (Unused columns and cells)
df = df.dropna(how='all', axis='columns')

#Drops columns 4 5 and 6 from the dataframe.  These are populated with each
#generation of the script. They are recreated to reflect the most recent status.
df = df.drop(df.columns[[4,5,6]], axis=1)

linkArr = []
wildlifeArr = []
availArr = []
matsArr = []
for i in df["Painted Sky Link"]:
    linkArr.append(i)   #Pull links of products from source for BS4


#Search for substring in reseller URL to determine materials used to make
#the current product we are searching for. 
for i in df["WildlifeWondersLink"]:
    wildlifeLink = str(i)
    wildlifeArr.append(wildlifeLink)
    if wildlifeLink.find("2-tone") != -1:
        matWord = "2-Tone"
    else:
        matWord = "Cast Iron"
    matsArr.append(matWord) #New column created with materials used. Allows for
    #sorting of entire spreadsheet by product materials. (Price difference between
    #the two).
    
#For every link in linkArr, use BS4 to open page and then look for the text
#Sold Out in the add to cart button.  
for i in linkArr:
    curLink = str(i)
    curLink = curLink.strip("'")
    page = requests.get(curLink)
    soup = BeautifulSoup(page.content, 'html.parser')
    maybe = repr(soup.find(class_="btn__text"))

    word = "Sold Out"
    if maybe.find(word) != -1:
        availArr.append("Sold Out")
    else:
        availArr.append("Available")

df2 = pd.DataFrame(availArr, columns=['Available?'])
df3 = pd.DataFrame(matsArr, columns=['2-Tone or Cast Iron'])

frames = [df, df2, df3] 
result = pd.concat(frames, axis=1) #combine all dataframes into 1 dataframe 

userName = getpass.getuser()  #pulls user name that is logged into computer.  Also used in path. 
datetime = time.localtime() # pulls date of the action associated.  For file name creation. 
current_time = time.strftime("%H:%M:%S", datetime) # pulls time of the action associated.  For file name creation. 
fixed_current_time = current_time.replace(":",".") #Formats the time output to replace all : with a .

today = date.today()  # Holds todays date in default format Mon-DD-YYYY
addDate = today.strftime("%b-%d-%Y")

fileName = userName + '-' + '-' + addDate + '.' + fixed_current_time + '.xlsx'

with pd.ExcelWriter(str(fileName)) as writer:
    df.to_excel(writer, sheet_name='Original', index=False)
    result.to_excel(writer, sheet_name='Available Benches', index=False)
