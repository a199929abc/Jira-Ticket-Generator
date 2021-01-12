"""
Author:Kaiheng Zhang
Mail : kaiheng365@gmail.com
Date: 2021-01-08
Version:2.0.2
"""
# Excel libraries
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import openpyxl
import pandas as pd
import numpy as np
from onc.onc import ONC
import json
import requests
# GUI libraries
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter.filedialog import askopenfilename
import tkinter.scrolledtext as tkst
from tkinter import messagebox as mb
from tkinter import filedialog
import tkinter.font as tkFont

# JIRA libraries
from jira.resources import IssueLink
import jira.client
from jira.client import JIRA
from jira import JIRA
import re

#Variables
username = ''
password = ''
mypath = ''
workbookTitle = ''
loginWindow = ''
GUI_flag= False
versionControl= '2.0.2'

# create a dataframe for storing the deploy row
df_out = pd.DataFrame()

# create a dictionary for storing the site ticket EN number
site_dict = {}

def openFile():
    """ Open File explorer and lets user select exsisting excel workbook and worksheet to be used """
    global mypath
    global workbookTitle
    mypath = filedialog.askopenfilename(initialdir = "C:",
                           filetypes = (("Excel Workbook", "*.xlsx"), ("Excel Macro-Enabled Workbook","*.xlsm")),
                           title = "Choose a file."
                           )

    workbookTitle = os.path.basename(mypath)
    wbkTitle = StringVar()
    wbkTitle.set(workbookTitle)
    
  #  Entry(labelframe1, textvariable = wbkTitle, state = DISABLED, width = 35, font = 'bold').place(x = 20, y = 25)
    print(wbkTitle.get())
def create_ticket(row,Instrument_Category,Instrument,Serial_Number):
    # Assign the values
    __SerialNumber = Serial_Number
    __Instrument=   Instrument
    __InstrumentCategory = Instrument_Category
    __DeviceID =      row['DeviceID']
    __SiteLocation = row['Site/Location']
    __TicketLink=   row['Ticket Link']
    __createdTicket=row['Created Ticket']
    __components=   row['Component'] 
    __outwardIssue = row['Linked To']
    __issueLink = row['Work Ticket']
    ''' if(row['Operation'] == 'Deploy'):
     __summaryTitle = 'Instrument Qualification'
    if(row['Operation'] == 'Recover'):
        __summaryTitle = 'Instrument Recovery'''

    # Connect to jira
    # Authentication done by using username and password
    username = 'mtcelec2'
    password = '1q2w3e4R!'

    jira = JIRA(
        basic_auth = (username, password),
        options = {'server': 'http://142.104.193.65:8080'}
        #options = {'server': 'https://jira.oceannetworks.ca/'}
    )

    #create a ticket for current row
    new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary ="'{0}',SI: '{1}',DI: '{2}'".format(__Instrument,__SerialNumber, __DeviceID),
        # "%s: %s SI: %s DI: %s" % ( __Instrument, __SerialNumber, __DeviceID),
        description = "Site Location: %s\n  Instrument Category: %s\n " % (__SiteLocation,  __InstrumentCategory), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID)                 # Device ID field

    # add the linkedto
    if(isinstance(row['Linked To'], str)):
        jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)

    return new_issue.key
def oncAPIget(row):
    Instrument_Category=''
    Instrument=''
    Serial_Number= ''
    deviceId=row['DeviceID']

    url = 'https://data.oceannetworks.ca/api/devices'
    parameters = {'method':'get',
            'token':'71f23a7a-8b7f-4b13-bd24-0948bc76eab0', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
            'deviceId':deviceId}
  
    response = requests.get(url,params=parameters)
  
    if (response.ok):
        devices = json.loads(str(response.content,'utf-8')) # convert the json response to an object
        for device in devices:
            Instrument=device.get('deviceName')
            Instrument_Category=device.get('deviceCategoryCode')
            

    else:
        if(response.status_code == 400):
            error = json.loads(str(response.content,'utf-8'))
            print(error) # json response contains a list of errors, with an errorMessage and parameter
        else:
            print ('Error {} - {}'.format(response.status_code,response.reason))

    
    
    SNtemp = Instrument.split()
    if len(SNtemp)<=1:
        Serial_Number=None
    else:
        Serial_Number=SNtemp[-1]
    return Instrument_Category, Instrument, Serial_Number
    

if __name__=='__main__':
    pos = 0
    index=0

    currSite_Location= ''
    currDeviceID = ''
    currTicketLink= ''
    currInstrumentCategory=''
    currInstrument=''
    currSerialNumber=''
    currCreatedTicket=''
    df_whole = pd.read_excel (r'C:\Users\mtcelec2\Desktop\kaiheng\JIRA_Auto\Test_Jira\simple_jira.xlsx')
    #print(df_whole.shape)
    #print(len(df_whole.keys()))
    df_whole.drop(df_whole.iloc[:, 7::], inplace = True, axis = 1)
    df_whole.insert(7, "rowNum",np.nan)
    df_whole.columns=['Site/Location','DeviceID','Ticket Link','Instrument Category','Instrument','Serial Number','Created Ticket','rowNum']
    df_whole.insert(8, "Component", "Test and Development")
    df_whole.insert(9, "Linked To", np.nan)
    df_whole.insert(10,"Work Ticket",np.nan)

    for index, row in df_whole.iterrows():
        Instrument_Category=''
        Instrument=''
        Serial_Number= ''
        Instrument_Category, Instrument, Serial_Number= oncAPIget(row)
        myKey = create_ticket(row,Instrument_Category, Instrument, Serial_Number)
        df_whole['rowNum'][index]=pos
        df_whole['Instrument Category'][index]=Instrument_Category
        df_whole['Instrument'][index]=Instrument
        df_whole['Serial Number'][index]=Serial_Number
        df_whole['Created Ticket'][index]=myKey
        df_whole['Ticket Link'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        print("Finished Create Ticket"+'\n')
        pos+=1

        

    df_whole.drop(df_whole.iloc[:, 8::], inplace = True, axis = 1)
    df_whole.to_excel("output_test.xlsx", sheet_name='S1',index=False) 
    


        
  
    

    
    



    

    
