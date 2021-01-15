"""
Author:Kaiheng Zhang
Mail : kaiheng365@gmail.com
Date: 2021-01-08
Version:2.0.2
"""
# Excel libraries
from request_jira import *
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
from request import *
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
    
'''def create_ticket(row,instrument_category,instrument,serial_number):
    # Assign the values
    __SerialNumber = serial_number
    __instrument=   instrument
    __instrumentcategory =instrument_category
    __DeviceID =      row['DeviceID']
  #  __SiteLocation = row['Site/Location']
    __TicketLink=   row['Ticket Link']
    __createdTicket=row['Created Ticket']
    __components=   row['Component'] 
    __outwardIssue = row['Linked To']
    __issueLink = row['Work Ticket']
    if(row['Operation'] == 'Deploy'):
     __summaryTitle = 'instrument Qualification'
    if(row['Operation'] == 'Recover'):
        __summaryTitle = 'instrument Recovery'

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
        summary ="'{0}',SI: '{1}',DI: '{2}'".format(__instrument,__SerialNumber, __DeviceID),
        # "%s: %s SI: %s DI: %s" % ( __instrument, __SerialNumber, __DeviceID),
        description = " instrument category: %s\n " % ( __instrumentcategory), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID,
        duedate="2021-01-15",
        assignee={'name': 	'around'}
        #assignee format take only name before the email
        #assignee={'sfaassfda'}
        #https://innovalog.atlassian.net/wiki/spaces/JMWEC/pages/108200050/Standard+JIRA+fields Very helpful link 
        )
    #print(new_issue.fields.status.name) 
    ##print(new_issue.fields.issuetype.name) 
    #print(new_issue.fields.reporter.displayName)
    #print(new_issue.fields.summary)
    #print(new_issue.fields.comment.comments)    
                     # Device ID field

    # add the linkedto
    if(isinstance(row['Linked To'], str)):
        jira.create_issue_link("Related", new_issue.key, __outwardIssue, None) 
    print(new_issue.key)

    return new_issue.key   
def check_status(ticket):
    username = 'mtcelec2'
    password = '1q2w3e4R!'
    jira = JIRA(
    basic_auth = (username, password),
    options = {'server': 'http://142.104.193.65:8080'}
    #options = {'server': 'https://jira.oceannetworks.ca/'}
                )
    new_issue = jira.issue(ticket)
    return new_issue.fields.status.name
   #
    print(new_issue.fields.status.name) 
    print(new_issue.fields.issuetype.name) 
    print(new_issue.fields.reporter.displayName)
    print(new_issue.fields.summary)
    print(new_issue.fields.duedate) '''

if __name__=='__main__':
    pos = 0
    index=0

    currSite_Location= ''
    currDeviceID = ''
    currTicketLink= ''
    currinstrumentcategory=''
    currinstrument=''
    currSerialNumber=''
    currCreatedTicket=''
    df_whole = pd.read_excel (r'C:\Users\mtcelec2\Desktop\kaiheng\JIRA_Auto\Test_Jira\jira_input_v2.1.0.xlsx')
    
    #print(len(df_whole.keys()))
    df_whole.drop(df_whole.iloc[:, 10::], inplace = True, axis = 1)
    df_whole.insert(10, "rowNum",np.nan)

    df_whole.columns=['DeviceID','Due Date','Assignee','Description','Ticket Link','Instrument Category',
    'Instrument','Serial Number','Created Ticket','status','rowNum']

    df_whole.insert(11, "Component", "Test and Development")
    df_whole.insert(12, "Linked To", np.nan)
    df_whole.insert(13,"Work Ticket",np.nan)
    df_whole = df_whole[:-1]
    #print(df_whole.shape)
    #print(df_whole.head)
    #df_whole.drop(df_whole.iloc[:, 8::], inplace = True, axis = 1)
    for index, row in df_whole.iterrows():
        local_instrument_category=''
        local_instrument=''
        serial_number= ''
        local_instrument,local_instrument_category=onc_request(row)
        local_instrument, serial_number = processString(local_instrument)  
        df_whole['rowNum'][index]=pos
        df_whole['Instrument Category'][index]=local_instrument_category
        df_whole['Instrument'][index]=local_instrument
        df_whole['Serial Number'][index]=serial_number
        pos+=1

        ''' myKey = create_ticket(row,local_instrument_category, local_instrument, serial_number)
        df_whole['Created Ticket'][index]=myKey
        df_whole['Ticket Link'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        print("http://142.104.193.65:8080/browse/%s" % myKey)'''
        status=check_status('EN-54071')
        
   
        #print("Finished Create Ticket"+'\n')
        
        

        


    #df_whole.drop(df_whole.iloc[:, 8::], inplace = True, axis = 1)
    #df_whole.to_excel("output_test32.xlsx", sheet_name='S1',index=False) 
    


        
  
    

    
    



    

    
