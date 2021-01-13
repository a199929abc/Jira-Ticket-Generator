"""
Author:Kaiheng Zhang
Mail : kaiheng365@gmail.com
Date: 2021-01-13
Version:VersionControl
"""
from onc.onc import ONC
import json
import requests
# Excel libraries
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import openpyxl
import pandas as pd
import numpy as np

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
versionControl= '2.0.3'
df_whole= pd.DataFrame()
# create a dataframe for storing the deploy row
df_out = pd.DataFrame()
# create a dictionary for storing the site ticket EN number
site_dict = {}

def save_textvariable():
    global username
    username = e1.get()
    global password
    password = e2.get()
    global GUI_flag
    try:
        jira = JIRA( 
            basic_auth = (username, password),
            options = {'server': 'http://142.104.193.65:8080'}
            #options = {'server': 'https://jira.oceannetworks.ca/'}
        )
        
        GUI_flag=True
        log_inWindow.destroy()
    except:
        mb.showerror("Error", "Login Unsuccessfull")
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
    Entry(labelframe1, textvariable = wbkTitle, state = DISABLED, width = 35, font = 'bold').place(x = 20, y = 25)
    print(wbkTitle.get())
def create_ticket(row,Instrument_Category,Instrument,Serial_Number):  
    # Assign the values
    #Ticket create row by row. This part is adding all the require information to the JIRA API
    #instrument, instrument Category, serial Numbers are get from the ONC API 
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
    #password can be set as user input in the future 
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
    #set variable you want to get from the web service and provide the data you provided from Excel
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
            #if the response success, you will receive the required info here
            Instrument=device.get('deviceName')
            Instrument_Category=device.get('deviceCategoryCode')
            

    else:
        if(response.status_code == 400):
            error = json.loads(str(response.content,'utf-8'))
            print(error) # json response contains a list of errors, with an errorMessage and parameter
        else:
            print ('Error {} - {}'.format(response.status_code,response.reason))

    
    #Since there isn't exactlly cell for serial number, they are all at the end of the instrument name
    #we need to process the string to get serial number
    SNtemp = Instrument.split()
    if len(SNtemp)<=1:
        Serial_Number=None
    else:
        Serial_Number=SNtemp[-1]
    #return data we want to get 
    return Instrument_Category, Instrument, Serial_Number
def processExcel():
    global mypath
    pos = 0
    index=0
    global df_whole
    df_whole = pd.read_excel (mypath)
    #print(df_whole.shape)
    #print(len(df_whole.keys()))
    #drop unused columns of the dataframe
    df_whole.drop(df_whole.iloc[:, 7::], inplace = True, axis = 1)
    #insert the cloumn to count the number
    df_whole.insert(7, "rowNum",np.nan)
    #set cloumns name to the dataframe
    df_whole.columns=['Site/Location','DeviceID','Ticket Link','Instrument Category','Instrument','Serial Number','Created Ticket','rowNum']
    #In JIRA API it requires component to fill in the ticket
    df_whole.insert(8, "Component", "Test and Development")
    #Linked To and work Ticket to show the relation of the ticket
    df_whole.insert(9, "Linked To", np.nan)
    df_whole.insert(10,"Work Ticket",np.nan)
    #process each ROW each iteration 
    for index, row in df_whole.iterrows():
        Instrument_Category=''
        Instrument=''
        Serial_Number= ''
        #using JIRA API get info required
        Instrument_Category, Instrument, Serial_Number= oncAPIget(row)
        #myKey = create_ticket(row,Instrument_Category, Instrument, Serial_Number)
        #fill back to the original sheet
        df_whole['rowNum'][index]=pos
        df_whole['Instrument Category'][index]=Instrument_Category
        df_whole['Instrument'][index]=Instrument
        df_whole['Serial Number'][index]=Serial_Number
        #df_whole['Created Ticket'][index]=myKey
        #df_whole['Ticket Link'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        #print("Finished Create Ticket"+'\n')
        #go to the next row 
        pos+=1
        #after all the process, we drop the cloumn we don't want
    df_whole.drop(df_whole.iloc[:, 8::], inplace = True, axis = 1)
    try:
        print("Successfully login")
        initWindow.destroy()
    except:
        # no records need to generate tickets
        mb.showerror("Error", "Nothing to generate.")
        return
def autoGenerate():
    pos=0
    global df_whole
    global cb_intvar
    print(cb_intvar)
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    pd.set_option('mode.chained_assignment', None)
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
        print("row number is : ")
        print(pos)
        pos+=1
    df_whole.drop(df_whole.iloc[:, 8::], inplace = True, axis = 1)
    df_whole.to_excel("output_test222.xlsx", sheet_name='S1',index=False)
    print("Missing Done")
def on_resize(event):
    """Resize canvas scrollregion when the canvas is resized."""
    canvas.configure(scrollregion=canvas.bbox('all'))   
if __name__=='__main__':
    log_inWindow = tk.Tk()
    log_inWindow.title('JIRA Login')
    log_inWindow.geometry("390x135")
    log_inWindow['bg'] = "#0E69C1"
    
    tk.Label(log_inWindow, text = 'Username', font = 'bold', bg = "#0E69C1", fg = "white").grid(row = 0, column = 0, padx = 10, pady = 10)
    tk.Label(log_inWindow, text = 'Password', font = 'bold', bg = "#0E69C1", fg = "white").grid(row = 1, column = 0, padx = 10, pady = 10)

    e1 = tk.Entry(log_inWindow, font = "bold", bg = "white", fg = "black", cursor = "heart", insertbackground = 'black')
    e2 = tk.Entry(log_inWindow, show = '*', font = "bold", bg = "white", fg = "black", cursor = "heart", insertbackground = 'black')

    e1.grid(row = 0, column = 1)
    e2.grid(row = 1, column = 1)
    tk.Button(log_inWindow, text = 'Login', font = "bold", command = save_textvariable, bd = 4, width = 8, bg = "#0E69C1", fg = "white", activeforeground = "gray").grid(row = 2, column = 2, padx = 10,)
    log_inWindow.mainloop()
###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###
    initWindow = tk.Tk()
    initWindow.title('Open Excel File')
    initWindow.geometry("500x300")
    initWindow.resizable(False, False)
    initWindow['bg'] = "#0E69C1"
    labelframe1 = LabelFrame(initWindow, text="Choose Excel Workbook", bg = "#0E69C1", fg = "white")
    labelframe1.pack(fill= "both", expand="yes", padx = 10, pady = 10) 
    tk.Button(labelframe1, text = "BROWSE", command = openFile, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 25)
    tk.Button(labelframe1, text = "ENTER", command = processExcel, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 100)
    tk.Label(labelframe1,text=" Version "+versionControl,font="Helvetica 10 italic",bg="#0E69C1",fg="white").place(x=10,y=230)
    initWindow.mainloop()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    mainWindow = tk.Tk()
    mainWindow.title("Import Row")
    mainWindow.state('zoomed')
    mainWindow.configure(bg = "#0E69C1")

    mainWindow.columnconfigure(0, weight = 1)
    mainWindow.rowconfigure(0, weight = 1)

    canvas = tk.Canvas(mainWindow) 
    frame = tk.Frame(canvas)
    total_column = df_whole.shape[1]
    total_row = df_whole.shape[0]
    
#update_button = tk.Button(frame, text = "UPDATE", command = mytest, bd = 4, relief = RAISED, width = 15, height = 2)
    create_button = tk.Button(frame, text = "CREATE",  command =autoGenerate, bd = 4, relief = RAISED, width = 15, height = 2)
    create_button.grid(row = 0, column = 1, padx = 15, pady = 10)
    tk.Label(frame, text = "Check the row NOT to create ticket.", font = 'bold').grid(row = 0, column = 3, padx = 10, pady = 10)
    start = 1
    for col in df_whole.columns: 
        tk.Label(frame, text = col, font = 'bold').grid(row = 1, column = start, padx = 10, pady = 10)
        start += 1

    for j in range(0, total_row):  
        for i in range(0, total_column):
            dataCell = df_whole.iloc[j,i]
            tk.Label(frame, text = dataCell).grid(row = j+2, column = i+1, padx = 10, pady = 10)
    cb_intvar=[]
    for j in range(1, total_row+1):  
        cb_intvar.append(IntVar())
        chbx = tk.Checkbutton(frame, variable=cb_intvar[-1])
        chbx.grid(row = j+1, column = 0, sticky = 'w')

    canvas.create_window(0, 0, anchor = 'nw', window = frame)
    vbar = ttk.Scrollbar(mainWindow, orient = 'vertical', command = canvas.yview)
    hbar = ttk.Scrollbar(mainWindow, orient = 'horizontal', command = canvas.xview)
    canvas.configure(xscrollcommand = hbar.set,
                yscrollcommand = vbar.set,
                scrollregion = canvas.bbox('all'))

    canvas.grid(row = 0, column = 0, sticky = 'eswn')
    vbar.grid(row = 0, column = 1, sticky = 'ns')
    hbar.grid(row = 1, column = 0, sticky = 'ew')   
    canvas.bind('<Configure>', on_resize)
    mainWindow.mainloop()