"""
Author:Kaiheng Zhang
Mail : kaiheng365@gmail.com
Date: 2021-01-13
Version:VersionControl
"""

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
import sys
import globalvar as gl

#Variables
username = ''
password = ''
mypath = ''
workbookTitle = ''
loginWindow = ''
GUI_flag= False
versionControl= '2.0.4'
df_whole= pd.DataFrame()
# create a dataframe for storing the deploy row
df_out = pd.DataFrame()
# create a dictionary for storing the site ticket EN number
site_dict = {}

    
def save_textvariable():
    global username
    global password
    username = e1.get()
    password = e2.get()
    gl._init()
    gl.set_value('username',username)
    gl.set_value('password',password)
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
def processExcel():
    global mypath
    pos = 0
    index=0
    global df_whole
    df_whole = pd.read_excel(mypath)
    df_whole.drop(df_whole.iloc[:, 10::], inplace = True, axis = 1)
    df_whole.insert(10, "rowNum",np.nan)
    df_whole.columns=['DeviceID','Due Date','Assignee','Description','Ticket Link','Instrument Category',
    'Instrument','Serial Number','Created Ticket','status','rowNum']
    df_whole.insert(11, "Component", "Test and Development")
    df_whole.insert(12, "Linked To", np.nan)
    df_whole.insert(13,"Work Ticket",np.nan)
    df_whole = df_whole[:-1]
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
    try:
        print("Successfully login")
        initWindow.destroy()
    except:
        # no records need to generate tickets
        mb.showerror("Error", "Nothing to generate.")
        return 0
def autoGenerate():
    pos=0
    drop_row = []
    for ctr, int_var in enumerate(cb_intvar):
        if int_var.get():
            drop_row.append(ctr)
    df_out=df_whole.copy()
    print(df_out)
    df_out= df_out.drop(index=drop_row)
                
    for index, row in df_out.iterrows():
        local_instrument_category=''
        local_instrument=''
        serial_number= ''
        status=''
        local_instrument,local_instrument_category=onc_request(row)
        local_instrument, serial_number = processString(local_instrument)  
        df_out['rowNum'][index]=pos
        df_out['Instrument Category'][index]=local_instrument_category
        df_out['Instrument'][index]=local_instrument
        df_out['Serial Number'][index]=serial_number
        pos+=1
        myKey = create_ticket(row,local_instrument_category, local_instrument, serial_number)
        df_out['Created Ticket'][index]=myKey
        status=check_status(myKey)
        df_out['status'][index]=status
        df_out['Created Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        print("http://142.104.193.65:8080/browse/%s" % myKey)
    df_out.drop(df_whole.iloc[:, 10::], inplace = True, axis = 1)
    df_out.to_excel("output_file.xlsx", sheet_name='S1',index=False) 
def on_resize(event):

    """Resize canvas scrollregion when the canvas is resized."""
    canvas.configure(scrollregion=canvas.bbox('all'))   
def update_status():
    global mypath
    df_check=pd.DataFrame()
    df_check = pd.read_excel(mypath)
    print(df_check)
    for index, row in df_check.iterrows():
        status=''
        link=''
        link=str(row['Created Ticket'])
        myKey =link.split('/')[-1]
        status=check_status(myKey)
        df_check['status'][index]=status
    df_check.to_excel(mypath, index = False)
    return 0
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###
    initWindow = tk.Tk()
    initWindow.title('Open Excel File')
    initWindow.geometry("500x300")
    initWindow.resizable(False, False)
    initWindow['bg'] = "#0E69C1"
    labelframe1 = LabelFrame(initWindow, text="Choose Excel Workbook", bg = "#0E69C1", fg = "white")
    labelframe1.pack(fill= "both", expand="yes", padx = 10, pady = 10) 
    #transfer_variable
    tk.Button(labelframe1, text = "Brose", command = openFile, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 25)
    tk.Button(labelframe1, text = "Enter", command = processExcel, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 100)
    tk.Button(labelframe1, text = "Update", command =update_status, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 175)
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