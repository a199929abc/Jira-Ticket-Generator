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
def extractRow(): 

    
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
###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~###
    initWindow = tk.Tk()
    initWindow.title('Open Excel File')
    initWindow.geometry("500x300")
    initWindow.resizable(False, False)
    initWindow['bg'] = "#0E69C1"
    labelframe1 = LabelFrame(initWindow, text="Choose Excel Workbook", bg = "#0E69C1", fg = "white")
    labelframe1.pack(fill= "both", expand="yes", padx = 10, pady = 10) 
    tk.Button(labelframe1, text = "BROWSE", command = openFile, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 25)
    tk.Button(labelframe1, text = "ENTER", command = extractRow, font = 'bold', bd = 4, width = 10, bg = "white", fg = "black", activeforeground = "gray").place(x = 350, y = 100)
    initWindow.mainloop()


    

    
