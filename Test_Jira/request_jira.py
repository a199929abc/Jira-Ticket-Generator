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
import datetime

def create_ticket(row,instrument_category,instrument,serial_number):
    # Assign the values
    __SerialNumber = serial_number
    __instrument=   instrument
    __instrumentcategory =instrument_category
    __DeviceID =      row['DeviceID']
  # __SiteLocation = row['Site/Location']
    __TicketLink=   row['Ticket Link']
    __createdTicket=row['Created Ticket']
    __components=   row['Component'] 
    __outwardIssue = row['Linked To']
    __assignee= row['Assignee']
    __assignee =__assignee[0:__assignee.rfind('@')]
    __duedate = row['Due Date']
    __description =row['Description']
    print(__description)
    
    
    #__issueLink = row['Work Ticket']
    ''' if(row['Operation'] == 'Deploy'):
     __summaryTitle = 'instrument Qualification'
    if(row['Operation'] == 'Recover'):
        __summaryTitle = 'instrument Recovery'''

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
        description = " instrument category: %s\n note: %s\n " % ( __instrumentcategory,__description), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID,
        duedate=str(__duedate),
        assignee={'name': 	__assignee}
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