import os
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import openpyxl
import pandas as pd
import numpy as np
#from onc.onc import ONC
import json
import requests

from request import *
# JIRA libraries
from jira.resources import IssueLink
import jira.client
from jira.client import JIRA
from jira import JIRA
import re
import datetime
import globalvar as gl


def create_ticket(row,instrument_category,instrument,serial_number):
    # Assign the values
    nat = np.datetime64('NaT')
    __SerialNumber = serial_number
    __instrument=   instrument
    __instrumentcategory =instrument_category
    __DeviceID =      row['DeviceID']
    __createdTicket=row['Created Ticket']
    __components=   row['Component'] 
    __outwardIssue = row['Ticket Link']
    __assignee= row['Assignee']
    if(isinstance(__assignee,str)):
        __assignee =__assignee[0:__assignee.rfind('@')]
    else: 
        __assignee= np.nan

    __duedate = row['Due Date']
    __description =row['Description']
    __title='Instrument Qualification '
        # Connect to jira
    # Authentication done by using username and password
    #global username,password
    #username,password=transfer_variable()
    username =gl.get_value('username')
    password=gl.get_value('password')
    jira = JIRA(
        basic_auth = (username, password),
        options = {'server': 'http://142.104.193.65:8080'}
        #options = {'server': 'https://jira.oceannetworks.ca/'}
    )
    #print(type(__duedate))
    #print(type(__assignee))
    #create a ticket for current row

    if((isinstance(__assignee, str )==True) and (isinstance(__duedate, pd._libs.tslibs.nattype.NaTType)==True)):
        new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary ="{0}: {1}, SN: {2} (DI:{3})".format(__title,__instrument,__SerialNumber, int(__DeviceID)),
        # "%s: %s SI: %s DI: %s" % ( __instrument, __SerialNumber, __DeviceID),
        description = " instrument category: %s\n note: %s\n " % ( __instrumentcategory,__description), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID,
        #duedate=str(__duedate),
        assignee={'name': 	__assignee}
        #assignee format take only name before the email
        #assignee={'sfaassfda'}
        #https://innovalog.atlassian.net/wiki/spaces/JMWEC/pages/108200050/Standard+JIRA+fields Very helpful link 
        )
        if(isinstance(row['Ticket Link'], str)):
            jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)
            
        return new_issue.key

    # due date NAT Assignee NaN
    elif((isinstance(__assignee, str )==False) and (isinstance(__duedate, pd._libs.tslibs.nattype.NaTType)==True)):

        new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary ="{0}: {1}, SN: {2} (DI:{3})".format(__title,__instrument,__SerialNumber, int(__DeviceID)),
        # "%s: %s SI: %s DI: %s" % ( __instrument, __SerialNumber, __DeviceID),
        description = " instrument category: %s\n note: %s\n " % ( __instrumentcategory,__description), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID,
        #duedate=str(__duedate),
        #assignee={'name': 	__assignee}
        #assignee format take only name before the email
        #assignee={'sfaassfda'}
        #https://innovalog.atlassian.net/wiki/spaces/JMWEC/pages/108200050/Standard+JIRA+fields Very helpful link 
        )
        if(isinstance(row['Ticket Link'], str)):
            jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)
            
        return new_issue.key

    elif((isinstance(__assignee, str )==True) and (isinstance(__duedate, pd._libs.tslibs.nattype.NaTType)==False)):
        new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary ="{0}: {1}, SN: {2} (DI:{3})".format(__title,__instrument,__SerialNumber, int(__DeviceID)),
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
        #https://innovalog.atlassian.net/wiki/spaces/JMWEC/pages/108200050/Standard+JIRA+fields Very helpful link 
        )
        if(isinstance(row['Ticket Link'], str)):
            jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)
            
        return new_issue.key

    else:

        new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary ="{0}: {1}, SN: {2} (DI:{3})".format(__title,__instrument,__SerialNumber, int(__DeviceID)),
        # "%s: %s SI: %s DI: %s" % ( __instrument, __SerialNumber, __DeviceID),
        description = " instrument category: %s\n note: %s\n " % ( __instrumentcategory,__description), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID,
        duedate=str(__duedate),
        )
    # add the Ticket Link
    if(isinstance(row['Ticket Link'], str)):
        jira.create_issue_link("Related", new_issue.key, __outwardIssue, None) 
    print(new_issue.key)

    return new_issue.key
    


def check_status(ticket):
    #global username,password
    #username,password=transfer_variable()
    username =gl.get_value('username')
    password=gl.get_value('password')
    jira = JIRA(    
    basic_auth = (username, password),
    options = {'server': 'http://142.104.193.65:8080'}
    #options = {'server': 'https://jira.oceannetworks.ca/'}
                )
    new_issue = jira.issue(ticket)
    return new_issue.fields.status.name
