from jira.resources import IssueLink
import jira.client
from jira.client import JIRA
from jira import JIRA
import re
import time 


def create_ticket():
    # Assign the values
    __SerialNumber = serial_number
    __instrument=   instrument
    __instrumentcategory =instrument_category
    __DeviceID =    31379
    #__TicketLink=   'EN-32123'
   # __createdTicket=row['Created Ticket']
    __components=   'test and devlopment'
    __outwardIssue = None
    __issueLink = None


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
        description = "Site Location: %s\n  instrument category: %s\n " % (__SiteLocation,  __instrumentcategory), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __SerialNumber,      # Serial # field
        customfield_10070 = __DeviceID)                 # Device ID field

    # add the linkedto
    if(isinstance(row['Linked To'], str)):
        jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)

    return new_issue.key