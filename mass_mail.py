# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 10:00:19 2020

@author: Brian
"""

import win32com.client
import glob
import os
import pandas as pd

df_list = pd.read_csv('2017facilities.csv')

# choose sender account
outlook = win32com.client.Dispatch('outlook.application')
send_account = None
for account in outlook.Session.Accounts:
    if account.DisplayName == 'ocs_aqs_team@weblakes.com':
        send_account = account
        break

# get absolute file path location
path = os.path.abspath(os.getcwd())

# make list of file attachments from dataframe
file_list = df_list['File'].tolist()

# for testing!
#recips = ['brian.freeman@weblakes.com', 'mohammad.munshed@weblakes.com']

# make list of email recipients from dataframe when ready to go!
recips = df_list['Email'].tolist()

for i in range (len(recips)):
    
    # select individual email recipient
    address = recips[i]
    attachment = path + '\\' + file_list[i]

    mail_item = outlook.CreateItem(0)   # 0: olMailItem
    
    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    
    mail_item.To = address
    mail_item.BCC = 'dave.lim@xatorcorp.com; brian.freeman@weblakes.com'
    mail_item.Subject =  'DO NOT REPLY: ' + df_list['2017_CO_NAME'][i] + ' Emissions Inventory facilities review'
    mail_item.BodyFormat = 2   # 2: Html format
    mail_item.Body = '''Hi,
    
This email has been generated for: 
    \n'''  + df_list['2017_CO_NAME'][i] + '''
BOEM Company Number: ''' + str(df_list['BOEM_CO_NUM'][i]) + ''' 

Your company is on record as having submitted a 2017 inventory and an account has been set up for you in OCS AQS for the 2021 inventory submission, as previously announced.
Your OCS AQS account will be set up with the facilities shown in the attached spreadsheet. Please review it carefully and let us know if the list is correct. 
If the list is not up-to-date due to transfers of ownership, decommissioning, or new facilities, please update the attachment accordingly and send it back to OCS.AQS_support@weblakes.com.  Once we receive the revised information, we will update your account and send you an email inviting you to log in to OCS AQS.

Thank you,

OCS AQS Support Team 
OCS.AQS_support@weblakes.com'''

    
    mail_item.Attachments.Add(attachment)
    mail_item.Send()