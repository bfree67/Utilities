# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 10:00:19 2020

@author: Brian
"""

import win32com.client
import glob
import os
import pandas as pd

df_list = pd.read_csv('no_facilities.csv')

# choose sender account
outlook = win32com.client.Dispatch('outlook.application')
send_account = None
for account in outlook.Session.Accounts:
    if account.DisplayName == 'ocs_aqs_team@weblakes.com':
        send_account = account
        break

#path = os.path.abspath(os.getcwd())
#file_list = glob.glob('*.xlsx')

#recips = ['brian.freeman@weblakes.com']#, 'mohammad.munshed@weblakes.com']
recips = df_list['Email'].tolist()

for i in range (len(recips)):
    
    address = recips[i]
    #attachment = path + '\\' + file_list[i]

    mail_item = outlook.CreateItem(0)   # 0: olMailItem
    
    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    
    mail_item.To = address
    mail_item.BCC = 'dave.lim@xator.com'
    mail_item.Subject =  df_list['2020_CO_NAME'][i] + ' missing OCS 2021 Emissions Inventory facilities'
    mail_item.BodyFormat = 2   # 2: Html format
    mail_item.Body = '''Hi,
    
This email has been generated for: 
    \n'''  + df_list['2020_CO_NAME'][i] + '''
BOEM Company Number: ''' + str(df_list['BOEM_CO_NUM'][i]) + ''' 

Your company was listed in the BOEM Active Company data list as a lease owner in the Gulf of Mexico for 2020. We assume that you will be taking part in the 2021 OCS Emissions Inventory but we currently do not have any facilities listed for your company. 
Please respond to ocs.aqs_support@weblakes.com with the requested information shown below so we can update your account accordingly: 
•	If you have facilities that were in the 2017 OCS Emissions Inventory, please email us with the facility platform ID number and lease number.
•	If you have new facilities and platforms that you would like us to include in the 2021 OCS Emissions inventory, please let us know and we will provide you a template that will include necessary information.
•	If you have no facilities or will not be participating in the 2021 OCS Emissions Inventory, please let us know so that we can remove you from our records.

Thank you,

OCS AQS Support Team '''

    
    #mail_item.Attachments.Add(attachment)
    mail_item.Send()