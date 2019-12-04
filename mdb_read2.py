# -*- coding: utf-8 -*-
"""
Created on Tue Dec  3 16:27:13 2019

Code to read a table in a MS Access DB

@author: Brian
"""
import pandas as pd
import pyodbc
import pandas.io.sql as psql

def call_db(PATH, MDB, sql):
    
    connStr = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=" + PATH + MDB + ";"
        )
    cnxn = pyodbc.connect(connStr)
    
    dataf = pd.read_sql(sql, cnxn)
    cnxn.close()
    
    return dataf.dtypes

'''
#2017 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2017/Emission Inventory/archive/'
MDB = '2017_Gulfwide_Platform_Inventory_20190705_CAP_GHG.accdb'
table = ['2017_Gulfwide_Platform_20190705_CAP_GHG']
'''
#######
'''
#2014 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2014/'
MDB = '2014_Gulfwide_Platform_Inventory_20161102.accdb'
table = ['2014_Gulfwide_Platform_20161102']
'''
#######
'''
#2011 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2011/'
MDB = '2011_Gulfwide_Platform_Inventory.accdb'
table = ['tblPointCE', 'tblPointEM', 'tblPointEP', 'tblPointER', 'tblPointEU', 'tblPointPE', 'tblPointSI', 'tblPointTR']
'''
#######
'''
#2008 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2008/'
MDB = '2008-Gulfwide-NIF platform inventory.mdb'
table = ['tblPointCE', 'tblPointEM', 'tblPointEP', 'tblPointER', 'tblPointEU', 'tblPointPE', 'tblPointSI', 'tblPointTR']
'''
#######
'''
#2005 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2005/'
MDB = '2005-GOADS-NIF-v1.mdb'
table = ['tblPointCE', 'tblPointEM', 'tblPointEP', 'tblPointER', 'tblPointEU', 'tblPointPE', 'tblPointSI', 'tblPointTR']
'''
#######
#2000 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2000/'
MDB = 'GOADS-NIF-SEP04.mdb'
table = ['tblPointCE', 'tblPointEP', 'tblPointER', 'tblPointEU', 'tblPointPE', 'tblPointEM']

########################################################
## start
data_type = pd.DataFrame()

for tb in table:
    
    sql = 'SELECT TOP 1 ' + tb + '.* FROM ' + tb +';'
    
    d1 = call_db(PATH, MDB, sql)

    data_type = pd.concat([data_type, d1])