# -*- coding: utf-8 -*-
"""
Created on Tue Dec  3 16:27:13 2019

Code to read a table in a MS Access DB

@author: Brian
"""
import pandas as pd
import pyodbc
import pandas.io.sql as psql

'''
#2017 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2017/Emission Inventory/archive/'
MDB = '2017_Gulfwide_Platform_Inventory_20190705_CAP_GHG.accdb'
sql = 'SELECT [2017_Gulfwide_Platform_20190705_CAP_GHG].* FROM 2017_Gulfwide_Platform_20190705_CAP_GHG;'
'''
#######
'''
#2014 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2014/'
MDB = '2014_Gulfwide_Platform_Inventory_20161102.accdb'
sql = 'SELECT [2014_Gulfwide_Platform_20161102].* FROM 2014_Gulfwide_Platform_20161102;'
'''
#######
#2011 Platform
PATH = 'C:/Users/Brian/Documents/_Active/BOEM/Historical Data/GOADS_Historic_EI/2011/'
MDB = '2011_Gulfwide_Platform_Inventory.accdb'
sql = 'SELECT tblPointCE.* FROM tblPointCE;'


# declare driver to connect to DB
DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'

connStr = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=" + PATH + MDB + ";"
    )
cnxn = pyodbc.connect(connStr)

dataf = pd.read_sql(sql, cnxn)
cnxn.close()

data_type = dataf.dtypes
data_type.to_clipboard()