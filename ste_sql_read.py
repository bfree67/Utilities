# -*- coding: utf-8 -*-
"""
Created on Tue Dec  12 17:27:13 2019

Code to read table from SQL Server

@author: Brian
"""
import pandas as pd
import pyodbc

def resample_df(column):
#### downsample time series to 1 hr averaged data into average and std dev.
    df1 = dataf.time.to_frame()
    df1 = df1.iloc[0:len(dataf),:]
    df1[column]=dataf[column]
    df1.set_index('time', inplace=True)
    df1_ave = df1.resample('1H', label='right', closed='right').mean()
    df1_std = df1.resample('1H', label='right', closed='right').std()
    return df1_ave, df1_std

x = -104.9178
y = 41.91665
source = 1
 
folder = 'Source' + str(source)
 
# prepare sql instruciton
sql = "SELECT * "\
      "FROM [CsvImport] "\
      "where [x] = " + str(x) + " and [y] = " + str(y) + " and [folderName] = 'Source1'" \
      " ORDER BY [time] ASC;"


# make DB connection
cnxn = pyodbc.connect('Driver={SQL Server};'
                      'Server=hvdev-sql-01;'
                      'Database=brian_csv_import;'
                      'UID=sa;'
                      'PWD=Riley++;')
# read into dataframe
dataf = pd.read_sql(sql, cnxn)

# close connection
cnxn.close()

#### downsample time series to 1 hr averaged data
dep_ave, dep_std = resample_df('DepTracer_kbm2')
conc_ave, conc_std = resample_df('ConcTracer_kbm3')

df_new = dep_ave

df_new['Source'] = source
df_new['x'] = x
df_new['y'] = y
