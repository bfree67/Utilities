# -*- coding: utf-8 -*-
"""
Created on Tue Dec  17 17:27:13 2019

Code to read table from SQL Server from a single sampler, downsample to 1 hr,
and put into a dataframe

Ver 2 - includes loop to do all 3 sources and save in separate worksheets

@author: Brian
"""
import pandas as pd
import pyodbc
import time

source_num = 1
x = -104.8815002
y = 42.13731003

for i in range(3):
    source_num = i+1
    source = 'Source' + str(source_num)
    
    print('\nLoading ' + source + '...')
    
    # prepare sql instruction
    sql1 = "SELECT * "\
          "FROM [CsvImport] "\
          "where [x] = '" + str(x) + "' and [y] = '" + str(y) + "' and [folderName] = '" + source + "'"\
          " ORDER BY [time] ASC;"
    
    # make DB connection
    cnxn = pyodbc.connect('Driver={SQL Server};'
                          'Server=hvdev-sql-01;'
                          'Database=brian_csv_import;'
                          'UID=sa;'
                          'PWD=Riley++;')
    # read into dataframe
    sql = sql1
    dataf = pd.read_sql(sql, cnxn)
    
    # close connection
    cnxn.close()
    
    # select time column for index
    df1 = dataf.time.to_frame()
    df2 = dataf.time.to_frame()
    
    df1 = df1.iloc[0:len(dataf),:]
    df2 = df2.iloc[0:len(dataf),:]
    
    df1['conc']=dataf.ConcTracer_kbm3
    df2['dep']=dataf.DepTracer_kbm2
    
    df1.set_index('time', inplace=True)
    df2.set_index('time', inplace=True)
    
    # downsample
    df1_ds = df1.resample('1H', label='right', closed='right').mean()
    df2_ds = df2.resample('1H', label='right', closed='right').mean()
    
    # add extra columns
    df1_ds['dep'] = df2_ds.dep
    df1_ds['Latitude'] = y
    df1_ds['Longitude'] = x
    df1_ds['Source'] = source_num
    
    df_ds = df1_ds[['Source', 'Latitude', 'Longitude', 'conc', 'dep']]
    
    # save random list 
    sampler_coord = str(x) + '_' + str(y)
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    writer = pd.ExcelWriter('s_'+ sampler_coord + time_stamp+'.xls')
    df_ds.to_excel(writer, sheet_name = source)
    writer.save()

