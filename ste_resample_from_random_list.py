# -*- coding: utf-8 -*-
"""
Created on Tue Dec  17 17:27:13 2019

Code to read table from SQL Server from a single sampler, downsample to 1 hr,
and put into a dataframe

Ver 3 - includes loop to do all 3 sources and save in separate worksheets
Reads file with list of sampler locations and loops through them

@author: Brian
"""
import pandas as pd
import pyodbc
import time

# file with list of sampler locations in Excel format
sampler_file = 'List_of_samplers_5.xls'
df_samplers = pd.read_excel(sampler_file)

samp_num = len(df_samplers)

x_long = df_samplers.Longitude.values
y_lat = df_samplers.Latitude.values

print('\nNumber of samplers = ' + str(samp_num))

## loop over list of sampler locations and save into different files
for j in range(samp_num):
    x = x_long[j]
    y = y_lat[j]
    
    print('\nSampler ' + str(j+1), str(x) + ', ' + str(y) + ':  ')

    # ******************* read sampler location from 3 sources ******************
    sampler_coord = str(x) + '_' + str(y)
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    writer = pd.ExcelWriter('s_'+ sampler_coord + '_' + time_stamp+'.xls')
    
    for i in range(3):
        source_num = i+1
        source = 'Source' + str(source_num)
        
        print('Loading ' + source + '...')
        
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
        df_ds.to_excel(writer, sheet_name = source)
        
    writer.save()