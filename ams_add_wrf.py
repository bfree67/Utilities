# -*- coding: utf-8 -*-
"""
Created on Fri May  8 15:20:20 2020

@author: Brian
Loads station AMS file and WRF file
Merges both and saves
"""

import glob
import pandas as pd
import numpy as np
import datetime



def clean_name(text_list):
    #### Reformat column names
    new_list = []
    for text in text_list:
        new_text = text.replace('Conc,','')
        new_text = new_text.replace(',1','')
        new_text = new_text.replace(',2','')
        new_text = new_text.replace(' conc','')
        new_text = new_text.replace(' Conc','')
        new_text = new_text.replace('Wind Direction','WD')
        new_text = new_text.replace('Wind Speed','WS')
        new_text = new_text.replace('Temperature','TEMP')
        new_text = new_text.replace('Relative Humidity','RH')
        new_text = new_text.replace('Pressure','ATP')
        new_text = new_text.strip()
        new_text = new_text.replace('yr','year')
        new_text = new_text.replace('mo','month')
        new_text = new_text.replace('dy','day')
        new_text = new_text.replace('hr','hour')
        new_list.append(new_text)
    return new_list

def load_ams(file):
    #### load AMS file and strip headers from AMS data file
    df_ams = pd.read_excel(file)
    a_headers = clean_name(df_ams.iloc[3,:].to_list()) # extract header names and remove extra text
    df_ams.columns = a_headers  # replace header names
    df_ams.drop(df_ams.index[[0,1,2,3,4,5]], inplace=True) # delete first 6 rows
    return df_ams.reset_index(drop=True)    
    
#### prepare output Excel file
file_out = 'AMS02_Jahra_AMS_WRF_Merged.xls'
writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

#### load data files
file_list = glob.glob('*.xlsx')
for file in file_list:    
    print('\nReading file: ' + file)
    
    name_check = file[0:3]
    
    if name_check == 'AMS':
        
        df_ams = load_ams(file)  
        df_datetime = df_ams[['Date']].apply(pd.to_datetime)
    
        parameters = ['O3', 'SO2', 'NO2', 'CO', 'PM10', 'PM2.5', 'WD',
                   'WS', 'TEMP', 'RH']
        df_aqi = df_ams[parameters].apply(pd.to_numeric)
        df_aqi.CO = df_aqi.CO / 1000 # convert CO ug/m3 to mg/m3
        df_aqi = pd.concat((df_datetime, df_aqi), axis = 1) # combine date with parameters
        
        this_year = 2015
        df_ams_n = df_aqi[df_aqi['Date'].dt.year >= this_year] ## select only values in the year
        df_ams_n = df_ams_n.reset_index(drop=True) # reset index for each year
        
    if name_check == 'MET':
        df_wrf = pd.read_excel(file)
        df_wrf.columns = clean_name(df_wrf.columns.to_list())
        wrf_list = df_wrf.columns.to_list()
        df_wrf_dt = pd.to_datetime(df_wrf[['day', 'month', 'year', 'hour']]).to_frame()
        df_wrf_dt.columns = ['Date']
        df_wrf_n = pd.concat((df_wrf_dt, df_wrf[wrf_list[4:len(wrf_list)]]), axis = 1)

### merge AMS and WRF data together
df_merge = df_ams_n.merge(df_wrf_n, left_on='Date', right_on='Date', 
                           suffixes=('_left', '_right'))   
        
df_merge.to_excel(writer, sheet_name = file[0:15]) # save each AMS to separate worksheet
writer.save()  # save by year workbook
writer.close() # close by year workbook
print('\nSaving in file: ' + file_out)




    