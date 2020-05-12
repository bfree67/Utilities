# -*- coding: utf-8 -*-
"""
Created on Fri May  8 15:20:20 2020

@author: Brian
Loads indiviudal excel files of AMS data and prepares headers
saves correlation matrix
"""

import glob
import pandas as pd
import numpy as np


def clean_name(text_list):
    new_list = []
    for text in text_list:
        new_text = text.replace('Conc,','')
        new_text = new_text.replace(',1','')
        new_text = new_text.replace(',2','')
        new_text = new_text.replace(' conc','')
        new_text = new_text.replace(' Conc','')
        new_list.append(new_text)
    return new_list

file_list = glob.glob('*.xlsx')

#### prepare output Excel file
file_out = 'Kuwait_AMS_corr.xls'
writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')


for file in file_list:    
    print('\nReading file: ' + file)
        
    df = pd.read_excel(file)

    a_headers = clean_name(df.iloc[3,:].to_list()) # extract header names and remove extra text
    df.columns = a_headers  # replace header names
    df.drop(df.index[[0,1,2,3,4,5]], inplace=True) # delete first 6 rows
    df = df.reset_index(drop=True)
    
    df_datetime = df[['Date']].apply(pd.to_datetime)
    #df_datetime['Date']=df_datetime['Date'].dt.date
    #df_datetime['Time']=df_datetime['Time'].dt.time
    
    parameters = a_headers[2:len(a_headers)] #['O3', 'SO2', 'NO2', 'CO', 'PM10', 'PM2.5', 'Wind Direction',
               #'Wind Speed', 'Temperature', 'Relative Humidity', 'Pressure']
    df_aqi = df[parameters].apply(pd.to_numeric)
    df_aqi.CO = df_aqi.CO / 1000 # convert to mg/m3
    
    df_corr = df_aqi.corr()
        
    df_corr.to_excel(writer, sheet_name = file[0:15]) # save each AMS to separate worksheet

            
writer.save()  # save by year workbook
writer.close() # close by year workbook
print('\nSaving in file: ' + file_out)
   




    