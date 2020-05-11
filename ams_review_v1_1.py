# -*- coding: utf-8 -*-
"""
Created on Fri May  8 15:20:20 2020

@author: Brian
Loads indiviudal excel files of AMS data and prepares headers
CO2, CH4, CO in mg/u3
Converts columns into numbers or datetime
Counts occupied cells by year.
Saves as individual worksheet
"""



import glob
import pandas as pd

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

file_out = 'AMS_count.xls'
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
    df_aqi.CH4 = df_aqi.CH4 / 1000 # convert to mg/m3
    
    df_aqi = pd.concat((df_datetime, df_aqi), axis = 1) # combine date with parameters
    
    count = df_aqi.groupby(df_aqi['Date'].dt.year).agg(['count']) # count occurrences by year
    
    worksheet_name = file[0:15]
    
    count.to_excel(writer, sheet_name = worksheet_name) # save each AMS to separate worksheet
    
writer.save()
    
writer.close()


    