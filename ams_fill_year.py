# -*- coding: utf-8 -*-
"""
Created on Fri May  8 15:20:20 2020

@author: Brian
Loads indiviudal excel files of AMS data and prepares headers
CO2, CH4, CO in mg/u3
Converts columns into numbers or datetime
Saves each year into separate worksheets
Then save each pollutant by year

"""

import glob
import pandas as pd
import os.path

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

def no_leap(df):
    ### removes hours in a leap year so everybody lines up
    df_1 = (df['Date'].dt.month != 2).to_frame() + 0
    df_2 = (df['Date'].dt.day != 29).to_frame() + 0
    df_mask = ((df_1 + df_2) > 0).to_numpy().flatten()
    return pd.Series(df_mask)

def check_file(file_out):
    ### see if file exists and if it does, add a prefix
    if os.path.isfile(file_out):
        file_1 = file_out
        file_out = '1_' + file_out
        print(file_1 + ' already exists. Saving as ' + file_out)
    return file_out

def save_pollutant(df, analyte):
        df.columns = yr_col_strings  # change column names
        df = df.replace(0,np.nan)
        df.to_excel(writer, sheet_name = analyte) # save each AMS to separate worksheet
    
file_list = glob.glob('*.xlsx')

file_out = check_file('AMS02_Jahra_by_year.xls')
    
file_pollutants = check_file('AMS02_Jahra_by_pollutant.xls')

writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

for file in file_list:    
    print('\nReading file: ' + file)
    
    df = pd.read_excel(file)

    a_headers = clean_name(df.iloc[3,:].to_list()) # extract header names and remove extra text
    df.columns = a_headers  # replace header names
    df.drop(df.index[[0,1,2,3,4,5]], inplace=True) # delete first 6 rows
    df = df.reset_index(drop=True)
    
    df_datetime = df[['Date']].apply(pd.to_datetime)
    df_o3 = pd.DataFrame()
    df_so2 = pd.DataFrame()
    df_no2 = pd.DataFrame()
    df_co = pd.DataFrame()
    df_pm10 = pd.DataFrame()
    
    parameters = ['O3', 'SO2', 'NO2', 'CO', 'PM10']
    
    df_aqi = df[parameters].apply(pd.to_numeric)
    
    df_aqi.CO = df_aqi.CO / 1000 # convert to mg/m3
    
    df_aqi = pd.concat((df_datetime, df_aqi), axis = 1) # combine date with parameters
    
    yrs = df_aqi['Date'].dt.year.unique() # find yrs in data
    
    for this_year in yrs:
        
        df_year = df_aqi[df_aqi['Date'].dt.year == this_year] ## select onlyu values in the year
        df_year = df_year.reset_index(drop=True) # reset index for each year
    
        worksheet_name = str(this_year)
        df_year.to_excel(writer, sheet_name = worksheet_name) # save each AMS to separate worksheet
        
        mask = no_leap(df_year)   #make mask to remove leap year hours (if any)
        df_noleap = df_year[mask] #remove leap year hours
        
        ##### add column to analytes
        df_o3 = pd.concat((df_o3, df_noleap.O3), axis = 1)
        df_so2 = pd.concat((df_so2, df_noleap.SO2), axis = 1)
        df_no2 = pd.concat((df_no2, df_noleap.NO2), axis = 1)
        df_co = pd.concat((df_co, df_noleap.CO), axis = 1)
        df_pm10 = pd.concat((df_pm10, df_noleap.PM10), axis = 1)
    
    writer.save()  # save by year workbook
    writer.close() # close by year workbook

#### save analytes in separate sheets

writer = pd.ExcelWriter(file_pollutants, engine = 'xlsxwriter')

yr_cols = df_aqi['Date'].dt.year.unique().tolist() # get list of years in data
yr_col_strings = ["%.0f" % x for x in yr_cols]  # convert number list to strings

for analyte in parameters:

    if analyte == 'O3':
        save_pollutant(df_o3, analyte)
    if analyte == 'SO2':
        save_pollutant(df_so2, analyte)
    if analyte == 'NO2':
        save_pollutant(df_no2, analyte)
    if analyte == 'CO':
        save_pollutant(df_co, analyte)
    if analyte == 'PM10':
        save_pollutant(df_pm10, analyte)
        
writer.save()  # save by pollutant workbook
writer.close() # close by pollutant workbook


    