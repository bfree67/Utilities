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

def add_stats(df):
    ### add mean, std dev, and median columns to dataframe
    mean = df.mean(axis = 1)
    std = df.std(axis = 1)
    median = df.median(axis = 1)
    df['Mean'] = mean
    df['STD'] = std
    df['Median'] = median
    return df

def save_pollutant(df, analyte):
    #df.columns = yr_col_strings  # change column names
    df = add_stats(df.replace(0,np.nan))
    df.to_excel(writer, sheet_name = analyte) # save each AMS to separate worksheet
    return
    
file_list = glob.glob('*.xlsx')

file_out = check_file('AMS02_Jahra_2015.xls')
    
#file_pollutants = check_file('AMS02_Jahra_by_pollutant.xls')

writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

for file in file_list:    
    print('\nReading file: ' + file)
    
    df = pd.read_excel(file)

    a_headers = clean_name(df.iloc[3,:].to_list()) # extract header names and remove extra text
    df.columns = a_headers  # replace header names
    df.drop(df.index[[0,1,2,3,4,5]], inplace=True) # delete first 6 rows
    df = df.reset_index(drop=True)
    
    df_datetime = df[['Date']].apply(pd.to_datetime)
    
    parameters = ['O3', 'SO2', 'NO2', 'CO', 'PM10', 'Wind Speed', 
                  'Wind Direction', 'Temperature', 'Relative Humidity', 'Pressure']
    
    df_aqi = df[parameters].apply(pd.to_numeric)
    
    df_aqi.CO = df_aqi.CO / 1000 # convert to mg/m3

    
    df_aqi = pd.concat((df_datetime, df_aqi), axis = 1) # combine date with parameters
    
    #yrs = df_aqi['Date'].dt.year.unique() # find yrs in data
    
    this_year = 2015
        
    df_year = df_aqi[df_aqi['Date'].dt.year == this_year] ## select only values in the year
    df_year = df_year.reset_index(drop=True) # reset index for each year

    worksheet_name = str(this_year)
    df_year.to_excel(writer, sheet_name = worksheet_name) # save each AMS to separate worksheet
        
    writer.save()  # save by year workbook
    writer.close() # close by year workbook




    