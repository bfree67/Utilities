# -*- coding: utf-8 -*-
"""
Created on Wed May 13 11:13:24 2020

@author: Brian
"""

import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler 
import math
import time

def time_prep(df):
    a = df.columns.to_list()
    np_df = df.values
    max_count = df.max()
    sector = 360/max_count   
    sector_rad = math.radians(sector)
    np_df_cos = np.round(np.cos(np_df*sector_rad),6) + 0
    np_df_sin = np.round(np.sin(np_df*sector_rad),6) + 0
    np_new = np.concatenate((np_df_cos, np_df_sin), axis = 1)
    df_new = pd.DataFrame(np_new)
    a_cos = a[0] + '_COS'; a_sin = a[0] + '_SIN'
    df_new.columns = [a_cos, a_sin]
    return df_new

def wind_prep(df):
    a = df.columns.to_list()
    np_df = np.radians(df.values)
    np_df_cos = np.round(np.cos(np_df),6) + 0
    np_df_sin = np.round(np.sin(np_df),6) + 0
    np_new = np.concatenate((np_df_cos, np_df_sin), axis = 1)
    df_new = pd.DataFrame(np_new)
    a_cos = a[0] + '_COS'; a_sin = a[0] + '_SIN'
    df_new.columns = [a_cos, a_sin]
    return df_new

# convert an array of values into a dataset tensor
# ASSUMES data is already in look_back packages
def TensorForm(data, look_back):
    #determine number of data samples
    rows_data,cols_data = np.shape(data)
    
    #determine # of batches based on look-back size
    tot_batches = int(rows_data/look_back)
    
    #initialize 3D tensor
    threeD = np.zeros(((tot_batches,look_back,cols_data)))
    
    # populate 3D tensor
    for i in range(tot_batches):  
        sample_num = i * look_back # skip by # of look_back
        
        for look_num in range(look_back):
            threeD[i,:,:] = data[sample_num:sample_num+(look_back),:]
    
    return threeD

print('\n****************************************************\n')
print('      Prep merged AMS & WRF data for imputing')
print('                   ver 13May2020_2_1')
print('\n****************************************************\n')
    
#### load raw merged AMS & WRF file
file_in = 'AMS02_Jahra_AMS_WRF_Merged.xls'
print('\nReading file: ' + file_in)
df_raw = pd.read_excel(file_in)

### drop rain column if present and remove rows with NaNs
if 'rain' in df_raw.columns:
    df_raw = df_raw.drop('rain', axis = 1)  # remove column 'rain' if present
#df_nn = df_raw.dropna()  ### remove any row with an NaN
df_nn = df_raw.reset_index(drop=True)   ### reset index

# isolate Date columns
df_Date = df_nn.Date.to_frame()

### convert cyclic data to sine and cosine components
df_hours = time_prep(df_nn.hour.to_frame())
df_month = time_prep(df_nn.month.to_frame())
df_wd = wind_prep(df_nn.wdir.to_frame())
df_cyclic = pd.concat((df_Date, df_month, df_hours, df_wd), axis = 1)

### scale continuous values using MinMax
df_scale = df_nn.drop(df_nn[['Date', 'month', 'day', 'hour', 'wdir']], axis = 1)
df_scale_col = df_scale.columns
np_scale = df_scale.values

#### save Max and Min values to setup Scaler later
scale_max = np.asmatrix(np_scale.max(axis=0))
scale_min = np.asmatrix(np_scale.min(axis=0))
scale_max_min = np.concatenate((scale_max, scale_min), axis = 0)
df_maxmin = pd.DataFrame(scale_max_min)
df_maxmin.columns = df_scale_col

### scale values
scalerX = MinMaxScaler(feature_range=(0, 1))
dataset = scalerX.fit_transform(np_scale)

df_scaled = pd.DataFrame(dataset)
df_scaled.columns = df_scale_col

### re-assemble processed data
df_data_raw = pd.concat((df_cyclic, df_scaled), axis = 1)

df_data = df_data_raw.dropna()  ### remove any row with an NaN
df_data = df_data.reset_index(drop=True)   ### reset index

# make timestamp for unique filname
stamp = str(time.perf_counter())  #add timestamp for unique name
stamp = stamp[0:4]

#### prepare output Excel file to save processed data
file_out = 'Jahra_processed_raw_data_' + stamp +'.xls'
writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

df_data_raw.to_excel(writer, index = False) # save to Excel
writer.save()  # save workbook
writer.close() # close workbook
print('\nSaving processed raw data in file: ' + file_out)

############## prepare tensor based training sets based on look-ahead
df_train = pd.DataFrame()
df_train_out = pd.DataFrame()
look_ahead = 1 # hr look-ahead
look_back = look_ahead + 2

print('\nPreparing training data tensor for look ahead = ' + str(look_ahead))

for i in range(look_back, len(df_data)-look_back):
    
    at = df_data.Date[i] - df_data.Date[i-look_back]
    at_hr = int(at.total_seconds()/3600)
    
    # find blocks of continuous hour rows that meet the look_back length
    # collect the rows and make a new df that is a multiple of look_back
    if at_hr == look_back:
        df_set = df_data.iloc[i-look_back:i]
        df_train = pd.concat((df_train, df_set), axis = 0)
        
        df_out = df_data.iloc[i+look_ahead].to_frame().T
        df_out = df_out[['O3', 'SO2', 'NO2', 'CO', 'PM10']]
        
        df_train_out = pd.concat((df_train_out, df_out), axis = 0)

df_train_noDate = df_train.drop(df_raw.columns[0], axis=1) # remove Date columns

#X_tensor = TensorForm(df_train_noDate.values, look_back)

# make timestamp for unique filname
stamp = str(time.perf_counter())  #add timestamp for unique name
stamp = stamp[0:4]

#### prepare output Excel file to save processed data
file_out = 'Jahra_Impute_TensorReady_ahead_' + str(look_ahead) + '_' + stamp +'.xls'
writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

df_train_noDate.to_excel(writer, sheet_name = 'X_Data', index = False) # save to Excel
df_train_out.to_excel(writer, sheet_name = 'Y_Data', index = False)
df_maxmin.to_excel(writer, sheet_name = 'MinMax', index = False)
writer.save()  # save workbook
writer.close() # close workbook
print('\nSaving in file: ' + file_out)



        
    



