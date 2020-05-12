# -*- coding: utf-8 -*-
"""
Created on Tue May 12 16:23:56 2020
Simple imputing for 1 empty space per row up to 8 consecutive gaps
Imputes using median strategy from all rows if available

@author: Brian
"""

import pandas as pd
from sklearn.impute import SimpleImputer
import numpy as np

#### prepare output Excel file
file_out = 'imputed_file.xls'
writer = pd.ExcelWriter(file_out, engine = 'xlsxwriter')

file_in = 'AMS02_Jahra_by_pollutant.xlsx'
print('\nReading file: ' + file_in)
xls = pd.ExcelFile(file_in)

work_sheets = xls.sheet_names

imp = SimpleImputer(missing_values = np.nan, strategy = 'median')

for sheet in work_sheets:
    print('\nImputing ' + sheet)
    df = pd.read_excel(file_in, sheet_name = sheet)
    col_list = df.columns.to_list()
    
    end_list = len(col_list)
    col_list = col_list[1:len(col_list)]
    targ_list = col_list[len(col_list)-5:len(col_list)]
    
    df_target = df[targ_list] # impute columns
    df_use = df[col_list]  # use to train by row
    
    np_target = df_target.values
    
    arr_nan = df_target.isnull().values.any(axis=1) + 0  # make Boolean rows with NaNs as 1
    arr_nan = np.argwhere(arr_nan == 1) # find indices of NaN rows
    
    for n in arr_nan:
        i = int(n)
        
        i_end = i + 8
        if i_end > arr_nan.max():
            i_end = arr_nan.max()
        
        ### look ahead to see if there are more than 8 gaps 
        eight_ahead = (np.isnan(np_target[i:i_end,:]) + 0).sum(axis=0)
        arr_eight = np.argwhere(eight_ahead == 8) # find indices of NaN rows
        
        a_row = np.asmatrix(df_use.iloc[i].values).T  # take row for imputer training        
        imputer = imp.fit(a_row)   ### train imputer on target row
        new_row = imp.transform(np.asmatrix(df_target.iloc[i].values).T).T
        
        if eight_ahead.max() < 8:
            
            np_target[i,:] = new_row  # replace row with imputed values
            
        if eight_ahead.max() == 8:
            arr_eight = np.argwhere(eight_ahead != 8) # find indices of non-max value
            
            for m in arr_eight:
                j = int(m)
                np_target[i,j] = new_row[0,j]  # replace row with imputed values
            
        
    df_new = pd.DataFrame(np_target)
    df_new.columns = targ_list
    df_new.to_excel(writer, sheet_name = sheet) # save to Excel
    
writer.save()  # save workbook
writer.close() # close workbook
print('\nSaving in file: ' + file_out)
    
    

