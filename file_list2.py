# -*- coding: utf-8 -*-
"""
Created on Wed Aug 21 14:53:06 2019
Reads excel files that have some EDD column headers from multiple xls files
in a folder. The read files must have an .xls extension or will be ignored.

The files are merged into a super dataframe, agg. 

The EDD format is read as an xlsx file (KEPA.xlsx).

The agg dataframe is broken up into Field and Lab EDD formats. Blank columns
are inserted if there is no match in the agg dataframe.

@author: Brian, Python 3.5
"""
import glob
import pandas as pd

txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

df_lab = pd.read_excel('KEPA.xlsx', sheet_name = 0)
lst_lab = list(df_lab)

df_field = pd.read_excel('KEPA.xlsx', sheet_name = 1)
lst_field = list(df_field)

# read the excel files in the directory and save to a list
for file in glob.glob("*.xls"):
    txtfiles.append(file)
    

    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    
    print(file, res)
    
    if i == 0:   # on the first sheet of the first file, read only
        agg = pd.read_excel(file, sheet_name = 0)
        i += 1
        
    else:
        
        if i == 1 and res > 1:  # on the next sheet, start one sheet over to avoid duplicate read
            
            for j in range (1, res):
                            
                agg1 = pd.read_excel(file, sheet_name = j)
                agg = pd.concat([agg,agg1])
                i += 1
                
        else: 
            
            for j in range (res):
                            
                agg1 = pd.read_excel(file, sheet_name = j)
                agg = pd.concat([agg,agg1])
                i += 1
                
# ******* Make Field EDD
                
df_newField = pd.DataFrame() # make blank dataframe

for field_col in lst_field: # list of field EDD columns
    count = 0 # reset counter used to see if there is matching column
    for agg_col in list(agg): #list of existing columns
        if field_col == agg_col and count == 0: # if columns match, add column to dataframe
            count += 1
            df_newField[field_col] = agg[field_col].tolist()
            
        if field_col != agg_col and count > 0: # if no match, break
            break
        
    if count == 0:  # if no match, add a blank column
         df_newField[field_col] = ""

# ******* Make Lab EDD
                
df_newLab = pd.DataFrame()
for lab_col in lst_lab:
    count = 0
    for agg_col in list(agg):
        if lab_col == agg_col and count == 0:
            count += 1
            df_newLab[lab_col] = agg[lab_col].tolist()
            
        if lab_col != agg_col and count > 0:
            break
        
    if count == 0:
         df_newLab[lab_col] = ""       
            
        

writer = pd.ExcelWriter('Kabd_results.xlsx')
df_newLab.to_excel(writer, sheet_name = 'KEPA_lab', index = False)
df_newField.to_excel(writer,sheet_name = 'KEPA_field', index = False)
writer.save()
