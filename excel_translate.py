# -*- coding: utf-8 -*-
"""
Created on Tue Jun  8 16:37:21 2021

@author: Brian
"""
import glob
import pandas as pd
from google_trans_new import google_translator 
from copy import deepcopy 

def translate(trans_txt):
    # translate text
    return translator.translate(trans_txt, lang_src= lang_org, lang_tgt='en') 

def translate_list(d_list):
    
    new_list_names = []
    for name in d_list:
         eng_name = translate(name)
         new_list_names.append(eng_name)
    return new_list_names


def translate_sheet(df):
    
    #deal with column names first
    col_names = list(df)
    '''
    new_col_names = []
    for name in col_names:
         eng_name = translate(name)
         new_col_names.append(eng_name)
    '''
    #rename columns    
    df.columns = translate_list(col_names)
    
    # translate content
    rows_n = len(df)
    print('There are ', rows_n, 'rows\n')
    cols_n = len(col_names)
    
    # pick a row
    for i in range (rows_n):
        print('Processing row', i)
        temp_row = deepcopy(df.loc[i])
        
        # pick the cell in the row
        for j in range (cols_n):
            
            # only translate if it is a string
            if isinstance(temp_row[j], str) == True:
                temp_row[j] = translate(temp_row[j])
                
        # replace row with translated row
        df.loc[i] = temp_row
    
    return df
    
translator = google_translator() 
lang_org = 'he' ## Hebrew = 'he'

# load file
for file in glob.glob("*.xls"):
    
    print('\nLoading ' + file + '....\n')
    
    f = pd.ExcelFile(file)
    sheet_names = f.parse()
    
    # translate sheet names for later use
    new_sheet_names = translate_list(sheet_names)
    
    for sheet in sheet_names:
        
        print('Processing sheet')
    
        df = pd.read_excel(file, sheet_name = sheet)
    
        new_df = translate_sheet(df)
    
    '''
    #deal with column names first
    col_names = list(df)
    
    new_col_names = []
    for name in col_names:
         eng_name = translate(name)
         new_col_names.append(eng_name)
    
    #rename columns    
    df.columns = new_col_names
    
    # translate content
    rows_n = len(df)
    print('There are ', rows_n, 'rows\n')
    cols_n = len(col_names)
    
    # pick a row
    for i in range (rows_n):
        print('Row', i)
        temp_row = df.loc[i]
        
        # pick the cell in the row
        for j in range (len(temp_row)):
            
            # only translate if it is a string
            if isinstance(temp_row[j], str) == True:
                temp_row[j] = translate(temp_row[j])
                
        # replace row with translated row
        df.loc[i] = temp_row
    '''    
