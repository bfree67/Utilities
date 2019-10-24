# -*- coding: utf-8 -*-
"""
Created on 18 Oct 2019

Utility to reduce sources based on emission contributions. Takes complete Facility Template
and removes unused release points and emission units.
Fixed bug that cut off first row of data.
And stupid shit...this one works

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import time
from copy import deepcopy

version = '2.3'
# global constants for country and company

current_path = os.getcwd()  # get current working path to save later

# function to add proper header in new template format after data has been remapped
def new_header(df):
    hd_list = list(df) # get column headers (name, unknown,,,,,)
    for i in range(1,len(hd_list)):
        hd_list[i] = '' # make blank
    
    df.columns = df.iloc[0] # make new column headers from first row
    df2 = pd.DataFrame(hd_list).T
    df2.columns = df.iloc[0]
    df3 = pd.concat((df2,df), axis = 0)
    df3 = drop_index(df3)
    return df3.iloc[0:3]

def new_header2(df):
    #df.columns = df.iloc[0] # make new column headers from first row
    return df.iloc[0:2]
    
# function to reset dataframe index and drop added index column
def drop_index(df):
    df = df.reset_index()
    return df.drop('index', axis=1)

def strip_data(x_string):
    
    new_sheet = new_template.parse(x_string)
    
    # strip header and name columns
    new_sheet.columns = new_sheet.iloc[0]
    return new_sheet.iloc[2:len(new_sheet)]

def reconstruct(new_sheet_hd, new_sheet):
    new_sheet = pd.concat((new_sheet_hd, new_sheet), axis = 0)
    return drop_index(new_sheet)

def make_percentile(pollutant, percentile):
    # *Get Emission sheet
    new_emissions = strip_data('Emissions')
    
    # set up cumulative columns
    es = deepcopy(new_emissions.loc[new_emissions['Pollutant'] == pollutant])
    es = drop_index(es.sort_values(by=['Emission'], ascending=False))
    es['cum_sum'] = es['Emission'].cumsum()
    es['cum_perc'] = 100*es['cum_sum']/es['Emission'].sum()
    
    # select only the top percentile% contributors
    es['top'] = (es['cum_perc'] < percentile) + 0 #( 1- if less than 96%, 0 if greater))
    es1 = deepcopy(es[es.top == 1])
    
    print('Top ' + str(percentile) + '% for ' + pollutant + ' gives ' + str(es1['Emission'].count()) )
    es2 = deepcopy(es[es.top == 0])
    
    es2.loc[:,'Emission'] = 0. # 0 out the bottom percentile
    
    es1 = es1.drop('top', axis=1)
    es1 = es1.drop('cum_perc', axis=1)
    es1 = es1.drop('cum_sum', axis=1)
    
    es2 = es2.drop('top', axis=1)
    es2 = es2.drop('cum_perc', axis=1)
    es2 = es2.drop('cum_sum', axis=1)
    
    new_es = drop_index(pd.concat((es1,es2), axis = 0))
    
    #get other emissions from the same emissions units
    esn = deepcopy(new_emissions.loc[new_emissions['Pollutant'] != pollutant])
    
    esn.loc[:,'Emission'] = 0.
    
    df = drop_index(pd.concat((new_es, esn), axis = 0))
    
    return drop_index(df.sort_values(by=['Unit ID', 'Pollutant'], ascending=True))

    
# ***************** START PROGRAM *********************************

print('*************************************************\n')
print('       Source Reduction Utility  ')
print('                Version ' + version + '\n')
print('*************************************************\n')

# ******************* make new headers ****************************

# location of blank template with only column headers    
new_file = 'Normal.xlsx'    
new_template = pd.ExcelFile(new_file)
new_names = new_template.sheet_names

# get template headers 

new_emissions_hd = new_template.parse(new_names[6])

# make new headers 

new_emissions_hd = new_header(new_emissions_hd)

voc = make_percentile('Volatile Organic Compounds', 96.)

nox = make_percentile('Nitrogen Oxides', 90.)

new_emissions1 = deepcopy(voc)
new_emissions1['Emission'] = voc['Emission'] + nox['Emission']
#
# **************************************************
# Save in new AQMIS format as xlsx 
save_df = input('\nSave reduced AQMIS sources? (y/n) ')


if save_df == 'y':   
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    
    writer = pd.ExcelWriter('Reduced_format_'+time_stamp+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    new_emissions1.to_excel(writer,sheet_name = 'Emissions', index = False, header = False)
    voc.to_excel(writer,sheet_name = 'VOC', index = False, header = False)
    nox.to_excel(writer,sheet_name = 'NOX', index = False, header = False)
    writer.save()
