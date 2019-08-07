# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 14:07:04 2019
Calculates hotspots based on ranked concentrations
@author: Brian
Python 3.5
"""

import os
import pandas as pd
from pandas import read_csv
from openpyxl import load_workbook

current_path = os.getcwd()  # get current working path to save later

#file paths for different runs
#file paths for different runs
fp2017 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2017\\Current-2017_post\\'
fp2016 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2016\\Current-2016_post\\'
fp2015 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2015\\Current-2015_post\\'
fp2014 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2014\\Current-2014_post\\'
fp2013 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2013\\Current-2013_post\\'
fp2012 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2012\\Current-2012_post\\'
fp2011 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2011\\Current-2011_post\\'
fp2010 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2010\\Current-2010_post\\'
fp2009 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2009\\Current-2009_post\\'
fp2008 = 'C:\\TARS\AAAActive\\Qatar Airshed Study Jul 2017\\CALPUFF-MIC\\Current\\2008\\Current-2008_post\\'

paths = [fp2017, fp2016, fp2015, fp2014, fp2013, fp2012, fp2011, fp2010, fp2009, fp2008] # make filepath list

#make variables of chemicals
nh3 = 'NH3'; hf = '7664393'; h2s = '7783064'; nox = 'NOX'; pm10 = 'PM10-PRI';
so2 = 'SO2'; voc = 'VOC';
chem_type = [nh3, hf, h2s, nox, pm10, so2, voc]
ch = len(chem_type)

#make variables of time averages
hr1 = '1HR'; hr3 = '3HR'; hr24 = '24HR'; hr8760 = '8760HR'; hr8784 = '8784HR'
time_ave = [hr1, hr3, hr24, hr8760]
tave = len(time_ave)

df2 = pd.DataFrame() # create blank dataframe to use for concatenating

ranklist = [] #initialize list to capture Rank column names

#saves sheets in an excel file

filename_out = 'hotspots_output.xlsx'

os.chdir(current_path) #reset to original working path


book = load_workbook(filename_out)
writer = pd.ExcelWriter(filename_out, engine = 'openpyxl')

#writer = pd.ExcelWriter(filename_out)

for chem_name in range(ch):    # select chemical and loop  
    ranklist = [] #initialize list to capture Rank column names
    chem = chem_type[chem_name]
    print('\nProcessing ' + chem + '..')
    year = 2017 #start with most recent year
    
    for i in range(len(paths)):     #loop over different runs in different folders
        file_path = paths[i]
    
        yr = str(year)
        print('Processing year ' + yr + '...')

        os.chdir(file_path + chem) ## set filepath and folder
    
        for t in range (tave):
    
            time = time_ave[t]  #set averaging period
            
            #compensate for leap years
            leap_year = False
            if paths[i] == fp2016 or paths[i] == fp2012 or paths[i] == fp2008:
                leap_year = True
                time = hr8784
            
            #load data file of chemical and time average from run
            if time == hr8760 or time == hr8784:
                rank = 'RANK(0)'
            else:
                rank = 'RANK(ALL)'
            
            #load data file of chemical and time average from run
            data_file_name = rank + '_' + chem + '_' + time + '_CONC_C.DAT'
            df = read_csv(data_file_name, engine='python', skiprows=4, sep = ' ')
            df = df.dropna(axis=1, how='all')
            
            chem_n = chem + '_' + time + '_' + yr #create name of run
            
            #prepare input data by changing column name, adding index as column and sorting
            #chemical concentration from max to min (descending)
            if time == hr8760 or time == hr8784:
                df.columns = ['X_' + chem_n, 'Y_' + chem_n, chem_n]
            else:
                df.columns = ['X_' + chem_n, 'Y_' + chem_n, chem_n, 'Lat_' + chem_n, 'Long_' + chem_n]
 
            df.drop(0, inplace=True)
            df = df.rename_axis('ID').reset_index()
            df.sort_values(by=[chem_n], ascending = False, inplace=True)
            
            #Extract only location ID and chemical concentration from main dataframe
            #Add an index to act as Ranking value column
            df1 = df[['ID', chem_n]]
            df1.columns = ['ID_' + chem_n, chem_n]
            df1 = df1.reset_index()
            df1 = df1.drop('index', axis=1)
            df1 = df1.reset_index()
            df1.columns = ['Rank_' + chem_n, 'ID_' + chem_n, chem_n]
            
            #add rank name to rank list
            ranklist.append('Rank_' + chem_n)
            
            #sort again to put locations in order from 0 to n
            df1.sort_values(by=['ID_' + chem_n], ascending = True, inplace=True)
            df1 = df1.reset_index()
            df1 = df1.drop('index', axis=1)
            
            #add new dataframe to master list
            df2 = pd.concat([df2, df1], axis=1)
            
        year -= 1 #go to next year
            
    df2 = df2.reset_index()
    df2 = df2.drop('index', axis=1)
    
    # average the ranks row-wise into new dataframe and sort ascending
    df3 = df2[ranklist].mean(axis=1)
    df3 = df3.reset_index()
    df3.columns = ['Loc_ID', chem + '_Rank_ave']
    df3.sort_values(by=[chem + '_Rank_ave'], ascending = True, inplace=True)
    df3 = df3.reset_index()
    df3 = df3.drop('index', axis=1)
    
    hotspots = df3.reset_index()
    hotspots = hotspots.drop('index', axis=1)
    hotspots.sort_values(by=['Loc_ID'], ascending = True, inplace=True)
    hotspots = hotspots.reset_index()
    hotspots = hotspots.drop('index', axis=1)
    
    df.sort_values(by=['ID'], ascending = True, inplace=True)
    
    #coords = df[['Lat_' + chem_n, 'Long_' + chem_n]]
    #coords.columns = ['Latitude', 'Longitude']
    
    #hotspots2 = pd.concat([hotspots, coords], axis=1)
    
    print('Processing complete')
    print('Saving sheet...')
    
    ### write results to spreadsheet in Excel
    os.chdir(current_path) 
    writer.book = book
    hotspots.to_excel(writer, chem)

    writer.save()

writer.close()

#hotspots.to_excel(r'hotspots_output.xlsx',)