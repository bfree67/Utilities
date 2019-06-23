# -*- coding: utf-8 -*-
"""
Created on Thu Jun 21 14:07:04 2019

Ranks gridded locations based on summing maximum concentrations
Sorts based on individual chemicals and stores prioritized locations in a common
Excel worksheet

@author: Brian
Python 3.5
"""

import os
import pandas as pd
from pandas import read_csv
from pandas import ExcelWriter

#file paths for different runs (add extra \ to tree separators for Python)
fp2 = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\AQMIS Runs\\3yrCurrent\\3yr_current_post\\'
fp1 = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\AQMIS Runs\\3yr-all-MIC\\3yr-all-MIC_post\\'
fp3 = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\AQMIS Runs\\10yr-all-MIC-facilities_post\\'
paths = [fp2]#, fp2, fp3]  # make filepath list

#make variables of chemicals
nh3 = 'NH3'; hf = '7664393'; h2s = '7783064'; nox = 'NOX'; pm10 = 'PM10-PRI';
so2 = 'SO2'; voc = 'VOC';
chem_type = [nh3, hf, h2s, nox, pm10, so2, voc]
ch = len(chem_type)

#make variables of time averages
hr1 = '1HR'; hr3 = '3HR'; hr24 = '24HR'; hr8760 = '8760HR'
time_ave = [hr1, hr24, hr8760]
tave = len(time_ave)

# set Excel output file pathway
active_dir = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\AQMIS Runs\\'
writer = ExcelWriter('hotspots3_current.xlsx')

for chem_name in range(ch):    # select chemical and loop  
    chem = chem_type[chem_name]
    df2 = pd.DataFrame() # create blank dataframe to use for concatenating
    ranklist = [] #initialize list to capture Rank column names

    
    for i in range(len(paths)):     #loop over different runs in different folders
        file_path = paths[i]
    
        # set variable for years of runs
        if file_path == fp1:
            yr = 1
        if file_path == fp2:
            yr = 3
        if file_path == fp3:
            yr = 10


        os.chdir(file_path + chem) ## set filepath and folder
    
        for t in range (tave):
    
            time = time_ave[t]  #set averaging period
            
            #load data file of chemical and time average from run
            data_file_name = 'RANK(ALL)_' + chem + '_' + time + '_CONC_C.DAT'
            df = read_csv(data_file_name, engine='python', skiprows=4, sep = ' ')
            df = df.dropna(axis=1, how='all')
            
            chem_n = chem + '_' + time + '_' + str(yr) #create name of run
            
            #prepare input data by changing column name, adding index as column and sorting
            #chemical concentration from max to min (descending)
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
            
    df2 = df2.reset_index()
    
    # some ranks row-wise into new dataframe and sort ascending
    df3 = df2[ranklist].sum(axis=1)
    df3 = df3.reset_index()
    df3.columns = ['Loc_ID', 'Rank_sum']
    df3.sort_values(by=['Rank_sum'], ascending = True, inplace=True)
    df3 = df3.reset_index()
    df3 = df3.drop('index', axis=1)
    
    hotspots = df3.reset_index()
    hotspots = hotspots.drop('index', axis=1)
    hotspots.sort_values(by=['Loc_ID'], ascending = True, inplace=True)
    hotspots = hotspots.reset_index()
    hotspots = hotspots.drop('index', axis=1)

    df.sort_values(by=['ID'], ascending = True, inplace=True)
    
    coords = df[['Lat_' + chem_n, 'Long_' + chem_n]]
    coords.columns = ['Latitude', 'Longitude']
    
    hotspots2 = pd.concat([hotspots, coords], axis=1)
    hotspots2.sort_values(by=['Rank_sum'], ascending = True, inplace=True)
    hotspots2 = hotspots2.reset_index()
    hotspots2 = hotspots2.drop('index', axis=1)
    
    # save to chemical rankings to worksheet
    os.chdir(active_dir)
    hotspots2.to_excel(writer, sheet_name = chem)  

#close worksheet
writer.save()
writer.close()