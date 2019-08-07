# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 10:46:58 2019

Calculates HQ for sensitive receptors
Uses exceedances in exceedances.py - so put in same folder

@author: Brian/Python 3.5
"""


import os
import pandas as pd
import numpy as np
from pandas import read_csv
from pandas import ExcelWriter
import exceedances as ex

current_path = os.getcwd()  # get current working path to save later

#get master list of grids
grid_file = 'Discrete-Receptors.xls'
grid_df = pd.read_excel(grid_file, sheet_name='Grid')

#get receptors with assigned grid locations
dist_file = 'receptor_grids.xlsx'
dist_df = pd.read_excel(dist_file)
dist_df = dist_df.reset_index()
dist_df = dist_df.drop('index', axis=1)

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
#time_ave = [hr1] # for acute HI
time_ave = [hr8760] # for HI
#time_ave = [h1, hr3, hr24, hr8760] #hotspots

tave = len(time_ave)

conc_df = pd.DataFrame(grid_df.LOC_ID) # create blank dataframe to use for concatenating

ranklist = [] #initialize list to capture Rank column names

for chem_name in range(ch):    # select chemical and loop  
    chem = chem_type[chem_name]
    
    for i in range(len(paths)):     #loop over different runs in different folders
        file_path = paths[i]
    
        yr = 1

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
            
            chem_n = chem + '_' + time + '_' + str(yr) #create name of run
            
            #prepare input data by changing column name, adding index as column and sorting
            #chemical concentration from max to min (descending)
            if time == hr8760 or time == hr8784:
                df.columns = ['X_' + chem_n, 'Y_' + chem_n, chem_n]
            else:
                df.columns = ['X_' + chem_n, 'Y_' + chem_n, chem_n, 'Lat_' + chem_n, 'Long_' + chem_n]
 
            df.drop(0, inplace=True) #remove column
            df = df.rename_axis('ID').reset_index()
            #df.sort_values(by=[chem_n], ascending = False, inplace=True)
            
            #regulatory threshold for chemical and time
            thresh1 = ex.mic_exceedance(chem, time)
            
            #Extract only chemical concentration from main dataframe
            df1 = df[[chem_n]]
            df1.columns = [chem_n]
            
            if thresh1 != 1.:
                df1 = df1/thresh1  #calculate Hazard quotient (HQ) = Ave Conc/RfC
            else:
                break
            
            conc_df = pd.concat([conc_df, df1], axis=1)

#find grid rows that have the receptor locations
receptor_conc = pd.DataFrame()            
for i in range(len(dist_df)):
    grid_desc = dist_df.Grid_dist[i]
    grid_conc = conc_df.query('LOC_ID == "{0}"'.format(grid_desc))
    receptor_conc = pd.concat([receptor_conc, grid_conc], axis=0)
    receptor_conc = receptor_conc.reset_index()
    receptor_conc = receptor_conc.drop('index', axis=1)


conc_df['HQ'] = conc_df.sum(axis=1)  #sums individual column HQ for aggregated HQ

# save to Excel
os.chdir(current_path) #reset to original working path
conc_df.to_excel(r'QRA_country_results.xlsx')
    