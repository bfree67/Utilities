# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 20:55:11 2019 Python 3.5
Reads Excel file with AQM data and converts ppb into ug/m3
Also creates rows to groupby dates
@author: Brian
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import calendar
import os
import time

def MW_gmole(MW_name):
    MWg =	{
      "NO": 30.01,
      "NO2": 46.0055,
      "NOx": 1964,
      "SO2": 64.066,
      "H2S": 34.1,
      "O3": 48,
      "CH4": 16.04,
      "CO": 28.01
    }
    return MWg[MW_name]

def ppb2ugm3(Cppb,P, T, MW_name):
    
    R = 8.3144 ### ideal gas constant in m3-kPa/K-mol
    
    MW = MW_gmole(MW_name)  ## get MW from MW_gmole function

    Cug = round((Cppb*MW*P*0.01)/(R*T),2) ## convert to ug/m3
    
    return Cug

########################Start program execution
start = time.time()

sheets = ['2013', '2014','2015', '2016']

Site_all = pd.DataFrame()

for i in range(len(sheets)):
    
    print('Loading Year ' + sheets[i] + ' ....')
    ##### read data file and choose chemical of interest
    Site = pd.read_excel('FS_AQMS.xlsx', sheet_name= sheets[i])
    
    #### break out time and data from time/date
    Site['date'] = pd.to_datetime(Site['Time']).dt.strftime('%D')
    Site['day'] = pd.to_datetime(Site['Time']).dt.strftime('%d')
    Site['month'] = pd.to_datetime(Site['Time']).dt.strftime('%m')
    Site['year'] = pd.to_datetime(Site['Time']).dt.strftime('%Y')
    
    #Site['day_of_week_n'] = Site['Time'].apply(lambda x: x.weekday()) # get the weekday index, between 0 and 6
    #Site['day_of_week'] = Site['day_of_week_n'].apply(lambda x: calendar.day_name[x])  ## weekday name
    Site['hour'] = pd.to_datetime(Site['Time']).dt.strftime('%H')
    
    chemical = ['NO', 'NO2', 'SO2', 'O3', 'H2S', 'CH4', 'CO']
    
    ## create new columns with chemical names that can be converted into ug/m3
    for i in range(len(chemical)):
        
        chem_name = chemical[i]
    
        Site[chem_name+'-ugm3'] = ppb2ugm3(Site[chem_name + '-ppb'], Site['BP-hpa'], Site['Temp-degC'], chem_name)
    
    Site['CO-mgm3'] = Site['CO-ugm3'] * 0.001
    Site['CH4-mgm3'] = Site['CH4-ugm3'] * 0.001
    Site['H2S-mgm3'] = Site['H2S-ugm3'] * 0.001
    
    Site['BP-kpa'] = Site['BP-hpa'] * 0.1
    
    Site_all = Site_all.append(Site)  ## collect each year

### reset index
Site_all = Site_all.reset_index(drop=True)

end = time.time()

duration = round(end - start,2) # total time in seconds

print('Loaded and processed in ' + str(duration) + ' seconds.\n')

### write results to spreadsheet in Excel
file_saved = 0
while file_saved == 0:
    try:
        filename_out = 'AMS-rawdata-out'
        writer = pd.ExcelWriter(filename_out + '.xlsx')
        Site_all.to_excel(writer,'FS')
        writer.save()
        file_saved = 1
        print('File ' + filename_out + '.xlsx saved')
        
    except:
        print('File ' + filename_out + '.xls is open. Close it and try again.')

#### make revised dataframe with sequence for AQMIS - no ppb only fields like NMHC
Site_new = Site_all[['date', 
                    'hour', 
                    'NO-ugm3',  
                    'O3-ugm3',
                    'SO2-ugm3',
                    'NO2-ugm3', 
                    'CH4-mgm3',
                    'H2S-mgm3',  
                    'CO-mgm3', 
                    'PM10-ug/m3', 
                    'WS-m/s', 
                    'WD', 
                    'Temp-degC', 
                    'RH-percent', 
                    'BP-kpa', 
                    'SW-w/m2']].copy()

Site_new = Site_new.replace(np.inf, '')

### write results to spreadsheet in Excel
file_saved = 0
while file_saved == 0:
    try:
        filename_out = 'AMS-AQMISdata-out'
        writer = pd.ExcelWriter(filename_out + '.xlsx')
        Site_new.to_excel(writer,'FS')
        writer.save()
        file_saved = 1
        print('File ' + filename_out + '.xlsx saved')
        
    except:
        print('File ' + filename_out + '.xls is open. Close it and try again.')