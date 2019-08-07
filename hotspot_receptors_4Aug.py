# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 15:07:16 2019

Takes prepared Excel files with coordinates of all gridded receptors
and results from Quantitative Risk Assesment (QRA) from other output.
USE country-qra.py to run QRA

Creates a distance matrix for each sensitive receptor and discrete/gridded receptor

Converts 0 values in the distance matrix to 10 m for calculation purposes

Hotspot = QRA/Distance matrix

@author: Brian/Python 3.5
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans

current_path = os.getcwd()  # get current working path to save later

#list of all possible data types
floats = ['float16', 'float32', 'float64']

#set filename for discrete/gridded and sensitive receptors (same file - different sheets)
grid_file = 'Discrete-Receptors.xls' # outliers (11 & 13) removed

#upload QRA results
hq_file = 'QRA_country_results.xlsx' # acute HQ file

#upload gridded receptors
grid_df = pd.read_excel(grid_file, sheet_name='Grid')

#upload sensitive receptors
receptors_df = pd.read_excel(grid_file, sheet_name='Receptor')

#upload QRA results
total_hq_df = pd.read_excel(hq_file)

acute_df = total_hq_df[['LOC_ID', 'HQ']]

#transpose for convenience
x_grid = np.asmatrix((grid_df.X)).T; y_grid = np.asmatrix((grid_df.Y)).T
x_rec = np.asmatrix((receptors_df.X)).T; y_rec = np.asmatrix((receptors_df.Y)).T

# find UTM distance from grid center to receptor
dist_df = pd.DataFrame(grid_df.LOC_ID)
dist = np.zeros((len(x_grid), len(x_rec))) #make distance matrix

for i in range (len(x_rec)):
    dist[:,i] = (np.sqrt(np.power(x_grid - x_rec[i],2) + np.power(y_grid - y_rec[i],2))).T
    
# calculate distance matrix    
dist_df = pd.concat([dist_df, pd.DataFrame(dist)], axis=1)
dist_df = dist_df.replace(0, 0.01) # set small distances to 10m for calculations

#convert to numpy matrices
dist = (dist_df.select_dtypes(include= floats)).values
acuteHQ = (acute_df.select_dtypes(include= floats)).values

#c = np.divide(acuteHQ,dist) #divide HQ/Distance
c = acuteHQ

#create new dataframe with grid ID as first row - take average of rows (HQ/distance)
c_df = pd.concat([grid_df.LOC_ID, pd.DataFrame(c).mean(axis=1)], axis = 1)
c_df.columns = ['LOC_ID', 'DistHQ']
c_df['DistHQ'] = c_df['DistHQ'].apply(lambda x: round(x, 2))
c_df = c_df.reset_index()
c_df = c_df.drop('index', axis=1)

hq_df = pd.concat([grid_df.X, grid_df.Y,c_df.DistHQ, grid_df.Lat, grid_df.Long], axis = 1)
hq_df.columns = ['X', 'Y', '1 RANK', 'NLAT_WGS84', 'ELON_WGS84']

#c_df.sort_values(by=['DistHQ'], ascending = True, inplace=True)

# save to Excel
os.chdir(current_path) #reset to original working path
#hq_df.to_excel(r'grid_dist_acute.xlsx', index = False)  # this saves an unformated worksheet

#### make CALPOST style CONC file
hq_np = hq_df.values
file1 = open("RANK(ALL)_HI_1HR_CONC_C.DAT","w") 
#make header
file1.write(" 1 RANKED     1   HOUR AVERAGE  CONCENTRATION VALUES AT EACH RECEPTOR   \n")
file1.write('\n')
file1.write('         HQ          1     (1/m)    \n')
file1.write('         (24 hours/day processed) \n')
file1.write('     RECEPTOR (x,y) km        1 RANK     NLAT_WGS84  ELON_WGS84 \n')
file1.write('\n')

#write data
for i in range(len(hq_np)):
    Line = '     ' + str(hq_np[i,0]) + '  ' + str(hq_np[i,1]) + '     ' + \
        str(hq_np[i,2]) + '      ' + str(hq_np[i,3]) + '  ' + str(hq_np[i,4]) + '\n'
    
    file1.writelines(Line) 
    
file1.close()  