# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 10:46:58 2019

Inputs Excel file with Grid and Receptor worksheets
Creates distance based ranking for each sensitive receptor to each grid
Outputs Excel file with Receptors and nearest grid


@author: Brian/Python 3.5
"""

import os
import pandas as pd
import numpy as np
from pandas import read_csv
from pandas import ExcelWriter

grid_file = 'Discrete-Receptors.xls'
grid_df = pd.read_excel(grid_file, sheet_name='Grid')
receptors_df = pd.read_excel(grid_file, sheet_name='Receptor')

x_grid = np.asmatrix((grid_df.X)).T; y_grid = np.asmatrix((grid_df.Y)).T
x_rec = np.asmatrix((receptors_df.X)).T; y_rec = np.asmatrix((receptors_df.Y)).T

# find UTM distance from grid center to receptor
dist_df = pd.DataFrame(grid_df.LOC_ID)
dist = np.zeros((len(x_grid), len(x_rec))) #make distance matrix
for i in range (len(x_rec)):
    dist[:,i] = (np.sqrt(np.power(x_grid - x_rec[i],2) + np.power(y_grid - y_rec[i],2))).T

# distance matrix    
dist_df = pd.concat([dist_df, pd.DataFrame(dist)], axis=1)

a = [] # initiate list

for i in range (1,len(dist_df.T)): # loop through receptors by column
    rec_dist = dist_df[dist_df.columns[i]]
    
    close_df = pd.concat([dist_df.LOC_ID, rec_dist], axis=1) # make new dataframe with loc ID and distance
    close_df.columns = ['LOC_ID', 'Dist']
    close_df.sort_values(by=['Dist'], ascending = True, inplace=True)
    
    
    ac = close_df.reset_index()
    ac = ac.reset_index()
    ac.sort_values(by=['index'], ascending = True, inplace=True)    


    closest_grd = close_df.head(1) # take first row (closest grid to receptor)
    a.append(closest_grd.iloc[0]['LOC_ID']) # add to list

r_names = list(receptors_df) # get list of column names
r_names.append('Grid_dist') # add new name
receptors_df = pd.concat([receptors_df, pd.DataFrame(a)], axis=1)

receptors_df.columns = [r_names] #rename with updated name list

# save to Excel
receptors_df.to_excel(r'receptor_grids.xlsx')