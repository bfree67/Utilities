# -*- coding: utf-8 -*-
"""
Created on Thu Oct 31 10:35:50 2019
Lists all csv files in a folder, opens them up and finds the largest concentration and deposition value
Saves lat/long in a list
Converts list into dataframe by first placing lat/long values into a numpy matrix
Then groups on uniques Lat/Long combinations and counts number of max values

@author: Brian, Python 3.5
4 Nov 2019
"""

import pandas as pd
import glob
import time
import numpy as np

sps_files = []
c_max = []
d_max = []

for file in glob.glob("*.csv"):
    sps_files.append(file)

    start = time.time()
    df = pd.read_csv(file)
    df_col = list(df)
    df.head()
    end = time.time()
        
    c_max.append(df.iloc[df[df_col[3]].idxmax()])
    d_max.append(df.iloc[df[df_col[4]].idxmax()])
    
    print(file + ' finished in ' + str( round( end-start,1)) + ' seconds')

'''
#print concentration max
print('\nConcentration max locations')
for i in range(len(c_max)):
    c = c_max[i]
    print(c[1],c[2])

#print deposition max
print('\nDeposition max locations')
for i in range(len(c_max)):
    c = d_max[i]
    print(c[1],c[2])
'''
  
df = pd.DataFrame()
cmat = np.zeros((2*len(c_max),2))

for i in range(len(c_max)):
    c = c_max[i]
    cmat[i,0]=c[1]
    cmat[i,1]=c[2]
    
for i in range(len(d_max), 2*len(d_max)):
    d = d_max[i-len(d_max)]
    cmat[i,0]=d[1]
    cmat[i,1]=d[2]
    
df = pd.DataFrame(cmat, columns = ['Long', 'Lat'])

#print list of counts and sort in descending order
df.groupby(['Long', 'Lat']).size().sort_values(ascending=False)