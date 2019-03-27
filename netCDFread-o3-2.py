# -*- coding: utf-8 -*-
"""
Created on Sun Sep 23 12:04:21 2018
Load HDF5/netCDF4 files and extract lat/long/values

@author: Brian

Run dataset.variables to find analyte name in top file list after OrderDict..
"""

from netCDF4 import Dataset
import numpy as np
from matplotlib import pyplot as plt
from matplotlib import dates
from datetime import datetime
import seaborn as sns
import os
import pandas as pd

year = ['2016','2017']
year_sel = year[1]
dir_list = ['Ozone2016','Ozone2017']
pollutant = 'o3'
units = 'Dobson Units' 
analyte_name = 'OMDOAO3e_003_ColumnAmountO3' #analyte variable
a_t = []

for pol_year in range(len(year)):
    active_year = year[pol_year]
    current_dir = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\Giovanni\\Ozone' + active_year
    
    print(current_dir)
    
    #change directory to get annual data
    os.chdir(current_dir)

    i = 1  #filenames start at 1, not 0
    
    #make initial filename
    file_name = pollutant+'-'+active_year+'-'+str(i)+'.nc'
        
    #get initial data to calculate matrix dimenstions
    dataset = Dataset(file_name)
    d_lat = np.asmatrix(dataset.variables['lat'][:])  #list of latitude coordinates
    d_lon = np.asmatrix(dataset.variables['lon'][:])  #list of longitude coordinates
    analyte_values = np.asmatrix(dataset.variables[analyte_name][:]) #analyte values in matrix
    
    n,p = np.shape(analyte_values)
    
    #initialize tensor
    analyte_tensor = np.zeros(((n,p,12)))  
    
    # include initial month into tensor
    analyte_tensor[:,:,0]=analyte_values 
    
    #create tensor from datafiles using netCDF4 Dataset command to load data
    for i in range(1,12):
        file_name = pollutant+'-'+active_year+'-'+str(i+1)+'.nc'
        print(file_name)
        dataset = Dataset(file_name)
        analyte_values = np.asmatrix(dataset.variables[analyte_name][:])
        analyte_tensor[:,:,i]=analyte_values
    
    #create list of 2 tensors    
    a_t.append(analyte_tensor)

#convert tensor list into 1 tensor
a_t3=(np.concatenate((a_t[0],a_t[1]), axis=2))
nt, pt, qt = np.shape(a_t3)

#### begin plotting
max_range = a_t3.max() * 1
min_range = a_t3.min() * 1

a_mean= a_t3.mean(axis=2)
a_std= a_t3.std(axis=2)

#### plot heatmaps of annual averages
ax = plt.axes()
sns.heatmap(a_mean, annot= False,  linewidths=.1, 
            vmin = min_range, vmax = max_range,
            cmap="YlGnBu",
            cbar_kws={'label': units},
            ax = ax)
ax.set_title('Average O$_{3}$ 2016-2017')
#plt.savefig('Ave_'+pollutant+'_2016-2017.png', dpi=300)
plt.show()

#### plot heatmaps of annual averages
ax = plt.axes()
sns.heatmap(a_std, annot= False,  linewidths=.1, 
            vmin = min_range, vmax = max_range,
            cmap="YlGnBu",
            cbar_kws={'label': units},
            ax = ax)
ax.set_title('Std Dev O$_{3}$ 2016-2017')
#plt.savefig('Std_'+pollutant+'_2016-2017.png', dpi=300)
plt.show()

#make an x range of months
df=pd.DataFrame({'x': pd.DatetimeIndex(start='2016', freq='M', periods=qt).map(lambda x: str(x.year) +'-'+ str(x.month))})

for i in range(nt):
    for j in range(pt):
        y = a_t3[i,j,:]
        y1 = str(j) + '/' + str(i)
        df[y1] = y

for column in df.drop('x', axis=1):
    
    #df.set_value(0,column, df.iloc[0][column] * 1000)
    #df.set_value(12,column, df.iloc[12][column] * 1000)
    plt.plot(df['x'], df[column], marker='', linewidth=1, alpha=0.9, label=column)
    #plt.scatter(df['x'], df[column])

# change directory to store images
os.chdir("C:\\TARS\\AAAActive\Qatar Airshed Study Jul 2017\\Giovanni\\Images")
#plt.savefig('AveDust'+year_sel+'.png', dpi=300)
#plt.legend(loc=2,ncol=4)
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.3),
          fancybox=True, shadow=True, ncol=5)
#plt.legend()
plt.xlabel("Month")
#labels = df['x']
plt.xticks(rotation=90)
plt.ylabel(units)
#plt.yscale('log')

#plt.savefig('TS_'+pollutant+'_2016-2017.png', dpi=300)
plt.show()