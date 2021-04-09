# -*- coding: utf-8 -*-
"""
Created on Wed Apr  7 16:50:13 2021

For CALPUFF
Reads each CON.Dat file and finds ave values with inner & outer radii from source


@author: Brian
"""
import os
import glob
import pandas as pd
from pandas import read_csv
import copy
import matplotlib.pyplot as plt

org_path = os.getcwd()

# get list of gridded receptors in order
folder = ['TCEQ_scenario', 'SCREEN3_scenario', 'Null_scenario']

writer = pd.ExcelWriter('calpuff_ave_results.xlsx')

distance = [5,10,15,20,25,30,35,40, 45, 50]
inner = distance[0:(len(distance)-1)]
outer = distance[1:len(distance)]
df_total = pd.DataFrame({'inner':inner, 'outer':outer})

for scenario in folder:

    pathname = 'C:/Users/Brian\\Documents/_Active/BOEM/Flares/Calpuff/' + scenario + '/'+ scenario +'_post/SO2/'
    os.chdir(pathname)
    
    scen = scenario[0:len(scenario)-9]
    print(scen)
    
    out ={}

        
    for grid_file in glob.glob("*.DAT"):   #read all the files in the directory with .DAT extension
    
        print(grid_file)
        df_new= read_csv(grid_file, engine='python', skiprows=4, delim_whitespace=True)
        new_cols = list(df_new)
        df_grid = df_new[new_cols[0:3]]
        df_grid = df_grid.reset_index(drop = True)
        df_grid.columns = ['x', 'y', 'conc']
        
        # add column with distance from source (source at 0,0)
        df_grid['distance'] = (df_grid['x']**2 + df_grid['y']**2)**.5
        
        
        
        df_out = pd.DataFrame({'inner':inner, 'outer':outer})
        
        ave_out =[]
        for i in range(1,len(distance)):
            near = distance[i-1] # inner radius
            far = distance[i] # outer radius
            
            # find values within the 2 radii 
            df_ave = df_grid.loc[(df_grid['distance'] >= near) & (df_grid['distance'] <= far)]
            
            # take average
            ave = df_ave['conc'].mean()
            ave_out.append(ave)
            
            print(near, far, ave)
        
        grid_name = copy.deepcopy(grid_file)
        
        out.update({grid_name:ave_out})
        print('\n')
        
    os.chdir(org_path)   
    df_conc = pd.DataFrame(out)
    
    df_conc.columns = ['annual_'+scen, '1hr_'+scen, '24hr_'+scen]
    
    df_out = pd.concat((df_out, df_conc), axis=1)
    df_out.to_excel(writer, sheet_name = scenario, index = False, header = True)
    

    df_total = pd.concat((df_total, df_conc), axis = 1)
    
    
    print(df_conc)
df_total.to_excel(writer, sheet_name = 'summary', index = False, header = True)  
writer.save()

name_list = list(df_total)
fig,ax = plt.subplots()
for name in [name_list[2],name_list[5],name_list[8]]:   #annual
#for name in [name_list[3],name_list[6],name_list[9]]:   #1hr
#for name in [name_list[4],name_list[7],name_list[10]]:   #24hr
    ax.plot(df_total.inner,df_total[name],label=name)

ax.set_xlabel("Inner radius [m]")
ax.set_ylabel("Ave concentration [$\mu g / m^{3}$]")
ax.legend(loc='best')
    
            

        
        

