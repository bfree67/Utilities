# -*- coding: utf-8 -*-
"""
Created on Wed Aug 21 14:53:06 2019

Utility to remap AQMIS templates using old import format to new import format.
Requires all old template spreadsheets to be in the same folder using *.xlsx format
and a modified blank template with only column headers to in an other folder.

Updated 26 Aug 2019 - fixed AQZ merging and made into a function add_aqz
Reads column headers of first old template and changes headers of all
other files to ensure standardized column headers. Just have to update the
first file.

Updated 28 Aug 2019 - Added check to see if column headers in Release Point
are in degrees or deg and exit temp is K or C. Added check to see if there is an
active or inactive added column

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np


current_path = os.getcwd()  # get current working path to save later

# function to reset dataframe index and drop added index column
def drop_index(df):
    df = df.reset_index()
    return df.drop('index', axis=1)

def make_stats(so2):
    stats = pd.DataFrame()
    stats = so2.groupby(['Source Code']).sum()
    stats = pd.concat([stats, so2.groupby(['Source Code']).mean()], axis = 1, sort = False)
    stats = pd.concat([stats, so2.groupby(['Source Code']).std()], axis = 1, sort = False)
    stats = pd.concat([stats, so2.groupby(['Source Code']).max()], axis = 1, sort = False)
    stats = stats.drop(columns=['Month'])
    stats = pd.concat([stats, so2.groupby(['Source Code']).min()], axis = 1, sort = False)
    stats = stats.drop(columns=['Month'])
    stats90 = so2.groupby(['Source Code']).mean() + (so2.groupby(['Source Code']).mean() 
                                + (so2.groupby(['Source Code']).std() * 1.363 / 3.464))
    return pd.concat([stats, stats90], axis = 1, sort = False)
    
def scenario_dev(chem):
    chemical = chem.groupby(['Source Code'])
    scenarios = []
    for i, (chemical_name, chemical_gdf) in enumerate(chemical):
        a = []
        chemical_gdf.columns =['a', 'b', 'c']
        normal = float(round(chemical_gdf.mean(),3))
        
        std_chem = float(chemical_gdf.std())
        
        t_var = float((2*normal) + (std_chem * 1.363/3.464))
        
        abnormal = round((chemical_gdf.loc[(chemical_gdf['c'] > t_var)]).mean(),3)
        
        print(chemical_name, normal, abnormal.c)
        
        a.append(chemical_name)
        a.append(normal)
        a.append(abnormal.c)
        
        scenarios.append(a)
    df_scene = pd.DataFrame(scenarios)
    df_scene.columns = ['Source Code', 'Normal', 'Abnormal']
        
    return df_scene

    
# **************** make templates **************************************

company = 'QChem'
txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

# merge the excel files into one file by reading all excel files in the folder
for file in glob.glob("*.xlsx"):
    txtfiles.append(file)
    
    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    xl_names = xl.sheet_names
    
    if i == 0:
        df_fac = xl.parse(xl_names[0])


        i += 1     
    
    else:
        df_fac = pd.concat([df_fac, xl.parse(xl_names[0])], sort = False)    

    
    print(file)

df_fac = df_fac.sort_values(['Source Code', 'Month'])
so2 = df_fac[['Source Code', 'Month', '(as SO2)']]
nox = df_fac[['Source Code', 'Month', '(as NOx)']]
voc = df_fac[['Source Code', 'Month', '(VOC)']]
pm10 = df_fac[['Source Code', 'Month', '(PM)']]

# make stats worksheet
stats = make_stats(so2)
stats = pd.concat([stats, make_stats(nox)], axis = 1, sort = False)
stats = pd.concat([stats, make_stats(voc)], axis = 1, sort = False)
stats = pd.concat([stats, make_stats(pm10)], axis = 1, sort = False)

stats.columns = ['SO2 Sum', 'SO2 Ave', 'SO2 Std', 'SO2 Max', 'SO2 Min', 'SO2 CI', 
		'NOX Sum', 'NOX Ave', 'NOX Std', 'NOX Max', 'NOX Min', 'NOX CI',
		'VOC Sum', 'VOC Ave', 'VOC Std', 'VOC Max', 'VOC Min', 'VOC CI',
		'PM10 Sum', 'PM10 Ave', 'PM10 Std', 'PM10 Max', 'PM10 Min', 'PM10 CI']

# *********** plot histograms ***************************************

chem = 'SO2'

chem = so2.drop(columns=['Month'])
chemical = pm10.groupby(['Source Code'])
plt.figure()

fig = plt.figure(figsize=(8,16))
#fig.set_size_inches(8,5)
# Iterate through continents

for i, (chemical_name, chemical_gdf) in enumerate(chemical):

    # create subplot axes in a 3x3 grid
    ax = plt.subplot(17, 3, i + 1) # nrows, ncols, axes position
    # plot the continent on these axes
    chemical_gdf.hist(ax=ax)
    # set the title
    ax.set_title(chemical_name)

plt.tight_layout()
plt.show()


# ******************* Make scenarios *****************
    
df_scene = scenario_dev(so2)
df_scene = pd.concat([df_scene, scenario_dev(nox)], axis = 1, sort = False)
df_scene = pd.concat([df_scene, scenario_dev(voc)], axis = 1, sort = False)
df_scene = pd.concat([df_scene, scenario_dev(pm10)], axis = 1, sort = False)


# ************ Save worksheets ***************************

# Save in new AQMIS format as xlsx 
save_df = input('\nSave new sheet for ' + company +'? (y/n) ')

#new_path = 'C:\TARS\AAALakes\Kuwait AQMIS\'
if save_df == 'y':   
    os.chdir("\TARS\AAAActive\Qatar Airshed Study Jul 2017\Compare") # reset to new folder path 
    writer = pd.ExcelWriter('Annual_'+company+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    df_scene.to_excel(writer, sheet_name = 'Scenarios', index = True, header = True)
    stats.to_excel(writer, sheet_name = 'Stats', index = True, header = True)
    df_fac.to_excel(writer, sheet_name = 'All', index = False, header = True)
    so2.to_excel(writer, sheet_name = 'SO2', index = False, header = True)
    nox.to_excel(writer, sheet_name = 'NOx', index = False, header = True)
    voc.to_excel(writer, sheet_name = 'VOC', index = False, header = True)
    pm10.to_excel(writer, sheet_name = 'PM10', index = False, header = True)

    writer.save()