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
from scipy import stats
import time

version = '1.903'

current_path = os.getcwd()  # get current working path to save later

# function to reset dataframe index and drop added index column
def drop_index(df):
    df = df.reset_index()
    return df.drop('index', axis=1)

def make_stats(so2):
    stats1 = pd.DataFrame()
    stats1 = so2.groupby(['Source Code']).sum()
    stats1 = pd.concat([stats1, so2.groupby(['Source Code']).mean()], axis = 1, sort = False)
    stats1 = pd.concat([stats1, so2.groupby(['Source Code']).std()], axis = 1, sort = False)
    stats1 = pd.concat([stats1, so2.groupby(['Source Code']).max()], axis = 1, sort = False)
    stats1 = stats1.drop(columns=['Month'])
    stats1 = pd.concat([stats1, so2.groupby(['Source Code']).min()], axis = 1, sort = False)
    stats1 = stats1.drop(columns=['Month'])
    stats90 = so2.groupby(['Source Code']).mean() + (so2.groupby(['Source Code']).mean() 
                                + (so2.groupby(['Source Code']).std() * 1.363 / 3.464))
    return pd.concat([stats1, stats90], axis = 1, sort = False)
    
def scenario_dev(chem):
    
    t_crit = stats.t.ppf(.9,df=11) # set confidence interval (90% = 0.9)
    chemical = chem.groupby(['Source Code'])
    scenarios = []
    for i, (chemical_name, chemical_gdf) in enumerate(chemical):
        a = []
        chemical_gdf.columns =['a', 'b', 'c']
        normal = float(round(chemical_gdf.mean(),3))
        
        std_chem = float(chemical_gdf.std())
        
        t_var = float((2*normal) + (std_chem * t_crit/3.464))
        
        abnormal = round((chemical_gdf.loc[(chemical_gdf['c'] > t_var)]).mean(),3)
               
        a.append(chemical_name)
        a.append(normal)
        a.append(abnormal.c)
        
        scenarios.append(a)
    df_scene = pd.DataFrame(scenarios)
    df_scene.columns = ['Source Code', 'Normal', 'Abnormal']
        
    return df_scene

def drop_source(df):
    return df.drop('Source Code', axis=1)

def plot_hist(chem_name, df):  # chem_name is a text string and df is a dataframe
    chemical = df.groupby(['Source Code'])
    
    check = float(df.mean())
    
    if  check > 0.: # check to see if the df has any values.
        
        plt.figure()
        fig = plt.figure(figsize=(8,16))
        # Iterate through sources
        
        for i, (chemical_name, chemical_gdf) in enumerate(chemical):
        
            # create subplot axes in a 3x3 grid
            ax = plt.subplot(17, 3, i + 1) # nrows, ncols, axes position
            # plot the continent on these axes
            chemical_gdf.hist(ax=ax)
            # set the title
            ax.set_title(chemical_name)
        
        plt.tight_layout()
        plt.savefig(company + '_' + chem_name + '.png')
        
        print(company + ' histograms plotted for ' + chem_name)
    return 
    
# ***************** START PROGRAM *********************************

print('*************************************************\n')
print('       AQMIS Emission Scenario Utility  ')
print('                Version ' + version + '\n')
print('*************************************************\n')
    
# **************** make templates **************************************

company = 'QSteel'
txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

file_list = glob.glob("*.xlsx")
# merge the excel files into one file by reading all excel files in the folder
print('Importing file ...\n')

for file in file_list:
    txtfiles.append(file)
    
    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    xl_names = xl.sheet_names
    
    if i == 0:
        df_fac = xl.parse(xl_names[0]).fillna(0)


        i += 1     
    
    else:
        df_fac = pd.concat([df_fac, xl.parse(xl_names[0]).fillna(0)], sort = False)
    
    print(file)

df_fac.columns = ['Source Code','Month', 'SO2', 'NOX', 'VOC', 'PM10', 'NH3', 'H2S', 'HF']        

print(str(len(file_list)) + ' file(s) imported.\n')

df_fac = df_fac.sort_values(['Source Code', 'Month'])
so2 = df_fac[['Source Code', 'Month', 'SO2']]
nox = df_fac[['Source Code', 'Month', 'NOX']]
voc = df_fac[['Source Code', 'Month', 'VOC']]
pm10 = df_fac[['Source Code', 'Month', 'PM10']]
nh3 = df_fac[['Source Code', 'Month', 'NH3']]
h2s = df_fac[['Source Code', 'Month', 'H2S']]
hf = df_fac[['Source Code', 'Month', 'HF']]

# make stats worksheet

stats2 = make_stats(so2)
stats2 = pd.concat([stats2, make_stats(nox)], axis = 1, sort = False)
stats2 = pd.concat([stats2, make_stats(voc)], axis = 1, sort = False)
stats2 = pd.concat([stats2, make_stats(pm10)], axis = 1, sort = False)
stats2 = pd.concat([stats2, make_stats(nh3)], axis = 1, sort = False)
stats2 = pd.concat([stats2, make_stats(h2s)], axis = 1, sort = False)
stats2 = pd.concat([stats2, make_stats(hf)], axis = 1, sort = False)

stats2.columns = ['SO2 Sum', 'SO2 Ave', 'SO2 Std', 'SO2 Max', 'SO2 Min', 'SO2 CI', 
		'NOX Sum', 'NOX Ave', 'NOX Std', 'NOX Max', 'NOX Min', 'NOX CI',
		'VOC Sum', 'VOC Ave', 'VOC Std', 'VOC Max', 'VOC Min', 'VOC CI',
		'PM10 Sum', 'PM10 Ave', 'PM10 Std', 'PM10 Max', 'PM10 Min', 'PM10 CI', 
        'NH3 Sum', 'NH3 Ave', 'NH3 Std', 'NH3 Max', 'NH3 Min', 'NH3 CI',
        'H2S Sum', 'H2S Ave', 'H2S Std', 'H2S Max', 'H2S Min', 'H2S CI',
        'HF Sum', 'HF Ave', 'HF Std', 'HF Max', 'HF Min', 'HF CI']

# *********** plot histograms ***************************************

plot_question = input('\nPlot histograms for ' + company +'sources? (y/n) ')

if plot_question == 'y': 
    plot_hist('SO2', so2)
    plot_hist('NOX', nox)
    plot_hist('VOC', voc)
    plot_hist('PM10', pm10)
    plot_hist('MH3', nh3)
    plot_hist('H2S', h2s)
    plot_hist('HF', hf)

# ******************* Make scenarios *****************
    
df_scene = scenario_dev(so2)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(nox))], axis = 1, sort = False)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(voc))], axis = 1, sort = False)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(pm10))], axis = 1, sort = False)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(nh3))], axis = 1, sort = False)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(h2s))], axis = 1, sort = False)
df_scene = pd.concat([df_scene, drop_source(scenario_dev(hf))], axis = 1, sort = False)

df_scene.columns = ['Source Code', 'SO2 Normal', 'SO2 Abnormal',
					'NOX Normal', 'NOX Abnormal',
					'VOC Normal', 'VOC Abnormal',
					'PM 10 Normal', 'PM10 Abnormal',
					'NH3 Normal', 'NH3 Abnormal',
					'H2S Normal', 'H2S Abnormal',
					'HF Normal', 'HF Abnormal']
 
# ************ Save worksheets ***************************

# Save in new AQMIS format as xlsx 
save_df = input('\nSave new sheet for ' + company +'? (y/n) ')

#new_path = 'C:\TARS\AAALakes\Kuwait AQMIS\'
if save_df == 'y':   
    os.chdir("\TARS\AAAActive\Qatar Airshed Study Jul 2017\Compare") # reset to new folder path 
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    
    writer = pd.ExcelWriter('Annual_'+company+'_'+time_stamp+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    df_scene.to_excel(writer, sheet_name = 'Scenarios', index = True, header = True)
    stats2.to_excel(writer, sheet_name = 'Stats', index = True, header = True)
    df_fac.to_excel(writer, sheet_name = 'All', index = False, header = True)
    so2.to_excel(writer, sheet_name = 'SO2', index = False, header = True)
    nox.to_excel(writer, sheet_name = 'NOx', index = False, header = True)
    voc.to_excel(writer, sheet_name = 'VOC', index = False, header = True)
    pm10.to_excel(writer, sheet_name = 'PM10', index = False, header = True)
    nh3.to_excel(writer, sheet_name = 'NH3', index = False, header = True)
    h2s.to_excel(writer, sheet_name = 'H2S', index = False, header = True)
    hf.to_excel(writer, sheet_name = 'HF', index = False, header = True)

    writer.save()