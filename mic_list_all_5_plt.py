# -*- coding: utf-8 -*-
"""
Created on Mon 4 Sep 2019

Utility to create Normal and Abnormal emissions rates for individual sources
from all companies in different folders.
Can read multiple files in xlsx format.
All files should be in the same folder as the company name (QAFCO - QAFCO folder)
Assumes all data files have the same column headers
['Source Code','Month', 'SO2', 'NOX', 'VOC', 'PM10', 'NH3', 'H2S', 'HF'] 

Update 5 Sep 2019: Fixed averaging so abnormal = Ave(>=HCI) and Normal = Ave (<HCI)
(HCI is the high CI limit)

Update 8 Sep 2019: Added plot feature back that puts vert lines of averages and Upper CI
Also makes ave & Upper CI a function

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import matplotlib.pyplot as plt
from scipy.stats import sem, t
import time
import numpy as np

version = '2.908.2'
con_int = 0.9 # set confidence interval (90% = 0.9)
units = 'mg/Nm3'

# dictionaries
chem_id = {'SO2':'SO2',
           'NOX':'NOX',
           'VOC':'VOC',
           'PM10':'PM10-PRI',
           'NH3':'NH3',
           'H2S':'7783064',
           'HF':'7664393'}

chem_desc = {'SO2':'Sulfur Dioxide',
             'NOX':'Nitrogen Oxides',
             'VOC':'Volatile Organic Compounds',
             'PM10':'PM10 Primary (Filt + Cond)',
             'NH3':'Ammonia',
             'H2S':'Hydrogen Sulfide',
             'HF':'Hydrogen Fluoride'}

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

def aqmis_col(chem_name, chem_list): # chem_name is a text sting and chem_list is a 2 column list
    
    df = pd.DataFrame(chem_list)
    df.columns = ['Emission Unit ID', 'Emission Value']

    df['Pollutant'] = chem_id[chem_name]
    df['Description'] = chem_desc[chem_name]
    df['Units'] = units
    
    return df

def confidence_int(val):  # input is a dataframe with one numerical column
    ave = float(val.mean(skipna = True))  
    val_ar = np.nan_to_num(val.values) # convert from dataframe to numpy array  
    std_err = sem(val_ar) # take standard error
    h = std_err * t.ppf(con_int, val.count())  # prepare CI using critical t stat      
    return val_ar, ave, ave + h  # returns array of values, average and upper CI 
    
def scenario_dev(chem_name, chem):   # chem_name is a text sting and chem is a dataframe
    
    #t_crit = stats.t.ppf(con_int, df=11) 
    chemical = chem.groupby(['Source Code'])
    
    norm_scenarios = []
    abnorm_scenarios = []
    
    for i, (chemical_name, chemical_gdf) in enumerate(chemical):
        norm = []
        abnorm = []
        chemical_gdf.columns =['a', 'b', 'result']
        val_df = drop_index(chemical_gdf.replace(0., np.NaN))
        val = val_df['result']
        
        val_ar, ave, abnormal_hi = confidence_int(val) 
               
        normal = round(val_ar[val_ar < abnormal_hi].mean(),3) # average values less than the CI
        abnormal = round(val_ar[val_ar >= abnormal_hi].mean(),3) # average values gte the CI
        
        if abnormal != abnormal: # check if it's nan and make normal value
            abnormal = normal
               
        norm.append(chemical_name)
        norm.append(normal)
        
        abnorm.append(chemical_name)
        abnorm.append(abnormal) 
        
        norm_scenarios.append(norm)
        abnorm_scenarios.append(abnorm)

    df_norm = aqmis_col(chem_name, norm_scenarios)
    df_abnorm = aqmis_col(chem_name, abnorm_scenarios)
            
    return df_norm, df_abnorm

def drop_source(df):
    return df.drop('Source Code', axis=1)

def plot_hist(chem_name, df, company):  # chem_name is a text string, df is a dataframe, company is comp string
    chemical = df.groupby(['Source Code'])
    
    check = float(df.mean())
    
    if  check > 0.: # check to see if the df has any values.
        
        plt.figure()
        fig = plt.figure(figsize=(8,16))
        # Iterate through sources
        
        for i, (chemical_name, chemical_gdf) in enumerate(chemical):
            
            val_ar, ave, abnormal_hi = confidence_int(chemical_gdf.iloc[ : , 2 ])
        
            # create subplot axes in a 3x3 grid
            ax = plt.subplot(17, 3, i + 1) # nrows, ncols, axes position
            # plot the continent on these axes
            chemical_gdf.hist(ax=ax)
            
            if ave == abnormal_hi:
                ax.axvline(x= ave, color = 'g')
            else:
                ax.axvline(x= ave, color = 'g')
                ax.axvline(x= abnormal_hi, color = 'r')
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

#****************** set directories ***********************************

# list of companies to check - should also be the folder name
#companies = ['QAFAC', 'QAPCO', 'SEEF', 'NGL', 'Refinery', 'MPCL', 'QSteel', 'QAFCO', 'Qatalum', 'Qchem']
companies = ['Qchem']

for comp in companies:  
    
    working_dir = 'C:\\TARS\\AAAActive\\Qatar Airshed Study Jul 2017\\Compare\\' + comp
    os.chdir(working_dir)    
               
    # **************** make templates **************************************
    
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
    
    print(str(len(file_list)) + ' file(s) imported.')
    
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
    
    # ******************* Make scenarios *****************
    df_scene_norm = pd.DataFrame()
    df_scene_abnorm = pd.DataFrame()
    
    if float(so2.mean()) > 0.: 
        plot_hist('SO2', so2, comp)
        df_so2_norm, df_so2_ab = scenario_dev('SO2', so2)
        df_scene_norm = pd.concat([df_scene_norm, df_so2_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_so2_ab], axis = 0, sort = False)
        
    if float(nox.mean()) > 0.:  
        plot_hist('NOX', nox, comp)
        df_nox_norm, df_nox_ab = scenario_dev('NOX', nox)
        df_scene_norm = pd.concat([df_scene_norm, df_nox_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_nox_ab], axis = 0, sort = False)
        
    if float(voc.mean()) > 0.:    
        plot_hist('VOC', voc, comp)
        df_voc_norm, df_voc_ab = scenario_dev('VOC', voc)
        df_scene_norm = pd.concat([df_scene_norm, df_voc_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_voc_ab], axis = 0, sort = False)
        
    if float(pm10.mean()) > 0.:    
        plot_hist('PM10', pm10, comp)
        df_pm10_norm, df_pm10_ab = scenario_dev('PM10', pm10)
        df_scene_norm = pd.concat([df_scene_norm, df_pm10_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_pm10_ab], axis = 0, sort = False)
        
    if float(nh3.mean()) > 0.:  
        plot_hist('NH3', nh3, comp)
        df_nh3_norm, df_nh3_ab = scenario_dev('NH3', nh3)
        df_scene_norm = pd.concat([df_scene_norm, df_nh3_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_nh3_ab], axis = 0, sort = False)
        
    if float(h2s.mean()) > 0.:    
        plot_hist('H2S', h2s, comp)
        df_h2s_norm, df_h2s_ab = scenario_dev('H2S', h2s)
        df_scene_norm = pd.concat([df_scene_norm, df_h2s_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_h2s_ab], axis = 0, sort = False)
        
    if float(hf.mean()) > 0.:    
        plot_hist('HF', hf, comp)
        df_hf_norm, df_hf_ab = scenario_dev('HF', hf)
        df_scene_norm = pd.concat([df_scene_norm, df_hf_norm], axis = 0, sort = False)
        df_scene_abnorm = pd.concat([df_scene_abnorm, df_hf_ab], axis = 0, sort = False)
    
    df_scene_norm = df_scene_norm.dropna()
    df_scene_abnorm = df_scene_abnorm.dropna()
    
     
    # ************ Save worksheets ***************************
    
    # Save in new AQMIS format as xlsx 
    save_df = input('Save new scenarios for ' + comp +'? (y/n) ')
    
    #new_path = 'C:\TARS\AAALakes\Kuwait AQMIS\'
    if save_df == 'y':   
        os.chdir("\TARS\AAAActive\Qatar Airshed Study Jul 2017\Compare") # reset to new folder path 
        time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
        
        writer = pd.ExcelWriter('Annual_'+comp+'_'+time_stamp+'.xlsx')
        
        # write worksheets. Save w/o index or column header
        df_scene_norm.to_excel(writer, sheet_name = 'Normal', index = False, header = True)
        df_scene_abnorm.to_excel(writer, sheet_name = 'Abnormal', index = False, header = True)
        stats2.to_excel(writer, sheet_name = 'Stats', index = True, header = True)
    
        writer.save()