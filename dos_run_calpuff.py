"""
Created on Sat Jan 30 16:15:53 2021
Loads CALPUFF.INP files and runs CALPUFF 7
Changes CONC.DAT output file to new name to prevent overwrite
@author: Brian
"""

import os
import shutil

home_path = os.getcwd()

# location of CALPUFF .exe file
calpuff_path = r"C:\Program Files (x86)\Lakes\CALPUFF View\Models_7"

# CALPUFF model exe file
calpuff = "calpuff_v7.2.1"
calpost = "calpost_v7.1.0"

# location of CALPUFF.INP files
# IMPORTANT!!!! - Update path to WRF data in Subgroup 0a 
calpuff_inp_path = r"C:\TARS\AAAActive\Envirozone\sharjah\sharjah_test"

# IMPORTANT!!!! - Update calpost.inp file path in the pollutant folders
calpost_inp_path = r"C:\TARS\AAAActive\Envirozone\sharjah\sharjah_test\sharjah_test_post\SO2"

#%%

calpuff_scenarios = ['source1'] ## list of CALPUFF.INP file names

calpost_scenarios = ['SO2'] ## make list of pollutants

#%%
## run model
for file_inp in calpuff_scenarios:
    
    cal_name = file_inp + '.INP'
    
    print('\nRunning CALPUFF model for ' + file_inp + '\n')
    
    # Run CALPUFF model
    os.chdir(calpuff_path)
    os.system(calpuff + ' ' + calpuff_inp_path + '\\' + cal_name)
    
    # run CALPOST for each pollutant
    for post_file in calpost_scenarios:
        
        post_inp = post_file + '_C.INP'
        
        print('\nRunning CALPOST for ' + post_file + '\n')
        
        os.chdir(calpuff_path)
        os.system(calpost + ' ' + calpost_inp_path + '\\' + post_inp)

#set back to home directory
os.chdir(home_path)