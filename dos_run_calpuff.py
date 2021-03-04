# -*- coding: utf-8 -*-
"""
Created on Sat Jan 30 16:15:53 2021

Loads CALPUFF.INP files and runs CALPUFF 7
Changes CONC.DAT output file to new name to prevent overwrite

@author: Brian
"""

import os
import shutil

# location of CALPUFF .exe file
calpuff_path = r"C:\Program Files (x86)\Lakes\CALPUFF View\Models_7"

# CALPUFF model exe file
calpuff = "calpuff_v7.2.1"

# location of CALPUFF.INP files
# IMPORTANT!!!! - Update path to WRF data in Subgroup 0a 
calpuff_inp_path = r"C:\Users\Brian\Documents\_Active\Hexagon\pre-incident\wrf\pre-external"

# location of CALPUFF CONC.DAT output files (need to run CALPOST for data)
post_path = r"C:\Users\Brian\AppData\Local\VirtualStore\Program Files (x86)\Lakes\CALPUFF View\Models_7"

# list of CALPUFF.INP file names
calpuff_scenarios = ['SOURCE_3',
'SOURCE_4',
'SOURCE_5',
'SOURCE_6',
'SOURCE_7',
'SOURCE_8',
'SOURCE_9',
'SOURCE_10',
'SOURCE_11',
'SOURCE_12']

for file_inp in calpuff_scenarios:
    
    cal_name = file_inp + '.INP'
    
    print('\nRunning CALPUFF model for ' + file_inp + '\n')
    os.chdir(calpuff_path)
    os.system(calpuff + ' ' + calpuff_inp_path + '\\' + cal_name)
    
    #rename CONC.DAT file
    os.chdir(post_path)
    shutil.copy('CONC.DAT', 'CONC_' + file_inp + '.DAT')



