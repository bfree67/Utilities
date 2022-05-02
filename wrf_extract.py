# -*- coding: utf-8 -*-
"""
Created on Sun Apr 26 12:36:37 2020
Extract single file from BZ2 compression and
saves it into a decompressed file

@author: Brian
"""
import bz2
import glob
import os

org_path = os.getcwd()

for file in glob.glob("*.bz2"): 
    
    # make new fiel name by remove .bz2 from end
    file_len = len(file)
    new_name = file[0:file_len-4]

    #decompress data
    print('Decompressing ' + file)
    with bz2.open(file, "rb") as f:
        content = f.read()
    
    #store decompressed data to .dat file
    print('Writing to ' + new_name + '\n')
    with open(new_name, 'wb') as f:
       f.write(content)
       f.close()
