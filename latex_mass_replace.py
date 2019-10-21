# -*- coding: utf-8 -*-
"""
Created on Fri Oct  4 16:58:46 2019

Replaces text in all LaTex (.tex) files within a folder

@author: Brian, Python 3.5
"""

import glob
import pandas as pd
import os
from colorama import Fore, Style

current_path = os.getcwd()  # get current working path to save later
file_list = glob.glob("*.tex")

original_text = 'Bijel-1'
new_text = 'Block 19'

print('\nReplacing ' + Fore.RED + original_text + Style.RESET_ALL + ' with ' 
      + Fore.RED + new_text + Style.RESET_ALL)

replace = input('Continue? (y/n) ')

if replace == 'y':

    for file in file_list: 
        
        print('Checking ' + file)
    
        with open(file, 'r') as infile:
            data = infile.read()
            data_n = data.replace(original_text, new_text)
        print('Replaced ' + original_text + ' ' + str(data_n.count(original_text)) + ' times.\n')
            
        with open(file, "w") as outfile:
            outfile.write(data_n)
