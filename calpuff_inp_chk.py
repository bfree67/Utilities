# -*- coding: utf-8 -*-
"""
Created on Mon Nov 11 15:55:17 2019

Opens CALPUFF.INP file and counts # of NOX and VOC sources
Assumes it is set up for SCICHEM

@author: Brian
"""

import pandas as pd
from itertools import islice

#open format file or example of aermet
format_file = 'SCICHEM_ABNORMAL2.INP'
src_lines = []
sources = 0

NOX = 0
VOC = 0

#read inp file
lines = [line.rstrip('\n') for line in open(format_file)]

#make a smaller list
for i in range(len(lines)):
    
    if lines[i] == 'Subgroup (13b)':
        sources = 1
    
    if lines[i] == 'Subgroup (13c)':
        sources = 0
        
    if sources == 1:
        src_lines.append(lines[i])

# search for only input lines and count emissions
for i in range(len(src_lines)):
    a =src_lines[i].split()
    
    if len(a) > 2:
        if a[2] == 'X':
            if a[12] != '0.00':
                NOX += 1
            if a[13] != '0.00':
                VOC += 1
                
print('NOX sources = ', NOX, 'VOC sources =', VOC)
        




