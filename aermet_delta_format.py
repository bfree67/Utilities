# -*- coding: utf-8 -*-
"""
Created on Mon Nov 11 15:55:17 2019

Opens aermet file, does calculation then reformats it for plotting

@author: Brian
"""

import pandas as pd
from itertools import islice

#open format file or example of aermet
format_file = 'o3.DV.normal.plt'
header = []
with open(format_file) as f:
    for line in islice(f, 8):
        header.append(line)
        print(line)

#open file with new aermet data
file_name = 'ozone_delta.csv'
df = pd.read_csv(file_name)

sp = ' '
with open('output.plt', 'a') as f1:
    for i in range(len(header)):
        f1.write(header[i])
    for i in range(len(df)):
        
        str_df = 5*sp + format(df.iloc[i][0], '.2f') + 4*sp + format(df.iloc[i][1], '.2f') \
                + 6*sp + format(df.iloc[i][2], '.5f') + 4*sp + format(df.iloc[i][3], '.2f') \
                + 4*sp + format(df.iloc[i][4], '.2f') + 5*sp + format(df.iloc[i][5], '.2f') \
              + 4*sp + df.iloc[i][6] + 2*sp + df.iloc[i][7] + 7*sp + str(df.iloc[i][8]) +'\n'

        f1.write(str_df)
f1.close
