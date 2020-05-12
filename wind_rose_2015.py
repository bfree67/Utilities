# -*- coding: utf-8 -*-
"""
Created on Tue May 12 10:57:50 2020

@author: Brian
"""


from windrose import WindroseAxes
from matplotlib import pyplot as plt
import matplotlib.cm as cm
import numpy as np
import pandas as pd

file = '2015_wrf.xlsx'
df = pd.read_excel(file)

ws = df['WS'].values
wd = df['WD'].values

ws = df['WS_WRF'].values
wd = df['WD_WRF'].values

## make wind rose
ax = WindroseAxes.from_ax()
ax.bar(wd, ws, normed=True, opening=0.8, edgecolor='white')
ax.set_legend()

# frequency table
ax.bar(wd, ws, normed=True, nsector=16)
table = ax._info['table']
wd_freq = np.sum(table, axis=0)

direction = ax._info['dir']
wd_freq = np.sum(table, axis=0)
plt.bar(np.arange(16), wd_freq, align='center')

'''
xlabels = ('N','','N-E','','E','','S-E','','S','','S-O','','O','','N-O','')
xticks=arange(16)
gca().set_xticks(xticks)
draw()
gca().set_xticklabels(xlabels)
draw()
'''