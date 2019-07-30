# -*- coding: utf-8 -*-
"""
Created on Tue Jul  9 15:06:09 2019

Reads excel sheet with facility, year, and annual emissions in TPA
Assumes first year only 1/2 of annual emissions are generated

Creates a 0/1 vector of years then multiplies times the annual amount by pollutant

Create a cumulative tally for each year and compares difference between previous
years to show deltas

Assumes that annual emissions are constant after the first year

@author: Brian
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

df = pd.read_excel("MIC-history.xlsx")
df.sort_values(by=['Year'], ascending = True, inplace=True)
df = df.reset_index()
df = df.drop('index', axis=1)

chem_type = ['NH3', 'NOX', 'PM10', 'SO2', 'VOC']

for chem in chem_type:
    
    cum = np.zeros(2020-1958)
   
    for j in range(19): # add contribution from each site
        years = np.arange(1958, 2020)
        a_year = int(df.Year[j])
        years = (years >= a_year) + 0. # see where site begins
        
        # at first year only do half emissions
        for k in range (len(years)):
            if years[k] == 1.:
                years[k] = 0.5
                break
        #sum each site
        cum += years * df.loc[j,chem]
    
    # calculate annual differences (delta)
    delta = np.zeros(2020-1958)   
    for i in range(len(cum)):
        if i == 0:
            delta[i] = 0
        else:
            delta[i] = round(abs(cum[i] - cum[i-1]),1)
    
    #plot cumulative and delta histories
    years = np.arange(1958, 2020)        
    fig, axs = plt.subplots(2)
    
    fig.suptitle('Growth of annual ' + chem + ' emissions')
    fig.set_figheight(5)
    fig.set_figwidth(8)
    
    # cumulative plot
    axs[0].plot(years, cum)
    axs[0].set_yscale('linear')
    axs[0].set_ylabel('Tonnes (cum.)')
    axs[0].axis([1958, 2020, 100, round(cum.max(),-3) + 2000])
    
    # delta plot
    axs[1].bar(years, delta)
    axs[1].set_yscale('log')
    axs[1].set_ylabel('Tonnes (difference)')
    axs[1].axis([1958, 2020, 1, 10000])
    plt.show()
    fig.savefig(chem + '.svg', format='svg', dpi=1200) ### save figure to file
    
# Create figure and plot a stem plot with the date
built = np.zeros(len(years))
for m in range(len(df)):
    
    for yr in years:
        
        if int(df.Year[m]) == yr:
            built[yr-1958] = 1.

fig = plt.figure()
ax = fig.add_subplot(111)
fig.set_figheight(5)
fig.set_figwidth(8)           
plt.bar(years,built, color = 'red')
plt.ylim(0, 2)
ax.yaxis.set_visible(False)
plt.show()
#fig.savefig('mic_timeline.svg', format='svg', dpi=1200) ### save figure to file
            
            
            
            


