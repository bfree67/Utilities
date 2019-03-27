# -*- coding: utf-8 -*-
"""
Created on Thu Mar 21 05:43:20 2019

@author: Brian
Converts WRF data into DNDC formatted data
Julian date, Max Day Temp (oC), Min Day Temp (oC), Precip (cm), and Wind speed (m/s)
Saves in tab separated file with .DAY extension

Initially reads an excel file with the WRF data and then breaks it into years

"""

import pandas as pd

def dndc_weather(opyear):
    #opyear is a dataframe of the year of interest
    
    Year = opyear['YEAR'].max()  ## extract year for filenaming
    
    MinT = (pd.DataFrame(opyear.groupby(['MONTH','DAY'])['T'].min())).reset_index()
    MaxT = (pd.DataFrame(opyear.groupby(['MONTH','DAY'])['T'].max())).reset_index()
    WS = (pd.DataFrame(opyear.groupby(['MONTH','DAY'])['WS'].mean())).reset_index()
    
    Jul = pd.DataFrame(pd.Series(range(1,len(MaxT)+1)))   ## create a df named Jul
    Jul.columns = ['Jday']
    Jul['MaxT'] = MaxT['T']
    Jul['MinT'] = MinT['T']
    Jul['Prec'] = 0
    Jul['WindSpeed'] = WS['WS'].round(1)
    file_name = 'ARD_DNDC_WX_'+str(Year)+'.DAY'
    Jul.to_csv(file_name, encoding='utf-8', sep='\t', header=False, index=False)
    
    return Jul

#####################################################################
######## BEGIN PROGRAM
    
### read input file
df = pd.read_excel('ARD_2016-2018.xlsx')

yr2016 = df.loc[df['YEAR']==2016]
yr2017 = df.loc[df['YEAR']==2017]
yr2018 = df.loc[df['YEAR']==2018]

yr2016_dndc = dndc_weather(yr2016)
yr2017_dndc = dndc_weather(yr2017)
yr2018_dndc = dndc_weather(yr2018)

