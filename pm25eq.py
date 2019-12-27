# -*- coding: utf-8 -*-
"""
Created on Tue Dec 24 03:53:38 2019

@author: Brian
"""

import pandas as pd

df = pd.read_excel('pm25eq_norm.xlsx')
df_pol = df[['Unit ID', 'Pollutant','Emission']]

chem_list = ['Sulfur Dioxide', 'Volatile Organic Compounds', 'Nitrogen Oxides',
        'PM10 Primary (Filt + Cond)', 'Ammonia']

unit_id = df_pol['Unit ID'].unique().tolist()

df_group = df_pol.groupby('Unit ID')

a_25 = []

for source in unit_id:
    
    pm25sum = 0
    for chem in chem_list:
        
        if len(df_pol.loc[(df_pol['Unit ID'] == source) & (df_pol['Pollutant'] == chem)]) > 0:
            
            chem_cal = df_pol.loc[(df_pol['Unit ID'] == source) & (df_pol['Pollutant'] == chem)].Emission.values
            
            if chem == 'Sulfur Dioxide':
                pm25sum += chem_cal * 0.298
                
            elif chem == 'Volatile Organic Compounds':
                pm25sum += chem_cal * 0.009
                
            elif chem == 'Nitrogen Oxides':
                pm25sum += chem_cal * 0.067
                
            elif chem == 'PM10 Primary (Filt + Cond)':
                pm25sum += chem_cal * 0.52
                
            elif chem == 'Ammonia':
                pm25sum += chem_cal * 0.194
        
            a = [source, round(pm25sum.item(),6)]
            
    a_25.append(a)
    
df_25 = pd.DataFrame(a_25)
df_25.columns = ['Source', 'PM2.5']

df_25.to_excel('PM25eq_gen_norm.xlsx')


            