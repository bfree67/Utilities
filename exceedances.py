# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 17:57:24 2019
Lists MIC/MME exceedance levels

@author: Brian
"""

nh3 = 'NH3'; hf = '7664393'; h2s = '7783064'; nox = 'NOX'; pm10 = 'PM10-PRI';
so2 = 'SO2'; voc = 'VOC';

hr1 = '1HR'; hr24 = '24HR'; hr8760 = '8760HR'

def mic_exceedance(chem,twa):
    
    thresh = 1.  # default value
    
    if chem == nh3:
        
        if twa == hr1:  
            thresh = 1800. 
    
    if chem == hf:
        
        if twa == hr1:
            thresh = 14.
        if twa == hr24:
            thresh = 4.6
            
    if chem == h2s:
        
        if twa == hr1:
            thresh = 40.
        if twa == hr24:
            thresh = 20.
    
    if chem == nox:
        
        if twa == hr1:
            thresh = 400.
        if twa == hr24:
            thresh = 150.
        if twa == hr8760:
            thresh = 50.           
            
    if chem == so2:
        
        if twa == hr1:
            thresh = 1300.
        if twa == hr24:
            thresh = 364.
        if twa == hr8760:
            thresh = 80.   

    if chem == pm10:
        
        if twa == hr24:
            thresh = 150.
        if twa == hr8760:
            thresh = 50.

    return thresh             

#print (mic_exceedance('NH3', '1HR'))
