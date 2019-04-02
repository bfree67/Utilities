# -*- coding: utf-8 -*-
"""
Created on Fri Mar 22 16:29:21 2019

Converts coordinates in Degree Min Sec to Decimal degrees
https://stackoverflow.com/questions/33997361/how-to-convert-degree-minute-second-to-degree-decimal-in-python
Format requires all quotation marks to be a single apostrophe and not quotes

Correct: 78°55'44.33324'N ---- not Correct: N78°55"44.33324'

@author: Brian
"""

import re
import pandas as pd

def dms2dd(degrees, minutes, seconds, direction):
    dd = float(degrees) + float(minutes)/60 + float(seconds)/(60*60);
    if direction == 'E' or direction == 'N':
        dd *= -1
    return dd;

def dd2dms(deg):
    d = int(deg)
    md = abs(deg - d) * 60
    m = int(md)
    sd = (md - m) * 60
    return [d, m, sd]

def parse_dms(dms):
    parts = re.split('[^\d\w]+', dms)
    lat = dms2dd(parts[0], parts[1], parts[2], parts[3])

    return (lat)

# put in format 78°55'44.33324'N with no spaces

### read input file
df = pd.read_excel('dms.xlsx')

n = len(df)

dd = []

for i in range(n):
    Lat_d = round(parse_dms(df.Lat[i].replace(" ", "")),4)
    Long_d = round(parse_dms(df.Long[i].replace(" ", "")),4)
    a = [Lat_d, Long_d]
    print(Lat_d, Long_d)
    dd.append(a)
    
df_d = pd.DataFrame(dd)
df_d.columns = ['Lat', 'Long']

file_name = 'dec_deg.xlsx'
df_d.to_excel(file_name)