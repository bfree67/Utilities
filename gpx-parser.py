# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 12:45:41 2019
Adapted from
https://towardsdatascience.com/how-tracking-apps-analyse-your-gps-data-a-hands-on-tutorial-in-python-756d4db6715d
"""

import gpxpy
import matplotlib.pyplot as plt
import datetime
from geopy import distance
from math import sqrt, floor
import numpy as np
import pandas as pd
import haversine

filename = '20181211-131110'
gpx_file = open(filename+'.gpx', 'r')
gpx = gpxpy.parse(gpx_file)

data = gpx.tracks[0].segments[0].points

## Start Position
start = data[0]
## End Position
finish = data[-1]

df = pd.DataFrame(columns=['lon', 'lat', 'alt', 'time'])
for point in data:
    df = df.append({'lon': point.longitude, 'lat' : point.latitude, 'alt' : point.elevation, 'time' : point.time}, ignore_index=True)

### initial lists
alt_dif = [0]
time_dif = [0]
dist_vin = [0]
dist_hav = [0]
dist_vin_no_alt = [0]
dist_hav_no_alt = [0]
dist_dif_hav_2d = [0]
dist_dif_vin_2d = [0]

for index in range(len(data)):
    if index == 0:
        pass
    else:
        start = data[index-1]
        
        stop = data[index]
        
        distance_vin_2d = distance.vincenty((start.latitude, start.longitude), (stop.latitude, stop.longitude)).m
        dist_dif_vin_2d.append(distance_vin_2d)
        
        distance_hav_2d = haversine.haversine((start.latitude, start.longitude), (stop.latitude, stop.longitude))*1000
        dist_dif_hav_2d.append(distance_hav_2d)
        
        dist_vin_no_alt.append(dist_vin_no_alt[-1] + distance_vin_2d)
        dist_hav_no_alt.append(dist_hav_no_alt[-1] + distance_hav_2d)
        
        alt_d = start.elevation - stop.elevation
        alt_dif.append(alt_d)
        
        distance_vin_3d = sqrt(distance_vin_2d**2 + (alt_d)**2)
        
        distance_hav_3d = sqrt(distance_hav_2d**2 + (alt_d)**2)
                
        time_delta = (stop.time - start.time).total_seconds()
        time_dif.append(time_delta)
                
        dist_vin.append(dist_vin[-1] + distance_vin_3d)
        dist_hav.append(dist_hav[-1] + distance_hav_3d)

##### append to dataframe        
df['dis_vin_2d'] = dist_vin_no_alt 
df['dist_hav_2d'] = dist_hav_no_alt
#df['dis_vin_3d'] = dist_vin
#df['dis_hav_3d'] = dist_hav
df['alt_dif'] = alt_dif
df['time_dif'] = time_dif
df['dis_dif_hav_2d'] = dist_dif_hav_2d
#df['dis_dif_vin_2d'] = dist_dif_vin_2d
df['dist_dif_per_sec'] = df['dis_dif_hav_2d'] / df['time_dif']
df['spd'] = (df['dis_dif_hav_2d'] / df['time_dif']) * 3.6

#print('Vincenty 2D : ', round(dist_vin_no_alt[-1],0))
print('Haversine 2D : ', round(dist_hav_no_alt[-1]),0)
#print('Vincenty 3D : ', round(dist_vin[-1]),0)
#print('Haversine 3D : ', round(dist_hav[-1]),0)
print('Total Time : ', floor(sum(time_dif)/60),' min ', int(sum(time_dif)%60),' sec ')

df_with_timeout = df[df['dist_dif_per_sec'] > 0.9]

df_idle = df[df['dist_dif_per_sec'] <= 0.9]

idle_time = sum(df_idle['time_dif'])

cruise_time = sum(df_with_timeout['time_dif'])

tot_time = sum(df['time_dif'])

avg_km_h = (sum((df_with_timeout['spd'] * 
                 df_with_timeout['time_dif'])) / 
            sum(df_with_timeout['time_dif']))

print('\nAverage cruise speed (km/h) is ', round(avg_km_h,1))
print('% idle time is ', round( (idle_time/tot_time) * 100,0),'%')
print('% cruise time is ', round((cruise_time/tot_time) * 100,0),'%')

### write results to spreadsheet in Excel
filename_out = filename + '_out'
writer = pd.ExcelWriter(filename_out + '.xlsx')
df.to_excel(writer,'GPX')
writer.save()