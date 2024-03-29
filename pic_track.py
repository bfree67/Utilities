# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 19:35:01 2022

Collects GPS Lat/Long coordinates from jpeg photos in a file
Prepares a dataframe and saves it in decimal degrees

@author: Brian
"""
import glob
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
import pandas as pd

def get_exif(filename):
    
    '''
    Extracts exif metadata from jpg
    returns exif dictionary
    
    '''
    exif = Image.open(filename)._getexif()

    if exif is not None:
        for key, value in exif.items():
            name = TAGS.get(key, key)
            exif[name] = exif.pop(key)

        if 'GPSInfo' in exif:
            for key in exif['GPSInfo'].keys():
                name = GPSTAGS.get(key,key)
                exif['GPSInfo'][name] = exif['GPSInfo'].pop(key)
                
    return exif

def get_coords(raw_list):
    
    '''
    converts raw list in deg/min/sec format to decimal degrees
    
    '''
    
    degs = float(raw_list[0])
    mins = float(raw_list[1])
    secs = float(raw_list[2])
    
    
    return round((degs + mins/60 + secs/3600), 6)

# collect all jpeg files in a fodler    
img_files = glob.glob('*.jpg')   

# initialize lists
count = 1
lat_list = []
long_list = []
count_list = []

for image_path in img_files:

    exif = get_exif(image_path)
    
    # get latitude
    lat_raw = list(exif['GPSInfo']['GPSLatitude'])
    lat_ref = exif['GPSInfo']['GPSLatitudeRef']
    lat = get_coords(lat_raw)
    
    # is point is south of the Equator, make it negative
    if lat_ref == 'S':
        lat = -1 * lat()
    else:
        pass
    
    # get longitude
    long_raw = list(exif['GPSInfo'][4])
    long_ref = exif['GPSInfo']['GPSLongitudeRef']
    long = get_coords(long_raw)
    
    # if pt is west of Greenwich, make it negative
    if long_ref == 'W':
        long = -1 * long
    else:
        pass
    
    # update lists
    lat_list.append(lat)
    long_list.append(long)
    count_list.append(count)
    
    count += 1

## put into a dataframe in Excel-KML format (https://www.earthpoint.us/ExcelToKml.aspx)
df_coords = pd.DataFrame({'Latitude':lat_list, 'Longitude':long_list, 'Name':count_list})

# save to excel
df_coords.to_excel('Flightpath.xlsx')