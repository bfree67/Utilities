# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 10:03:38 2019

@author: Brian
Original script from https://gist.github.com/nishadhka/3ba801ca980da5b76004631c1935f604
Extract Lat/Long coords from a KML file
"""
from pykml import parser
import pandas as pd
from zipfile import ZipFile

kmz_name = 'mic_intersections.kmz'

### open KMZ file and extract into KML_folder
with ZipFile(kmz_name,'r') as zip_ref:
    zip_ref.extractall('KML_Folder')

##### open KML file in the KML folder - Google will always store as doc.kml
filename= 'KML_folder/doc.kml'

with open(filename) as f:
    folder = parser.parse(f).getroot().Document.Folder

plnm=[]
cordi=[]

for pm in folder.Placemark:

    plnm1=pm.name
    plcs1=pm.Point.coordinates
    plnm.append(plnm1.text)
    cordi.append(plcs1.text)
    
df = pd.DataFrame()
df['place_name'] = plnm
df['coordinates'] = cordi

## unzip lat/long from coordinates
df['Longitude'], df['Latitude'],df['value'] = zip(*df['coordinates'].apply(lambda x: x.split(',', 2)))

### make new dataframe with basic columns
dfa = df.filter(['place_name', 'Longitude', 'Latitude'],axis=1)

### convert objects to numeric format
dfa["Longitude"] = pd.to_numeric(dfa.Longitude, errors='coerce')
dfa["Latitude"] = pd.to_numeric(dfa.Latitude, errors='coerce')

### save to Excel file
dfa.to_excel('intersections.xlsx')

