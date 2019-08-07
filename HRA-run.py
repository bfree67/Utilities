# -*- coding: utf-8 -*-
"""
Created on Tue Aug  6 10:09:12 2019

@author: Brian
"""
#1 Make sure there is 

#2 Assign grid # to sensitive receptor
exec(open('receptors-grid-distance_sort_all_4Aug.py').read())
print('Running receptors-grid-distance_sort_all_4Aug.py to assign sensitive receptors...\n')

#3 Calculate QRA values 
exec(open('country-qra-4Aug.py').read())
print('Running country-qra-4Aug.py to calculate QRA values...\n')

#4 Calculate hotspots and format Conc.dat file
exec(open('hotspot_receptors_4Aug.py').read())
print('Running hotspot_receptors_4Aug.py to calculate country wide hotspots...')