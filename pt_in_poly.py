# -*- coding: utf-8 -*-
"""
Created on Sat Nov 20 22:49:32 2021

Determines whether a pt is within a polygon using residual signs
based on C172 COG chart at:
http://aviatech.co.za/wp-content/uploads/2014/03/mb_7.jpg

@author: Brian
"""
import numpy as np
from sklearn.linear_model import LinearRegression

## input test coordinates
x_in = 70 # Loaded acft Moment/1000 (lb-in)
y_in = 1600

## define envelope polygon ---- select type (utility or normal)
model_type = 'utility'

utility = np.matrix([[52, 1500],
                     [68, 1960],
                     [71, 2000],
                     [81, 2000],
                     [60, 1500]])

normal = np.matrix([[60,1500],
                    [81,2000],
                    [71,2000],
                    [88,2300],
                    [109,2300],
                    [70, 1500]])

## select envelope baed on model coordinates
if model_type == 'utility':
    envelope = utility
else:
    envelope =  normal

## define regression model
model = LinearRegression()

## make an index to pull from (0, 1, 2,...)
len_r = list(range(len(envelope)))

linears = []

## make combinations of points to fit lines
for i in len_r:
    x1 = envelope[i,0]
    y1 = envelope[i,1]
    x2 = envelope[i-1,0]
    y2 = envelope[i-1,1]
    
    x = np.array([x1,x2]).reshape(-1, 1)
    y = np.array([y1,y2]).reshape(-1, 1)

    model = LinearRegression().fit(x, y)
    params = [model.coef_[0,0], model.intercept_[0]] 

    linears.append(params)

## make matrix of slopes/y-int in form y = ax + b
coefs = np.matrix(linears)

a = coefs[:,0].A1 # slope
b = coefs[:,1].A1 # y-int

## calc new y based on regression values        
y = a*x_in + b

## calc residuals
y_residual = y - y_in

## find signs of residuals (1 pos and 0 neg)
yr_signs = (y_residual > 0.) + 0

## model residuals sign for a point to be with the polygon 
utility_model = np.matrix([0,1,1,1,0])

normal_model = np.matrix([[0,0,0,1,1,0],
                          [0,1,0,1,1,0],
                          [0,1,1,1,1,0]])

## select residual model based on model type
if model_type == 'utility':
    residuals_chk = utility_model
else:
    residuals_chk = normal_model

## if chk = 1, then in envelope
chk = 0

## loop through allowable model profiles
for arr in residuals_chk:
    residuals_in = np.array(arr.A1)

    if (yr_signs == residuals_in).sum() == len(a):
        chk +=1

## if chk = 1, then in envelope        
if chk > 0:
    print('In {} envelope'.format(model_type))
else:
    print('Not in {} envelope'.format(model_type))

print(yr_signs)

