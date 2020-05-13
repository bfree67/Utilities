# -*- coding: utf-8 -*-
"""
Created on Wed May 13 16:23:52 2020

Uses results from prep_merged_data_2.py
@author: Brian
"""

import numpy as np
import matplotlib.pyplot as plt
from pandas import read_csv
import pandas as pd
import math
from keras.models import Sequential
from keras.layers import Dense, Dropout
from keras.layers import LSTM
from keras import metrics
from keras.wrappers.scikit_learn import KerasClassifier
from sklearn.preprocessing import MinMaxScaler 
from sklearn.metrics import mean_squared_error, mean_absolute_error
import time
import sys


def TensorForm(data, look_back):
    # convert an array of values into a dataset tensor
    # ASSUMES data is already in look_back packages
    # Uses results from prep_merged_data_2.py
    
    #determine number of data samples
    rows_data,cols_data = np.shape(data)
    
    #determine # of batches based on look-back size
    tot_batches = int(rows_data/look_back)
    
    #initialize 3D tensor
    threeD = np.zeros(((tot_batches,look_back,cols_data)))
    
    # populate 3D tensor
    for i in range(tot_batches):  
        sample_num = i * look_back # skip by # of look_back
        
        for look_num in range(look_back):
            threeD[i,:,:] = data[sample_num:sample_num+(look_back),:]
    
    return threeD

print('\n****************************************************')
print('              Train LSTM RNNs to impute')
print('                   ver 13May2020_1')
print('****************************************************\n')

# fix random seed for reproducibility
np.random.seed(7)

# 
# ***
# 1) load dataset
# ***
file_in = 'Jahra_Impute_TensorReady_ahead_1_4209.xls'
print('Loading data from ' + file_in)
df_x = pd.read_excel(file_in, sheet_name= 'X_Data')
df_y = pd.read_excel(file_in, sheet_name= 'Y_Data')
df_minmax = pd.read_excel(file_in, sheet_name= 'MinMax')

# ***** prepare scaler inverters for analytes
scalerO3 = MinMaxScaler(feature_range=(0, 1)) ## initialize scaler for O3
scalerNO2 = MinMaxScaler(feature_range=(0, 1)) ## initialize scaler for O3
scalerSO2 = MinMaxScaler(feature_range=(0, 1)) ## initialize scaler for O3
scalerCO = MinMaxScaler(feature_range=(0, 1)) ## initialize scaler for O3
scalerPM10 = MinMaxScaler(feature_range=(0, 1)) ## initialize scaler for O3

scalerO3.fit(np.asmatrix(df_minmax.O3.values))
scalerNO2.fit(np.asmatrix(df_minmax.NO2.values))
scalerSO2.fit(np.asmatrix(df_minmax.SO2.values))
scalerCO.fit(np.asmatrix(df_minmax.CO.values))
scalerPM10.fit(np.asmatrix(df_minmax.PM10.values))

# **** Model input parameters

look_ahead = 1
look_back = look_ahead + 2

pollutants = list(df_y)

pollutant = 'O3'

xdata = TensorForm(df_x.values, look_back)
ydata = np.asmatrix(df_y[pollutant].values).T

# split into train and test sets
train_size = int(len(df_y) * 0.8)
test_size = len(df_y) - train_size
trainX, testX = xdata[0:train_size,:,:], xdata[train_size:len(df_y),:,:]
trainY, testY = ydata[0:train_size], ydata[train_size:len(df_y)]

print('Number of training samples is ' + str(len(trainX)))
print('Number of test samples is ' + str(len(testX)))

n_epochs = 30
n_batch = 64

print('\nBuilding the model...')
    
# ***
# 3) Build the RNN model
# ***

input_nodes = int(trainX.shape[2] * 2)

model = Sequential()

# LSTM layer
model.add(LSTM(input_nodes, activation='relu', recurrent_activation='relu', 
                input_shape=(trainX.shape[1], trainX.shape[2])))

# 1 neuron on the output layer
model.add(Dense(1, activation='tanh'))

# compiles the model
model.compile(loss='mae', 
              optimizer='adam', 
              metrics = ['accuracy'])

# ***
# 4) Increased the batch_size to 72. This improves training performance by more than 50 times
# and loses no accuracy (batch_size does not modify the final result, only how memory is handled)
# ***
#### Start the clock
start1 = time.perf_counter()  

history = model.fit(trainX, trainY, epochs=n_epochs, 
                    batch_size=n_batch, 
                    validation_data=(testX, testY), 
                    shuffle=False)

    # stop clock
end1 = time.perf_counter() 

if (end1-start1 > 60):
    print ("\nModel trained in {0:.1f} minutes".format((end1-start1)/60.))
else:
    print ("\nModel trained in {0:.1f} seconds".format((end1-start1)/1.))

# ***
# 5) test loss and training loss graph. It can help understand the optimal epochs size and if the model
# is overfitting or underfitting.
# ***
xhistory = len(history.history['loss'])
xlin = range(1,xhistory+1)
plt.close('all')
plt.plot(xlin,history.history['loss'],color="black", label='Train')
plt.plot(xlin,history.history['val_loss'], color = "black", linestyle = ':', label='Validate')
plt.xlabel('Epochs')
plt.ylabel('MSE Loss')
plt.xticks(np.arange(min(xlin), max(xlin)+1, (max(xlin)+1 - min(xlin))/10))
plt.legend()
plt.show()

# *****
# Impute
# *****

file_load = 'Jahra_processed_raw_data_7618.xls'
df_ams = pd.read_excel(file_load)

if 'Date' in df_ams.columns:
    df_ams_1= df_ams.drop('Date', axis = 1)  # remove column 'rain' if present

data_gap = df_ams_1[pollutant].values
gaps = np.argwhere(np.isnan(data_gap))

for i in gaps:
    i = int(i)
    df_xp = df_ams_1.iloc[i-look_back:i]  #get look_back rows to predict
    xp = df_xp.values
    yp = model.predict(xp)


