# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 15:07:16 2019

@author: Brian
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans

#grid_file = 'Discrete-Receptors.xls'
grid_file = 'Discrete-Receptors.xls' # outliers (11 & 13) removed

grid_df = pd.read_excel(grid_file, sheet_name='Grid')
receptors_df = pd.read_excel(grid_file, sheet_name='Receptor')

x_grid = np.asmatrix((grid_df.X)).T; y_grid = np.asmatrix((grid_df.Y)).T
x_rec = np.asmatrix((receptors_df.X)).T; y_rec = np.asmatrix((receptors_df.Y)).T

# find UTM distance from grid center to receptor
dist_df = pd.DataFrame(grid_df.LOC_ID)
dist = np.zeros((len(x_grid), len(x_rec))) #make distance matrix
for i in range (len(x_rec)):
    dist[:,i] = (np.sqrt(np.power(x_grid - x_rec[i],2) + np.power(y_grid - y_rec[i],2))).T
# calculate distance matrix    
dist_df = pd.concat([dist_df, pd.DataFrame(dist)], axis=1)


### do kmeans for receptors locations
dataset1 = receptors_df[['X', 'Y']]

# find the appropriate cluster number
plt.figure(figsize=(10, 8))
wcss = []
for i in range(1, 11):
    kmeans = KMeans(n_clusters = i, init = 'k-means++', random_state = 42)
    kmeans.fit(dataset1)
    wcss.append(kmeans.inertia_)
plt.plot(range(1, 11), wcss)
#plt.title('The Elbow Method')
plt.xlabel('Number of clusters (k)')
plt.ylabel('Inertia')
plt.show()

cluster_n = 2

# create kmeans object
kmeans = KMeans(n_clusters= cluster_n)
# fit kmeans object to data
kmeans.fit(dataset1)
# print location of clusters learned by kmeans object
print(kmeans.cluster_centers_)
# save new clusters for chart
y_km = kmeans.fit_predict(dataset1)
cluster_centers = kmeans.cluster_centers_

plt.figure(figsize=(5, 10))
plt.xlim(546, 563) 
plt.scatter(dataset1.X[y_km ==0], dataset1.Y[y_km == 0], s=100, c='red')
plt.scatter(dataset1.X[y_km ==1], dataset1.Y[y_km == 1], s=100, c='black')
plt.scatter(dataset1.X[y_km ==2], dataset1.Y[y_km == 2], s=100, c='blue')
plt.scatter(dataset1.X[y_km ==3], dataset1.Y[y_km == 3], s=100, c='green')
plt.scatter(dataset1.X[y_km ==4], dataset1.Y[y_km == 4], s=100, c='cyan')
plt.scatter(cluster_centers[:,0],cluster_centers[:,1], s=100, c='cyan')
plt.xlabel('X')
plt.ylabel('Y')
plt.show()