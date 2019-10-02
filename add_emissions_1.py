# -*- coding: utf-8 -*-
"""
Created on Wed 2 October 2019

Creates PM2.5 by backcalculating throughput from NOx emissions using
Emission factors for NOX from AP-42 averages and then used AP-42 value
for PM2.5

Also calculates PM2.5 from PM10 values assuming 54% from AP-42

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import time

version = '1.0'

current_path = os.getcwd()  # get current working path to save later

# function to add proper header in new template format after data has been remapped
def new_header(df):
    hd_list = list(df) # get column headers (name, unknown,,,,,)
    for i in range(1,len(hd_list)):
        hd_list[i] = '' # make blank
    
    df.columns = df.iloc[0] # make new column headers from first row
    df2 = pd.DataFrame(hd_list).T
    df2.columns = df.iloc[0]
    df3 = pd.concat((df2,df), axis = 0)
    df3 = df3.reset_index()
    return df3.drop('index', axis=1)

# function to reset dataframe index and drop added index column
def drop_index(df):
    df = df.reset_index()
    return df.drop('index', axis=1)


# ***************** START PROGRAM *********************************

print('*************************************************\n')
print('       AQMIS Add Emission Utility  ')
print('                Version ' + version + '\n')
print('*************************************************\n')

# **************** make templates **************************************

txtfiles = []
agg = pd.DataFrame

# merge the excel files into one file by reading all excel files in the folder
print('Importing data files...')
file_list = glob.glob("*.xlsx")

file = file_list[0]  #use only first file in the list - ignores open files
txtfiles.append(file)

new_template = pd.ExcelFile(file) # read the indvidual excel file
res = len(new_template.sheet_names) # count the number of sheets in the file
new_names = new_template.sheet_names

# assign sheets to dataframes
new_fac = new_header(new_template.parse(new_names[0]))
new_rp = new_header(new_template.parse(new_names[1]))
new_eu = new_header(new_template.parse(new_names[2]))
new_pro = new_header(new_template.parse(new_names[3]))
new_apport = new_header(new_template.parse(new_names[4]))
new_period = new_header(new_template.parse(new_names[5]))
new_emissions = new_header(new_template.parse(new_names[6]))

print(file)
print(str(len(txtfiles)) + ' imported. \n')

#####################################################
#### first find NOX to make PM2.5 based on combustion
#####################################################
# find specific emission (make sure it's spelled right!)
emission_1 = 'Nitrogen Oxides'
emission_new = 'PM2.5 Primary (Filt + Cond)'

df_emissions = drop_index(new_emissions.loc[new_emissions['Pollutant'] == emission_1])
df_emissions['Unit ID'] = df_emissions['Unit ID'].astype(str)
df_emissions = drop_index(df_emissions.sort_values('Unit ID'))

df_new = df_emissions
df_new['Emission_lb'] = df_new['Emission'] * 2204.62

df_new['Pollutant'] = emission_new

df_period = drop_index(new_period.iloc[3:])
df_period['Unit ID'] = df_period['Unit ID'].astype(str)
df_period = drop_index(df_period.sort_values('Unit ID'))

# find type of fuel
print('\nUpdating emissions...')
df_new['Emission_new'] = 0
for i in range(len(df_new)):
    
    fuel_row = df_period.loc[df_period['Unit ID'] == df_new['Unit ID'].iloc[i]]
    fuel_type = fuel_row.iloc[0]['Material']
    #print(df_new['Unit ID'].iloc[i], fuel_type)
    
    if fuel_type == 'Natural Gas':
        NOX_ef = 232 #lb/MMCF
        pm25_ef = 7 #lb/MMCF
            
    if fuel_type == 'Process Gas' or fuel_type == 'Refinery Gas':
        NOX_ef = 218 #lb/MMCF
        pm25_ef = 7 #lb/MMCF
        
    if fuel_type == 'Crude Oil':
        NOX_ef = 42 #lb/1000 Gallons
        pm25_ef = 6.1 #lb/1000 Gallons
        
    df_new['Emission_new'].iloc[i] = round((df_new['Emission_lb'].iloc[i] / NOX_ef) * pm25_ef / 2204.62,3)
    
df_emissions['Emission'] = df_new['Emission_new']

# drop columns 
df_emissions = df_emissions.drop('Emission_lb', 1)   
df_emissions = df_emissions.drop('Emission_new', 1)       

############# add new NOX derived pm2.5 emissions to emission list
new_emissions = pd.concat((new_emissions, df_emissions), axis = 0)
new_emissions = drop_index(new_emissions)

##### Find PM10 and calculated PM2.5 as 54% of the total
emissions_2 = 'PM10 Primary (Filt + Cond)'
df_pm2 = drop_index(new_emissions.loc[new_emissions['Pollutant'] == emissions_2]) 
df_pm2['Emission'] = df_pm2['Emission'] * 0.54
df_pm2['Pollutant'] = emission_new

############# add newPM10 derived pm2.5 emissions to emission list
new_emissions = pd.concat((new_emissions, df_pm2), axis = 0)
new_emissions = drop_index(new_emissions)

# **************************************************
# Save in new AQMIS format as xlsx 
save_df = input('\nSave new AQMIS Template for ' + emission_new +'? (y/n) ')

#new_path = 'C:\TARS\AAALakes\Kuwait AQMIS\'
if save_df == 'y':   
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    
    writer = pd.ExcelWriter('Emissions_' + emission_new + '_' +time_stamp+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    new_fac.to_excel(writer, sheet_name = 'Facilities', index = False, header = False)
    new_rp.to_excel(writer,sheet_name = 'Release Points', index = False, header = False)
    new_eu.to_excel(writer,sheet_name = 'Emission Units', index = False, header = False)
    new_pro.to_excel(writer,sheet_name = 'Processes', index = False, header = False)
    new_apport.to_excel(writer,sheet_name = 'Apportionment', index = False, header = False)
    new_period.to_excel(writer,sheet_name = 'Emission Periods', index = False, header = False)
    new_emissions.to_excel(writer,sheet_name = 'Emissions', index = False, header = False)
    writer.save()

