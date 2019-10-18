# -*- coding: utf-8 -*-
"""
Created on 18 Oct 2019

Utility to reduce sources based on emission contributions. Takes complete Facility Template
and removes unused release points and emission units.

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import time

version = '1.1018'
# global constants for country and company


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
    df3 = drop_index(df3)
    return df3.iloc[0:3]

def new_header2(df):
    #df.columns = df.iloc[0] # make new column headers from first row
    return df.iloc[0:2]
    

# function to reset dataframe index and drop added index column
def drop_index(df):
    df = df.reset_index()
    return df.drop('index', axis=1)

def strip_data(x_string):
    
    new_sheet = new_template.parse(x_string)
    
    # strip header and name columns
    new_sheet.columns = new_sheet.iloc[0]
    return new_sheet.iloc[3:len(new_sheet)]

def reconstruct(new_sheet_hd, new_sheet):
    new_sheet = pd.concat((new_sheet_hd, new_sheet), axis = 0)
    return drop_index(new_sheet)


# ***************** START PROGRAM *********************************

print('*************************************************\n')
print('       AQMIS Sources in Template Utility  ')
print('                Version ' + version + '\n')
print('*************************************************\n')

# ******************* make new headers ****************************

# location of blank template with only column headers    
new_file = 'Facilities_Template.xlsx'    
new_template = pd.ExcelFile(new_file)
new_names = new_template.sheet_names

# get template headers 
new_fac_hd = new_template.parse(new_names[0])
new_rp_hd = new_template.parse(new_names[1])
new_eu_hd = new_template.parse(new_names[2])
new_pro_hd = new_template.parse(new_names[3])
new_apport_hd = new_template.parse(new_names[4])
new_period_hd = new_template.parse(new_names[5])
new_emissions_hd = new_template.parse(new_names[6])

# make new headers 
new_fac_hd = new_header(new_fac_hd)
new_rp_hd = new_header(new_rp_hd)
new_eu_hd = new_header(new_eu_hd)
new_pro_hd = new_header(new_pro_hd)
new_apport_hd = new_header(new_apport_hd)
new_period_hd = new_header(new_period_hd)
new_emissions_hd = new_header(new_emissions_hd)

# ****************************************
# * Facility sheet
new_fac = strip_data('Facilities')

# ****************************************
# * Release Point sheet
new_rp = strip_data('Release Points')

# ****************************************
# * New Emission Units sheet
new_eu = strip_data('Emission Units')

# ****************************************
# * New Processes sheet
new_pro = strip_data('Processes')

# ****************************************
# * New Apportionment sheet
new_apport = strip_data('Apportionment')

# ****************************************
# * New Emission Periods sheet
new_period = strip_data('Emission Periods')

# ****************************************
# * New Emission sheet
new_emissions = strip_data('Emissions')

# set up cumulative columns
es = new_emissions.loc[new_emissions['Pollutant'] == 'Volatile Organic Compounds']
es = drop_index(es.sort_values(by=['Emission'], ascending=False))
es['cum_sum'] = es['Emission'].cumsum()
es['cum_perc'] = 100*es['cum_sum']/es['Emission'].sum()

# select only the top 96% contributors
es['top'] = (es['cum_perc'] < 96.) + 0 #( 1- if less than 96%, 0 if greater))
es1 = es[es.top == 1]  
es1 = es1.drop('top', axis=1)
es1 = es1.drop('cum_perc', axis=1)
es1 = es1.drop('cum_sum', axis=1)

new_es = es1
new_per = pd.DataFrame()
new_app = pd.DataFrame()
new_process = pd.DataFrame()
new_euts = pd.DataFrame()
new_rpts = pd.DataFrame()
new_facilities = pd.DataFrame()

# get list of emission units that contribute the most VOCs
eu_list = list(es1['Unit ID'])

#get other emissions from the same emissions units
esn = new_emissions.loc[new_emissions['Pollutant'] != 'Volatile Organic Compounds']
per = new_period
app = new_apport
pro = new_pro
eu = new_eu
rp = new_rp
facilities = new_fac

for unit in eu_list:
    esn1 = esn.loc[esn['Unit ID'] == unit]
    new_es = pd.concat((new_es, esn1), axis = 0)
    
    per1 = per.loc[per['Unit ID'] == unit]
    new_per = pd.concat((new_per, per1), axis = 0)
    
    app1 = app.loc[app['Unit ID'] == unit]
    new_app = pd.concat((new_app, app1), axis = 0)
    
    pro1 = pro.loc[app['Unit ID'] == unit]
    new_process = pd.concat((new_process, pro1), axis = 0)
    
    eu1 = eu.loc[app['Unit ID'] == unit]
    new_euts = pd.concat((new_euts, eu1), axis = 0)
    
    rp1 = rp.loc[app['Unit ID'] == unit]
    new_rpts = pd.concat((new_rpts, rp1), axis = 0)
      
new_emissions1 = reconstruct(new_emissions_hd, new_es)
new_period1 = reconstruct(new_period_hd, new_per)
new_apport1 = reconstruct(new_apport_hd, new_app)
new_pro1 = reconstruct(new_pro_hd, new_process)
new_eu1 = reconstruct(new_eu_hd, new_euts)
new_rp1 = reconstruct(new_rp_hd, new_rpts)

# find contributing facilties
fac_list = list(new_rpts['Facility ID'].drop_duplicates())

for fac in fac_list:
    
    fac1 = facilities.loc[facilities['Facility ID'] == fac]
    new_facilities = pd.concat((new_facilities, fac1), axis = 0)

new_fac1 = reconstruct(new_fac_hd, new_facilities)    

# **************************************************
# Save in new AQMIS format as xlsx 
save_df = input('\nSave reduced AQMIS sources? (y/n) ')


if save_df == 'y':   
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    
    writer = pd.ExcelWriter('Reduced_format_'+time_stamp+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    new_fac1.to_excel(writer, sheet_name = 'Facilities', index = False, header = False)
    new_rp1.to_excel(writer,sheet_name = 'Release Points', index = False, header = False)
    new_eu1.to_excel(writer,sheet_name = 'Emission Units', index = False, header = False)
    new_pro1.to_excel(writer,sheet_name = 'Processes', index = False, header = False)
    new_apport1.to_excel(writer,sheet_name = 'Apportionment', index = False, header = False)
    new_period1.to_excel(writer,sheet_name = 'Emission Periods', index = False, header = False)
    new_emissions1.to_excel(writer,sheet_name = 'Emissions', index = False, header = False)
    writer.save()
