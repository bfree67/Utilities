# -*- coding: utf-8 -*-
"""
Created on Wed Aug 21 14:53:06 2019

Utility to remap AQMIS templates using old import format to new import format.
Requires all old template spreadsheets to be in the same folder using *.xlsx format
and a modified blank template with only column headers to in an other folder.

Updated 26 Aug 2019 - fixed AQZ merging and made into a function add_aqz
Reads column headers of first old template and changes headers of all
other files to ensure standardized column headers. Just have to update the
first file.

Updated 28 Aug 2019 - Added check to see if column headers in Release Point
are in degrees or deg and exit temp is K or C. Added check to see if there is an
active or inactive added column

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os
import time

version = '2.912'
# global constants for country and company
country = 'Kuwait'
company = 'KNPC'
start_date = '01/01/2016'
end_date = '12/31/2016'

current_path = os.getcwd()  # get current working path to save later

# function to merge AQZ with facility in other worksheets
def add_aqz(x):
    # Merge AQZ from new_fac
    df_1 = x[['Facility Name']]
    df_aqz = pd.merge(df_1, df_zones, right_on='Facility Name', left_on='Facility Name')

    x = x.reset_index()
    new_aqz = x.drop('index', axis=1)

    new_aqz['Air Quality Zone'] = df_aqz['AQZ']
    
    return new_aqz

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

# function to check the release point sheet in the old format for common headers
def check_rp(df_rp):
        if 'Latitude [decimal degrees]' in df_rp:
            df_rp = df_rp.rename(columns = {'Latitude [decimal degrees]':'Latitude [decimal deg]'})
        
        if 'Longitude [decimal degrees]' in df_rp:
            df_rp = df_rp.rename(columns = {'Longitude [decimal degrees]':'Longitude [decimal deg]'})
            
        if 'Exit Temperature [K]' in df_rp:
            df_rp = df_rp.rename(columns = {'Exit Temperature [K]':'Exit Temperature [C]'})
        
        return df_rp

# function to check if column exists in old template    
def col_check(old_df, col_ck, new_df):
    if col_ck in old_df.columns:
        new_df[col_ck] = old_df[col_ck]    
    else:
        new_df[col_ck] = ''
    return new_df[col_ck]

# ***************** START PROGRAM *********************************

print('*************************************************\n')
print('       AQMIS Old to New Template Utility  ')
print('                Version ' + version + '\n')
print('*************************************************\n')

# ******************* make new headers ****************************

# location of blank template with only column headers    
new_file = 'C:\\TARS\\AAALakes\\Kuwait AQMIS\\Blank_Template.xlsx'    
new_template = pd.ExcelFile(new_file)
new_names = new_template.sheet_names

# get headers from blank template
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

# **************** make templates **************************************

txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

# merge the excel files into one file by reading all excel files in the folder
print('Importing data files...')
for file in glob.glob("*.xlsx"):
    txtfiles.append(file)
    
    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    xl_names = xl.sheet_names
    
    if i == 0:
        df_fac = xl.parse('Facility Information')
        df_rp = check_rp(xl.parse('Release Points')) # uses check_rp function to standardize headers
        df_source = xl.parse('Source Information')     
        df_emissions = xl.parse('Emissions')

        i += 1     
    
    else:
        df_fac = pd.concat([df_fac, xl.parse('Facility Information')], sort = False)    
        df_rp = pd.concat([df_rp, check_rp(xl.parse('Release Points'))], sort = False)        
        df_source = pd.concat([df_source, xl.parse('Source Information')], sort = False)      
        df_emissions = pd.concat([df_emissions, xl.parse('Emissions')], sort = False)

    
    print(file)
print(str(len(txtfiles)) + ' imported. \n')

# make sub dataframe with AQZ
df_zones = df_fac[['Facility Name', 'AQZ']]
df_zones = drop_index(df_zones)

# location of blank template with only column headers    
# use new_file from before   
new_template2 = pd.ExcelFile(new_file)
new_names = new_template2.sheet_names

# ****************************************
# * New Facility sheet
new_fac = new_template2.parse('Facilities')
new_fac.columns = new_fac.iloc[0]
new_fac = new_fac.iloc[0:0]

new_fac['Facility Name'] = df_fac['Facility Name']
new_fac['Country'] = 'Kuwait'
new_fac['Air Quality Zone'] = df_fac['AQZ']
new_fac['Company Name'] = 'MEW'
new_fac['Latitude [deg]'] = df_fac['Latitude']
new_fac['Longitude [deg]'] = df_fac['Longitude']
new_fac['Governorate'] = df_fac['Location']
new_fac['SIC'] = df_fac['SIC']

new_fac = pd.concat((new_fac_hd, new_fac), axis = 0)
new_fac = drop_index(new_fac)

# ****************************************
# * New Release Points sheet
new_rp = new_template2.parse('Release Points')
new_rp.columns = new_rp.iloc[0]
new_rp = new_rp.iloc[0:0]

new_rp['Facility Name'] = df_rp['Facility Name']
new_rp['Country'] = country
new_rp['Release ID'] = df_rp['Point Of Release']
new_rp['Release Type'] = df_rp['Source Type']
new_rp['Stack Height [m]'] = df_rp['Release Height [m]']
new_rp['Stack Diameter [m]'] = df_rp['Diameter [m]']
new_rp['Exit Temperature [C]'] = df_rp['Exit Temperature [C]']
new_rp['Exit Velocity [m/s]'] = df_rp['Exit Velocity [m/s]']
new_rp['Latitude [deg]'] = df_rp['Latitude']
new_rp['Longitude [deg]'] = df_rp['Longitude']
new_rp['UTM-X [m]'] = ''
new_rp['UTM-Y [m]'] = ''
new_rp['UTM Zone'] = ''
new_rp['Source ID'] = df_rp['Point Of Release']
new_rp['Description'] = df_rp['Comment']
new_rp['Base Elevation [m]'] = df_rp['Base Elevation']
new_rp['Notes'] = df_rp['Comment']
new_rp['Flow Rate [m^3/s]'] = ''
new_rp['Fugitive Type'] = 'Area Fugitive Source' # make it general for all
new_rp['Release Height [m]'] = df_rp['Area Release Height [m]']
new_rp['Initial Lateral Dimension [m]'] = df_rp['Volume Initial Lateral Dimension [m]']
new_rp['Initial Vertical Dimension [m]'] = df_rp['Area Initial Vertical Dimension [m]']
new_rp['Horizontal Area [sq. m]'] = ''
new_rp['Length of X Side [m]'] = df_rp['Area Length Side X [m]']
new_rp['Length of Y Side [m]'] = df_rp['Area Length Side Y [m]']
new_rp['Orientation Angle [deg]'] = df_rp['Area Orientation Angle [deg]']
new_rp['Projected Datum'] = 'WGS84'

# see if there is an active column in the work sheet
if 'Active' in df_rp.columns:
    new_rp['Active'] = df_rp['Active']    
else:
    new_rp['Active'] = 'T'

# Merge AQZ from new_fac by calling the add_aqz function
new_rp = add_aqz(new_rp)

new_rp = pd.concat((new_rp_hd, new_rp), axis = 0)
new_rp = drop_index(new_rp)

# ****************************************
# * New Emission Units sheet
new_eu = new_template2.parse('Emission Units')
new_eu.columns = new_eu.iloc[0]
new_eu = new_eu.iloc[0:0]

new_eu['Facility Name'] = df_source['Facility Name']
new_eu['Country'] = country
new_eu['Emission Unit ID'] = df_source['Emission Unit']
new_eu['Unit Status'] = 'Operating'
new_eu['Notes'] = df_source['Emission Unit Description']
new_eu['Unit Type'] = ''
new_eu['Description'] = df_source['Emission Unit Description']
new_eu['Unit Status Year'] = ''
new_eu['Unit Operation Date'] = ''
      
# Merge AQZ from new_fac
new_eu = add_aqz(new_eu)

new_eu = pd.concat((new_eu_hd, new_eu), axis = 0)
new_eu = drop_index(new_eu)

# ****************************************
# * New Processes sheet
new_pro = new_template2.parse('Processes')
new_pro.columns = new_pro.iloc[0]
new_pro = new_pro.iloc[0:0]

new_pro['Emission Unit ID'] = df_source['Emission Unit']
new_pro['Facility Name'] = df_source['Facility Name']
new_pro['Country'] = country
new_pro['Process ID'] = df_source['Emission Process']
new_pro['Process Description'] = df_source['Emission Process Description']
new_pro['Source Classification Code (SCC)'] = df_source['SCC']
new_pro['Process Comment'] = df_source['SCC_Name']
new_pro['Last Emissions Year'] = ''
new_pro['Percent Winter Activity'] = ''
new_pro['Percent Spring Activity'] = ''
new_pro['Percent Summer Activity'] = ''
new_pro['Percent Fall Activity'] = ''
new_pro['Average Days per Week'] = ''
new_pro['Average Weeks per Year'] = ''
new_pro['Average Hours per Day'] = ''
new_pro['Average Hours per Year'] = ''

# Merge AQZ from new_fac
new_pro = add_aqz(new_pro)

new_pro = pd.concat((new_pro_hd, new_pro), axis = 0)
new_pro = drop_index(new_pro)

# ****************************************
# * New Apportionment sheet
new_apport = new_template2.parse('Apportionment')
new_apport.columns = new_apport.iloc[0]
new_apport = new_apport.iloc[0:0]

new_apport['Process ID'] = df_source['Emission Process']
new_apport['Emission Unit ID'] = df_source['Emission Unit']
new_apport['Facility Name'] = df_source['Facility Name']
new_apport['Country'] = country
new_apport['Release Point ID'] = df_source['Point Of Release']
new_apport['Emissions Apportionment [%]'] = 100
new_apport['Comment'] = df_source['Comment']
new_apport['Notes'] = ''
          
# Merge AQZ from new_fac
new_apport = add_aqz(new_apport)

new_apport = pd.concat((new_apport_hd, new_apport), axis = 0)
new_apport = drop_index(new_apport)

# ****************************************
# * New Emission Periods sheet
new_period = new_template2.parse('Emission Periods')
new_period.columns = new_period.iloc[0]
new_period = new_period.iloc[0:0]

new_period['Process ID'] = df_source['Emission Process']
new_period['Emission Unit ID'] = df_source['Emission Unit']
new_period['Facility Name'] = df_source['Facility Name']
new_period['Country'] = country
new_period['Reporting Period'] = 'Annual'
new_period['Start Date'] = start_date
new_period['End Date'] = end_date
new_period['Operation Type'] = 'Routine'
new_period['Emission Category'] = 'Actual'
new_period['Emission Type'] = 'ENTIRE PERIOD'
new_period['Period Description'] = ''
new_period['Calculation Data Year'] = '2016'
new_period['Comments'] = ''
new_period['Throughput Value'] = df_emissions['Fuel Use']
new_period['Throughput Unit'] = df_emissions['Fuel Use (Units)']
new_period['Material'] = df_emissions['Fuel Type']
new_period['Material Process Type'] = 'PROCESS MATERIAL USED (INPUT)'
new_period['Throughput Calculation Source'] = ''
new_period['Average Days per Week'] = ''
new_period['Average Weeks per Period'] = ''
new_period['Average Hours per Day'] = ''
new_period['Hours per Period'] =  df_emissions['Op Hours']
new_period['Heat Content'] = ''
new_period['[Million BTU Per]'] = col_check(df_emissions, '[Million BTU Per]', new_period)
new_period['Ash Content [mass %]'] = col_check(df_emissions, 'Ash Content [mass %]', new_period)
new_period['Sulfur Content [mass %]'] = col_check(df_emissions, 'Sulfur Content [mass %]', new_period)

# Merge AQZ from new_fac
new_period = add_aqz(new_period)

new_period = pd.concat((new_period_hd, new_period), axis = 0)
new_period = drop_index(new_period)

# ****************************************
# * New Emission sheet
new_emissions = new_template.parse('Emissions')
new_emissions.columns = new_emissions.iloc[0]
new_emissions = new_emissions.iloc[0:0]

##### Emissions must be prepared row by row
'''
new_emissions['Reporting Period'] = '
new_emissions['Process ID'] = df_source['
new_emissions['Emission Unit ID'] = df_source['
new_emissions['Facility Name'] = df_source['
new_emissions['Country'] = df_source['
new_emissions['Air Quality Zone'] = df_source['
new_emissions['Pollutant'] = df_source['
new_emissions['Emission Value'] = df_source['
new_emissions['Emissions Unit of Measure'] = df_source['
new_emissions['Emission Calculation Method'] = df_source['
new_emissions['Emission Estimation Method'] = df_source['
new_emissions['Emission Factor (EF)'] = df_source['
new_emissions['EF Unit Numerator'] = df_source['
new_emissions['EF Unit Denominator'] = df_source['
new_emissions['Short Term Emission'] = df_source['
new_emissions['Short Term Emission Unit Numerator'] = df_source['
new_emissions['Short Term Emission Unit Denominator'] = df_source['
new_emissions['Emission Reliability Indicator'] = df_source['
new_emissions['EF Reliability Indicator'] = df_source['
new_emissions['Rule Effectiveness ['%']'] = df_source['
new_emissions['Rule Effectiveness Method'] = df_source['
new_emissions['Emissions Comment'] = df_source['
'''
# Merge AQZ from new_fac
new_emissions = add_aqz(new_emissions)

new_emissions = pd.concat((new_emissions_hd, new_emissions), axis = 0)
new_emissions = drop_index(new_emissions)


# **************************************************
# Save in new AQMIS format as xlsx 
save_df = input('\nSave new AQMIS Template for ' + company +'? (y/n) ')

#new_path = 'C:\TARS\AAALakes\Kuwait AQMIS\'
if save_df == 'y':   
    os.chdir("\TARS\AAALakes\Kuwait AQMIS") # reset to new folder path 
    time_stamp = str(int(time.time()))[-6:] # add unique stamp to file name
    
    writer = pd.ExcelWriter('AQMIS_new_format_'+company+'_'+time_stamp+'.xlsx')
    
    # write worksheets. Save w/o index or column header
    new_fac.to_excel(writer, sheet_name = 'Facilities', index = False, header = False)
    new_rp.to_excel(writer,sheet_name = 'Release Points', index = False, header = False)
    new_eu.to_excel(writer,sheet_name = 'Emission Units', index = False, header = False)
    new_pro.to_excel(writer,sheet_name = 'Processes', index = False, header = False)
    new_apport.to_excel(writer,sheet_name = 'Apportionment', index = False, header = False)
    new_period.to_excel(writer,sheet_name = 'Emission Periods', index = False, header = False)
    new_emissions.to_excel(writer,sheet_name = 'Emissions', index = False, header = False)
    writer.save()
