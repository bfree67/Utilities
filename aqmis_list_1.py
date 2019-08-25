# -*- coding: utf-8 -*-
"""
Created on Wed Aug 21 14:53:06 2019

Utility to remap AQMIS templates using old import format to new import format.
Requires all old template spreadsheets to be in the same folder using *.xlsx format
and a modified blank template with only column headers to in an other folder.

@author: Brian, Python 3.5
"""
import glob
import pandas as pd
import os

current_path = os.getcwd()  # get current working path to save later

txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

# merge the excel files into one file by reading all excel files in the folder
for file in glob.glob("*.xlsx"):
    txtfiles.append(file)
    
    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    xl_names = xl.sheet_names
    
    if i == 0:
        df_fac = xl.parse('Facility Information')
        df_rp = xl.parse('Release Points')
        df_source = xl.parse('Source Information')
        df_emissions = xl.parse('Emissions')
        i += 1     
    
    else:
        df_fac = pd.concat([df_fac, xl.parse('Facility Information')], sort = False)
        df_rp = pd.concat([df_rp, xl.parse('Release Points')], sort = False)
        df_source = pd.concat([df_source, xl.parse('Source Information')], sort = False)
        df_emissions = pd.concat([df_emissions, xl.parse('Emissions')], sort = False)
    
    print(file)

# location of blank template with only column headers    
new_file = 'C:\\TARS\\AAALakes\\Kuwait AQMIS\\Blank_Template_Mod.xlsx'    
new_template = pd.ExcelFile(new_file)
new_names = new_template.sheet_names

# global constants for country and company
country = 'Kuwait'
company = 'MEW'

# ****************************************
# * New Facility sheet
new_fac = new_template.parse('Facilities')
new_fac['Facility Name'] = df_fac['Facility Name']
new_fac['Country'] = 'Kuwait'
new_fac['Air Quality Zone'] = df_fac['AQZ']
new_fac['Company Name'] = 'MEW'
new_fac['Latitude [deg]'] = df_fac['Latitude [decimal deg]']
new_fac['Longitude [deg]'] = df_fac['Longitude [decimal deg]']
new_fac['Governorate'] = df_fac['Location']
new_fac['SIC'] = df_fac['SIC']

# ****************************************
# * New Release Points sheet
new_rp = new_template.parse('Release Points')

new_rp['Facility Name'] = df_rp['Facility Name']
new_rp['Country'] = country
new_rp['Release ID'] = df_rp['Point Of Release']
new_rp['Release Type'] = df_rp['Source Type']
new_rp['Stack Height [m]'] = df_rp['Release Height [m]']
new_rp['Stack Diameter [m]'] = df_rp['Diameter [m]']
new_rp['Exit Temperature [C]'] = df_rp['Exit Temperature [C]']
new_rp['Exit Velocity [m/s]'] = df_rp['Exit Velocity [m/s]']
new_rp['Latitude [deg]'] = df_rp['Latitude [decimal deg]']
new_rp['Longitude [deg]'] = df_rp['Longitude [decimal deg]']
new_rp['UTM-X [m]'] = ''
new_rp['UTM-Y [m]'] = ''
new_rp['UTM Zone'] = ''
new_rp['Source ID'] = ''
new_rp['Description'] = ''
new_rp['Base Elevation [m]'] = df_rp['Base Elevation [m]']
new_rp['Active'] = 'T'
new_rp['Notes'] = df_rp['Comment']
new_rp['Flow Rate [m^3/s]'] = ''
new_rp['Fugitive Type'] = ''
new_rp['Release Height [m]'] = df_rp['Area Release Height [m]']
new_rp['Initial Lateral Dimension [m]'] = df_rp['Volume Initial Lateral Dimension [m]']
new_rp['Initial Vertical Dimension [m]'] = df_rp['Area Initial Vertical Dimension [m]']
new_rp['Horizontal Area [sq. m]'] = ''
new_rp['Length of X Side [m]'] = df_rp['Area Length Side X [m]']
new_rp['Length of Y Side [m]'] = df_rp['Area Length Side Y [m]']
new_rp['Orientation Angle [deg]'] = df_rp['Area Orientation Angle [deg]']
new_rp['Projected Datum'] = ''

# Merge AQZ from new_fac
df_aqz = pd.merge(left=new_rp,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_rp['Air Quality Zone'] = df_aqz['Air Quality Zone_y']

# ****************************************
# * New Emission Units sheet
new_eu = new_template.parse('Emission Units')

new_eu['Facility Name'] = df_source['Facility Name']
new_eu['Country'] = country
new_eu['Emission Unit ID'] = df_source['Emission Unit']
new_eu['Unit Status'] = ''
new_eu['Notes'] = df_source['Emission Unit Description']
new_eu['Unit Type'] = ''
new_eu['Description'] = df_source['Emission Unit Description']
new_eu['Unit Status Year'] = ''
new_eu['Unit Operation Date'] = ''
      
# Merge AQZ from new_fac
df_aqz = pd.merge(left=new_eu,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_eu['Air Quality Zone'] = df_aqz['Air Quality Zone_y']

# ****************************************
# * New Processes sheet
new_pro = new_template.parse('Processes')

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
df_aqz = pd.merge(left=new_pro,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_pro['Air Quality Zone'] = df_aqz['Air Quality Zone_y']

# ****************************************
# * New Apportionment sheet
new_apport = new_template.parse('Apportionment')

new_apport['Process ID'] = df_source['Emission Process']
new_apport['Emission Unit ID'] = df_source['Emission Unit']
new_apport['Facility Name'] = df_source['Facility Name']
new_apport['Country'] = country
new_apport['Release Point ID'] = df_source['Point Of Release']
new_apport['Emissions Apportionment [%]'] = 100
new_apport['Comment'] = df_source['Comment']
new_apport['Notes'] = ''
          
# Merge AQZ from new_fac
df_aqz = pd.merge(left=new_apport,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_apport['Air Quality Zone'] = df_aqz['Air Quality Zone_y']

# ****************************************
# * New Emission Periods sheet
new_period = new_template.parse('Emission Periods')

new_period['Process ID'] = df_source['Emission Process']
new_period['Emission Unit ID'] = df_source['Emission Unit']
new_period['Facility Name'] = df_source['Facility Name']
new_period['Country'] = country
new_period['Reporting Period'] = 'Annual'
new_period['Start Date'] = ''
new_period['End Date'] = ''
new_period['Operation Type'] = ''
new_period['Emission Category'] = 'Actual'
new_period['Emission Type'] = 'ENTIRE PERIOD'
new_period['Period Description'] = ''
new_period['Calculation Data Year'] = ''
new_period['Comments'] = ''
new_period['Throughput Value'] = ''
new_period['Throughput Unit'] = ''
new_period['Material'] = ''
new_period['Material Process Type'] = ''
new_period['Throughput Calculation Source'] = ''
new_period['Average Days per Week'] = ''
new_period['Average Weeks per Period'] = ''
new_period['Average Hours per Day'] = ''
new_period['Hours per Period'] = ''
new_period['Heat Content'] = ''
new_period['[Million BTU Per]'] = ''
new_period['Ash Content [mass %]'] = ''
new_period['Sulfur Content [mass %]'] = ''

# Merge AQZ from new_fac
df_aqz = pd.merge(left=new_period,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_period['Air Quality Zone'] = df_aqz['Air Quality Zone_y']

# ****************************************
# * New Emission sheet
new_emissions = new_template.parse('Emissions')

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

# Merge AQZ from new_fac
df_aqz = pd.merge(left=new_emissions,right=new_fac, 
                  how='left', 
                  on='Facility Name')

new_emissions['Air Quality Zone'] = df_aqz['Air Quality Zone_y']
'''

# **************************************************
# Save in new AQMIS format as xlsx 
save_df = input('\nSave new AQMIS Template for ' + company +'? (y/n) ')
if save_df == 'y':   
    os.chdir(current_path) # reset to current folder path just in case it was changed
    writer = pd.ExcelWriter('AQMIS_new_format_template_'+company+'.xlsx')
    
    # write worksheets
    new_fac.to_excel(writer, sheet_name = 'Facilities', index = False)
    new_rp.to_excel(writer,sheet_name = 'Release Points', index = False)
    new_eu.to_excel(writer,sheet_name = 'Emission Units', index = False)
    new_pro.to_excel(writer,sheet_name = 'Processes', index = False)
    new_apport.to_excel(writer,sheet_name = 'Apportionment', index = False)
    new_period.to_excel(writer,sheet_name = 'Emission Period', index = False)
    new_emissions.to_excel(writer,sheet_name = 'Emissions', index = False)
    writer.save()
