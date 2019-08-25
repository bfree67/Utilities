# -*- coding: utf-8 -*-
"""
Created on Wed Aug 21 14:53:06 2019
Reads excel files that have some EDD column headers from multiple xls files
in a folder. The read files must have an .xls extension or will be ignored.

The files are merged into a super dataframe, agg. 

The EDD format is read as an xlsx file (KEPA.xlsx).

The agg dataframe is broken up into Field and Lab EDD formats. Blank columns
are inserted if there is no match in the agg dataframe.

Some columns are populated with known values.

@author: Brian, Python 3.5
"""
import glob
import pandas as pd

txtfiles = []
agg = pd.DataFrame
i = 0 # counter to begin with

df_lab = pd.read_excel('KEPA.xlsx', sheet_name = 0)
lst_lab = list(df_lab)

df_field = pd.read_excel('KEPA.xlsx', sheet_name = 1)
lst_field = list(df_field)


# read the excel files in the directory and save to a list
for file in glob.glob("*.xls"):
    txtfiles.append(file)
    

    xl = pd.ExcelFile(file) # read the indvidual excel file
    res = len(xl.sheet_names) # count the number of sheets in the file
    
    print(file, res)
    
    if i == 0:   # on the first sheet of the first file, read only
        agg = pd.read_excel(file, sheet_name = 0)
        i += 1
        
    else:
        
        if i == 1 and res > 1:  # on the next sheet, start one sheet over to avoid duplicate read
            
            for j in range (1, res):
                            
                agg1 = pd.read_excel(file, sheet_name = j)
                agg = pd.concat([agg,agg1])
                i += 1
                
        else: 
            
            for j in range (res):
                            
                agg1 = pd.read_excel(file, sheet_name = j)
                agg = pd.concat([agg,agg1])
                i += 1
                
# remove non-tertiary treaments amples
sub ='tertiary'
agg = agg[agg["Sample_Name"].str.contains(sub)==True]

# re-name all sample parameters 
agg["#facility_code"]="WWKBD"
agg["Location_Code"] = "SPS3"
     

print ('\n',len(txtfiles), 'files processed')     

# ************************************         
# ******* Make Field EDD
                
df_newField = pd.DataFrame() # make blank dataframe

for field_col in lst_field: # list of field EDD columns
    count = 0 # reset counter used to see if there is matching column
    for agg_col in list(agg): #list of existing columns
        if field_col == agg_col and count == 0: # if columns match, add column to dataframe
            count += 1
            df_newField[field_col] = agg[field_col].tolist()
            
        if field_col != agg_col and count > 0: # if no match, break
            break
        
    if count == 0:  # if no match, add a blank column
         df_newField[field_col] = ""
         
# remove duplicate samples and keep first occurence
df_newField = df_newField.drop_duplicates(subset=['Sample_Code'], keep = 'first')

# pre-populate some fields
df_newField['Sample_Name'] = 'Tertiary Effluent'
df_newField['Sample_Matrix_Code'] = 'WD'
df_newField['Sample_Type_Code'] = 'N'
df_newField['sample_method'] = 'Grab'
df_newField['Sampling_Company_Code'] = 'MPW'

# ********************************
# ******* Make Lab EDD
                
df_newLab = pd.DataFrame()
for lab_col in lst_lab:
    count = 0
    for agg_col in list(agg):
        if lab_col == agg_col and count == 0:
            count += 1
            df_newLab[lab_col] = agg[lab_col].tolist()
            
        if lab_col != agg_col and count > 0:
            break
        
    if count == 0:
         df_newLab[lab_col] = ""       
         
# remove rows without result values
df_newLab = df_newLab[df_newLab['Result_Value'].notnull()]

# remove rows with ROE chemical name
df_newLab = df_newLab[df_newLab['Chemical_Name']!='ROE']

# reset index
df_newLab = df_newLab.reset_index()
df_newLab = df_newLab.drop('index', axis=1)

# revise data format to mm/dd/yyyy 
sdates = df_newLab.Sample_Date.str.split('/')
num_dates = len(sdates)
new_date = []
for i in range(num_dates):
    new_date_str = str(sdates[i][1]) + '/' + str(sdates[i][0]) + '/' + str(sdates[i][2])
    new_date.append(new_date_str)
    
df_newLab['Sample_Date'] = pd.DataFrame({'Sample_date':new_date})
    
# pre-populate some fields
df_newField['Sample_Name'] = 'Tertiary Effluent'
df_newLab['Lab_Name_Code'] = 'MPW'    
df_newLab['Sample_Matrix_Code'] = 'WD'
df_newLab['Lab_Matrix_Code'] = 'WD'
df_newLab['Lab_Sample_Id'] = df_newLab['Sample_Code']
df_newLab['Analysis_Date'] = df_newLab['Sample_Date']
df_newLab['Analysis_Time'] = '10:00'
df_newLab['Result_Type_Code'] = 'TRG'

# **************************************************
# Save in KEDD format as xlsx filesave_df = input('\nSave new EDDs? (y/n)')
save_df = input('\nSave new EDD? (y/n) ')
if save_df == 'y':   

    writer = pd.ExcelWriter('Kabd_new_EDD.xlsx')
    df_newLab.to_excel(writer, sheet_name = 'KEPA_lab', index = False)
    df_newField.to_excel(writer,sheet_name = 'KEPA_field', index = False)
    writer.save()

