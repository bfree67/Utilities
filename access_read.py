# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 12:08:18 2019
Loads exported excel sheet from Access table and counts unique occurrences for each field
Counts blanks for each field and sticks on bottom of worksheet as 0
Makes summary table and percentage of blanks
@author: Brian
"""

import pandas as pd
version = '2.1.1'

#for 2017
df_file = '2017_Gulfwide_Platform_20190705_CAP_GHG.xlsx'; df_out = '2017_platform_summary.xlsx'
#df_file = 'BOEM_2017_Monthly_NonPlatformHAP.xlsx'; df_out = '2017_non_platformHAP_summary.xlsx'
#df_file = 'BOEM_2017_NonPlatformCAP_GHG_Summary.xlsx'; df_out = '2017_non_platformGHG_summary.xlsx'

#### for 2014
#df_file ='Query1.xlsx'; df_out = 'output_1_summary.xlsx'
#df_file ='Query2.xlsx'; df_out = 'output_2_summary.xlsx'
#df_file = 'Query1_NonPlatform2014.xlsx' ; df_out = 'output_1_nonplatform_summary.xlsx'

print('\nLoading ' + df_file)
df = pd.read_excel(df_file)

a_col = df.columns.values.tolist()
tot_obs = len(df)

a_percent = []
with pd.ExcelWriter(df_out) as writer:
    print('Saving to ' + df_out)
    df_list = df.nunique()
    df_list.to_excel(writer, sheet_name='Summary', index_label = 'FIELD')
    for col_name in a_col:
        a_p = []
        if col_name != 'EMISSIONS_VALUE' or col_name != 'THROUGHPUT_VALUE':
            df_list = df[col_name].value_counts()
            df_list.columns = ['Count']
            df_list.sort_index(inplace=True)
            df_list = df_list.to_frame()
            diff = tot_obs-df_list.sum()
            df1 = pd.DataFrame([diff]) #count number of blanks
            a_p = [col_name, float(diff/tot_obs) ] #calc percentage of blanks in field
            
            df1.columns = [col_name]
            df_list = df_list.append(df1)
            
            #write field category count to worksheet
            df_list.to_excel(writer, sheet_name=col_name[0:30], index_label = col_name)
            
            #append blanks percentage to list
            a_percent.append(a_p)
            
    df_percent=pd.DataFrame(a_percent)
    df_percent.to_excel(writer, sheet_name='Blanks', index_label = 'FIELD')