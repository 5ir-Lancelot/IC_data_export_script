# -*- coding: utf-8 -*-
"""
Created on Mon May  8 16:31:25 2023


"""
#open the IC data


import phreeqpython
import pandas as pd

import pandas as pd
import numpy as np

import datetime
from datetime import date

import sys
import os

import plotly.graph_objects as go

#for color of lines and scatter
import plotly.express as px

import copy

import matplotlib.pyplot as plt

import plotly.io as pio

from itertools import cycle

pio.renderers.default='browser'





'''

type in the filepath were the data can be found


'''

#where all the data are
path='./IC_data_export/round_5'

data_dirs=os.listdir(path)



'''
Download the IC and TA data to check if all data are present in both tables
'''

df_IC_1=pd.read_csv(path+'/'+data_dirs[0],sep=';',decimal=',',encoding='latin-1')

df_IC_2=pd.read_csv(path+'/'+data_dirs[1],sep=';',decimal=',',encoding='latin-1')





#create matching column names
df_IC_1.rename(columns={'Sample_ID':'Unique_ID'},inplace=True)

df_IC_2.rename(columns={'Sample_ID':'Unique_ID'},inplace=True)



#merge/concat dataframes

df_IC=pd.concat([df_IC_1,df_IC_2])



#compare with alkalinity to get the unknown number

df_TA=pd.read_excel('data/Round 5/TA.xlsx',
                          sheet_name='Final Sample',na_filter=True)



#get columsn names
cols=df_IC.columns


#create the batch_id from the unique
df_IC['batch_id']=[int(item.split('-')[1]) if not '?' in item.split('-')[1] else item.split('-')[1]  for item in df_IC['Unique_ID'] ]
    
    
    

#use merge to find the differences


lol=df_TA.merge(df_IC, how='left', on='batch_id')



cols=lol.columns

#missing samples


miss=lol.dropna(subset=['Unique_ID'])


missing=lol[lol['Unique_ID'].isnull()]



missing.to_excel('BAM_round_5_missing_samples.xlsx')



'''
code to calculate the average values for the samples

from the 3 IC shots (one shot = one measurement)


calculate average and standard deviation 

'''


#exceptions for specific batch_id's



# make some groups for the same measurement (groupby the Unique_ID)
# and calculate the mean for each column/ variable
df_IC_avg=df_IC.groupby(['batch_id']).mean()

#get the standard deviations of the measurement
df_IC_std=df_IC.groupby(['batch_id']).std()

#coefficient of variation 
df_IC_CV=100*df_IC.groupby(['batch_id']).std()/df_IC.groupby(['batch_id']).mean()


#add 


#make the batch_id a column again

df_IC_avg['batch_id'] = df_IC_avg.index
df_IC_std['batch_id'] = df_IC_std.index
df_IC_CV['batch_id'] = df_IC_CV.index

#remove the old index
# remove the own index added by groupby function with default index 0-n
df_IC_avg.reset_index(inplace = True,drop = True)
df_IC_std.reset_index(inplace = True,drop = True)
df_IC_CV.reset_index(inplace = True,drop = True)

#merge both dataframes 

# add one column for each variable for  the decision 'remove'

list_c=list(df_IC_avg.columns)

list_c.remove('batch_id')

# add 'removed' suffix to the dataframe

list_c=[sub+'_removed' for sub in list_c]


# merge all daframes pure data, mean, std  and CV

df_IC_merge=df_IC.merge(df_IC_avg, how='left', on='batch_id',suffixes=('','_mean')).merge(df_IC_std,how='left', on='batch_id',suffixes=('','_std')).merge(df_IC_CV,how='left', on='batch_id',suffixes=('','_CV'))



#add the '_removed' columns
df_IC_merge[[list_c]]=''



#move the batch id to front




#for the BAM export just use the average values (avg)



column_list = list(df_IC_merge.columns)

#Now rearrange the list the way you want the columns to be
#Then do

column_list.sort(reverse=False)

final_df = df_IC_merge[column_list]



first_column = final_df.pop('batch_id')
  
# insert column using insert(position,column_name,
# first_column) function
final_df.insert(0, 'batch_id', first_column)


first_column = final_df.pop('Unique_ID')
  
# insert column using insert(position,column_name,
# first_column) function
final_df.insert(0,'Unique_ID', first_column)


final_df.sort_values(by=['Unique_ID'],inplace=True)




#drop unnecessary columns for BAM (Hamburg internal stuff)

#df_IC_avg.drop(columns=['NR', 'pos','dil'])

df=final_df

df.to_excel('IC_data_export/round_5/round_5_IC_data.xlsx', index=False)



# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter('IC_data_export/round_5/pre-processed_BAM_round_5.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Define the format for the rows with different values in Column1
row_color_format = workbook.add_format({'bg_color': '#F0EEA0'})
black_line_format = workbook.add_format({'bottom': 2, 'border_color': 'black'})

# Iterate through the DataFrame and apply the formats based on the condition
for i in range(1, len(df)):
    if df.iloc[i]['batch_id'] != df.iloc[i - 1]['batch_id']:
        # Apply color format to the entire row
        worksheet.set_row(i, None, row_color_format)

        # OR apply black line format to the entire row
        
        # worksheet.set_row(i + 1, None, black_line_format)
        
        
        # coefficient of variation (CV)
        #apply functions
        
        #worksheet.write_furmula(row_num, col, formula_to_write)
        
        worksheet.write_formula(i,2,'=SUM(1,2,3)')

# Close the Pandas Excel writer and output the Excel file
writer.save()

