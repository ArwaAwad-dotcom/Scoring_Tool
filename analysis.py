# -*- coding: utf-8 -*-
"""
Created on Fri Jan 24 11:34:26 2025

@author: user
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta



# Function Date transformation

def transform_date_column(df, column_name):

    # Iterate through DataFrame, transform datetime -> strings
    for i in range(len(df)):
        if 'datetime.datetime' in str(type(df[column_name].iloc[i])):
            value = df[column_name].iloc[i]
            year = value.year
            month = value.month
            day = value.day
            df.at[i, column_name] = f"{month}/{day}/{year}"

    # Convert the column to the desired datetime format
    df[column_name] = pd.to_datetime(df[column_name], format="%d/%m/%Y", errors='coerce')

    return df

def get_relevant_data(file_name): 
    
    df=pd.read_excel(file_name, sheet_name='Talent Info',converters={'Answers': str})
    df.columns=['Questions','Answers']
    answers=df['Answers']
    date_of_engagement=answers[11]
    engagement_name=answers[9]
    engagement_id=answers[10]
    
    
    if 'datetime.datetime' in str(type(date_of_engagement)): 
            year = date_of_engagement.year
            month =date_of_engagement.month
            day = date_of_engagement.day
            date_of_engagement= f"{month}/{day}/{year}"

    
  
    #print("Date of engagement : ", date_of_engagement)


    #Read the assessment sheet

    dg=pd.read_excel(file_name,sheet_name='Engagement Assessment')


    data_lookup=dg[7:70]
    #get the name of the columns that need to be added
    column_names=list(data_lookup.iloc[0])

    data_lookup=data_lookup[1:70]
    data_lookup.columns=column_names

    evaluator=[]
    #Check who is filling the assessment
    save_value=""
    save_value_level=""
    
    for i in range(len(data_lookup)): 
        if pd.isnull(data_lookup['Level 2 and Description '].iloc[i]): 
            evaluator.append('Reviewer') 
            data_lookup['Level 2 and Description '].iloc[i]=save_value 
        else: 
            evaluator.append('Talent')
            save_value=data_lookup['Level 2 and Description '].iloc[i]
        
        if pd.isnull(data_lookup['Level 1'].iloc[i]): 
            data_lookup['Level 1'].iloc[i]=save_value_level 
        else: 
            save_value_level=data_lookup['Level 1'].iloc[i]
        
        
     
    data_lookup['Evaluater']=evaluator
    data_lookup['Engagement Name']=engagement_name
    data_lookup['Engagement ID']=engagement_id


    data_final=data_lookup[['Level 1', 'Level 2 and Description ','Exposure & Stretch\n(0-no exposure, 0.2-low stretch & exposure, 0.5-normal stretch & exposure, 0.8-high stretch & exposure)','Analyst', 'Sr. Analyst',
'Venture Builder', 'Sr. Venture Buidler', 'Portfolio Manager',
'Sr. Portfolio Manager', 'Director', 'Evaluater','Engagement Name','Engagement ID']]

    data_final.columns=['Level 1', 'Level 2','Exposure & Stretch','Analyst', 'Sr. Analyst',
'Venture Builder', 'Sr. Venture Buidler', 'Portfolio Manager',
'Sr. Portfolio Manager', 'Director', 'Evaluater','Engagement Name','Engagement ID']

    data_final=data_final[data_final['Evaluater']=="Reviewer"]


    data_final['Date of Reviewing']=pd.to_datetime(date_of_engagement,dayfirst=True)
    
    return data_final


def calculate_weights(all_data): 
    
    all_data_levels_two=list(all_data['Level 2'].unique()) 
    list_dates_all_position=list(all_data.columns[3:9])
    matrix_level_pos=np.zeros((len(all_data_levels_two),len(list_dates_all_position)))
    
    for dd in all_data_levels_two:
        
        index_dd=all_data_levels_two.index(dd)
        
        data_sub=all_data[all_data['Level 2']==dd]
        for pos in list_dates_all_position:
            index_pos=list_dates_all_position.index(pos)
            data_sub_two=data_sub[data_sub[pos].notna()]
            
            if len(data_sub_two)!=0:
                list_weights= list(data_sub_two['Exposure & Stretch'])
                list_date_delivers=list(data_sub_two['Date of Reviewing'])
                list_grades=list(data_sub_two[pos])
                combined = list(zip(list_date_delivers, list_weights, list_grades))
                combined.sort(key=lambda x: x[0], reverse=True)
                
                selected = []
                current_weight_sum = 0
                
                for item in combined:
                    if current_weight_sum < 1:
                        selected.append(item)
                        current_weight_sum += item[1] 
                        
                        
                sorted_dates, sorted_weights, sorted_grades = zip(*selected)
                sorted_weights=np.array(list(sorted_weights))
                sorted_grades=np.array(list(sorted_grades))
                grades_weigh=(np.dot(sorted_weights,sorted_grades))/(np.sum(sorted_weights))

                
                if np.sum(sorted_weights) <1 :
                    grades_weigh=(np.sum(sorted_weights))* grades_weigh 
                    
                matrix_level_pos[index_dd,index_pos]=grades_weigh
            else: 
                matrix_level_pos[index_dd,index_pos]=np.nan

    
    return matrix_level_pos




def combine_baseline_data(baseline_file,all_data):
    bas_data=pd.read_excel(baseline_data,sheet_name="CF-Baselining")

    bas_data_data_frame=bas_data.iloc[5:36]
    bas_data_col=list(bas_data.iloc[4])
    bas_data_col=bas_data_col[:3]
    bas_data_column_two=list(bas_data.iloc[3,3:10])
    bas_data_col.extend(bas_data_column_two)

    bas_data_data_frame.columns=bas_data_col
    save_value_level=""
    for i in range(len(bas_data_data_frame)): 
        if pd.isnull(bas_data_data_frame['Level 1'].iloc[i]): 
            bas_data_data_frame['Level 1'].iloc[i]=save_value_level 
        else: 
            save_value_level=bas_data_data_frame['Level 1'].iloc[i]


    minimum_date=min(all_data['Date of Reviewing']) - relativedelta(years=1)

    bas_data_data_frame['Evaluator']='Reviewer'
    bas_data_data_frame['Engagement Name']="Previous Year"
    bas_data_data_frame['Engagement ID']=""
    bas_data_data_frame['Date of Reviewing']=minimum_date

    bas_data_data_frame.columns=all_data.columns
    all_data=pd.concat([all_data,bas_data_data_frame])
    
    return all_data
    


    



                
                


                    

    

