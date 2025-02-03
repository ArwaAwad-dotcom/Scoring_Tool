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



    



def get_relevant_data(file_name,level_one_list,level_two_list): 
    
    df=pd.read_excel(file_name, sheet_name='Talent Info',converters={'Answers': str})
    
    df=df.iloc[:,:2]
    
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

    



    #Read the assessment sheet

    dg=pd.read_excel(file_name,sheet_name='Engagement Assessment')


    data_lookup=dg[7:70]
    #get the name of the columns that need to be added
    column_names=list(data_lookup.iloc[0])
    column_names[0]='Level 1'
    column_names[1]='Level 2'
    column_names[2]='Weight'

    data_lookup=data_lookup[1:70]
    data_lookup.columns=column_names
    
    
    print(data_lookup.columns)

    evaluator=[]
    #Check who is filling the assessment
    save_value=""
    save_value_level=""
    
    for i in range(len(data_lookup)): 
        if pd.isnull(data_lookup['Level 2'].iloc[i]): 
            evaluator.append('Reviewer') 
            data_lookup['Level 2'].iloc[i]=save_value 
        else: 
            evaluator.append('Talent')
            save_value=data_lookup['Level 2'].iloc[i]
        
        if pd.isnull(data_lookup['Level 1'].iloc[i]): 
            data_lookup['Level 1'].iloc[i]=save_value_level 
        else: 
            save_value_level=data_lookup['Level 1'].iloc[i]
        
        
     
    data_lookup['Evaluater']=evaluator
    data_lookup['Engagement Name']=engagement_name
    data_lookup['Engagement ID']=engagement_id


    data_final=data_lookup[['Level 1', 'Level 2','Weight','Analyst', 'Sr. Analyst',
'Venture Builder', 'Sr. Venture Buidler', 'Portfolio Manager',
'Sr. Portfolio Manager', 'Director', 'Evaluater','Engagement Name','Engagement ID']]

    data_final.columns=['Level 1', 'Level 2','Weight','Analyst', 'Sr. Analyst',
'Venture Builder', 'Sr. Venture Buidler', 'Portfolio Manager',
'Sr. Portfolio Manager', 'Director', 'Evaluater','Engagement Name','Engagement ID']

    data_final=data_final[data_final['Evaluater']=="Reviewer"]


    data_final['Date of Reviewing']=pd.to_datetime(date_of_engagement,dayfirst=True)
    
    data_final['Level 1']=level_one_list
    data_final['Level 2']=level_two_list
    
    return data_final






def calculate_weights(all_data): 
    
    all_data_levels_two=list(all_data['Level 2'].unique()) 
    list_dates_all_position=list(all_data.columns[3:10])
    matrix_level_pos=np.zeros((len(all_data_levels_two),len(list_dates_all_position)))
    
    for dd in all_data_levels_two:
        
        #print(dd)
        
        index_dd=all_data_levels_two.index(dd)
        
        data_sub=all_data[all_data['Level 2']==dd]
        for pos in list_dates_all_position:
            #print(pos)
            index_pos=list_dates_all_position.index(pos)
            data_sub_two=data_sub[data_sub[pos].notna()]
            
            if len(data_sub_two)!=0:
                list_weights= list(data_sub_two['Weight'])
                #print("Weights",list_weights)
                list_date_delivers=list(data_sub_two['Date of Reviewing'])
                #print("Date of delivering",list_date_delivers)
                list_grades=list(data_sub_two[pos])
                combined = list(zip(list_date_delivers, list_weights, list_grades))
                combined.sort(key=lambda x: x[0], reverse=True)
                
                selected = []
                current_weight_sum = 0
                
                for item in combined:
                    if current_weight_sum <= 1:
                        selected.append(item)
                        current_weight_sum += item[1] 
                        
                        
                sorted_dates, sorted_weights, sorted_grades = zip(*selected)
                sorted_weights=np.array(list(sorted_weights))
                sorted_grades=np.array(list(sorted_grades))
                grades_weigh=(np.dot(sorted_weights,sorted_grades))/(np.sum(sorted_weights))
                #print(grades_weigh)

                
                if np.sum(sorted_weights) <=1 :
                    grades_weigh=(np.sum(sorted_weights))* grades_weigh 
                    
                matrix_level_pos[index_dd,index_pos]=grades_weigh
            else: 
                matrix_level_pos[index_dd,index_pos]=np.nan

    
    return matrix_level_pos




def combine_baseline_data(baseline_data,all_data,level_one_list,level_two_list):
    bas_data=pd.read_excel(baseline_data,sheet_name="CF-Baselining")

    bas_data_data_frame=bas_data.iloc[5:36]
    bas_data_col=list(bas_data.iloc[4])
    bas_data_col=bas_data_col[:3]
    bas_data_col[0]='Level 1'
    bas_data_col[1]='Level 2'
    bas_data_col[2]='Weight'
    bas_data_column_two=list(bas_data.iloc[3,3:11])
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
    bas_data_data_frame['Level 1']=level_one_list
    bas_data_data_frame['Level 2']=level_two_list
    
    all_data=pd.concat([all_data,bas_data_data_frame])
    
    return all_data

def get_top_values_avg(arr, n=4):
    # Sort the array in descending order
    sorted_arr = sorted(arr, reverse=True)
    
    # Get the top n values or all values if the array length is less than n
    top_values = sorted_arr[:min(len(arr), n)]
    
    avg_value=sum(top_values)/n
    return  avg_value





def calculate_psg_score(file_psg,file_path_ctf):
    
    #Read the PSG file
    df_psg=pd.read_excel(file_psg)
    matrix=df_psg.iloc[:,1:].to_numpy()
    
    #Read the CTF file
    df_ctf=pd.read_excel(file_path_ctf,sheet_name='Assessment')

    df_matrix=df_ctf.iloc[5:36,:9]

    column_one=list(df_ctf.iloc[4,:2])

    column_two=list(df_psg.columns[1:])

    column_total=column_one+column_two

    df_matrix.columns=column_total
    
    #Fill Level 1
    save_value_level=""
    for i in range(len(df_matrix)): 
        if pd.isnull(df_matrix['Level 1'].iloc[i]): 
            df_matrix['Level 1'].iloc[i]=save_value_level 
        else: 
            save_value_level=df_matrix['Level 1'].iloc[i]
            
            
    all_level_one=list(df_matrix['Level 1'].unique())

    aggregate_level=np.zeros(matrix.shape)
    
    for i,lev in enumerate(all_level_one):
        for j,psg in enumerate(column_two):
            
            
            df_sub=df_matrix[df_matrix['Level 1']==lev]
            scores=np.array(list(df_sub[psg]))
            scores =scores [~np.isnan(scores )]
            top_n=int(matrix[i,j])
            avg_value=get_top_values_avg(scores, n=top_n)
            
            
            aggregate_level[i,j]=avg_value
            
            
            
            
     
    aggregated_level_data_frame=pd.DataFrame(aggregate_level)
    aggregated_level_data_frame.columns=column_two
    aggregated_level_data_frame.index=all_level_one

    sum_psg_step_one=aggregate_level*matrix

    sum_factors_psg=np.sum(matrix,axis=0)
    sum_factors_aggregated_level=np.sum(sum_psg_step_one,axis=0)

    final_psg_score=np.divide(sum_factors_aggregated_level,sum_factors_psg)
    
    
    return final_psg_score





def calculate_psg_score_v2(weights,levels_one_two, all_psg_levels,file_psg):
    
    
        
    #Read the matrix for each psg and level
        
    df_psg=pd.read_excel(file_psg)
    
    
    column_two=list(df_psg.columns[1:])
    
    
    matrix=df_psg.iloc[:,1:].to_numpy()
    
    
    
    
    
    #Place the weights in a dataframe 
    
    psg_grade_data_frame=pd.DataFrame(weights)
    
    psg_grade_data_frame.columns=column_two
    
    psg_grade_data_frame['Level 2']=levels_one_two['Level 2']
    
    psg_grade_data_frame['Level 1']=levels_one_two['Level 1']
    
    col_all=list(levels_one_two.columns)
    
    col_all.extend(column_two)
    
    
    psg_grade_data_frame=psg_grade_data_frame[col_all]
    
    
    all_level_one=list(levels_one_two['Level 1'].unique())
    

    

    #Start with the agregation

    aggregate_level=np.zeros(matrix.shape)
    
    
    #Get the top n per PSG and Level
    
    
    for i,lev in enumerate(all_level_one):
        for j,psg in enumerate(column_two):
            
            
            df_sub=psg_grade_data_frame[psg_grade_data_frame['Level 1']==lev]
            
            
            scores=np.array(list(df_sub[psg]))
            scores =scores [~np.isnan(scores )]
            top_n=int(matrix[i,j])
            avg_value=get_top_values_avg(scores, n=top_n)
            
            
            aggregate_level[i,j]=avg_value
    
    
    
    aggregated_level_data_frame=pd.DataFrame(aggregate_level)
    aggregated_level_data_frame.columns=column_two
    aggregated_level_data_frame.index=all_level_one

    sum_psg_step_one=aggregate_level*matrix

    sum_factors_psg=np.sum(matrix,axis=0)
    sum_factors_aggregated_level=np.sum(sum_psg_step_one,axis=0)

    final_psg_score=np.divide(sum_factors_aggregated_level,sum_factors_psg)
    
    
    return final_psg_score





    
def level_one_aggregation(weights,levels_one_two, all_psg_levels,file_psg):
    
    
        
    #Read the matrix for each psg and level
        
    df_psg=pd.read_excel(file_psg)
    
    
    column_two=list(df_psg.columns[1:])
    
    
    matrix=df_psg.iloc[:,1:].to_numpy()
      
    
    #Place the weights in a dataframe 
    
    psg_grade_data_frame=pd.DataFrame(weights)
    
    psg_grade_data_frame.columns=column_two
    
    psg_grade_data_frame['Level 2']=levels_one_two['Level 2']
    
    psg_grade_data_frame['Level 1']=levels_one_two['Level 1']
    
    col_all=list(levels_one_two.columns)
    
    col_all.extend(column_two)
    
    
    psg_grade_data_frame=psg_grade_data_frame[col_all]
    
    
    all_level_one=list(levels_one_two['Level 1'].unique())
    

    

    #Start with the agregation

    aggregate_level=np.zeros(matrix.shape)
    
    
    #Get the top n per PSG and Level
    
    
    for i,lev in enumerate(all_level_one):
        for j,psg in enumerate(column_two):
            
            
            df_sub=psg_grade_data_frame[psg_grade_data_frame['Level 1']==lev]
            
            
            scores=np.array(list(df_sub[psg]))
            scores =scores [~np.isnan(scores )]
            avg_value=sum(scores)
            
            
            aggregate_level[i,j]=avg_value
    
    
    
    aggregated_level_data_frame=pd.DataFrame(aggregate_level)
    aggregated_level_data_frame.columns=column_two
    aggregated_level_data_frame.index=all_level_one


    
    
    return aggregated_level_data_frame
    
    
    
            
            
                
    
    
    
    
    
    











    


    
    
    
    

    
