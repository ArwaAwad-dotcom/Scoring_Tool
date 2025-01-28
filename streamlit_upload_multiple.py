# -*- coding: utf-8 -*-
"""
Created on Fri Jan 24 13:16:22 2025

@author: user
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from analysis import get_relevant_data
from analysis import calculate_weights
from openpyxl import load_workbook
from analysis import combine_baseline_data





def process_files(uploaded_files):
    combined_data = pd.DataFrame()

    for file in uploaded_files:
        # Read each Excel file into a DataFrame
        df = pd.read_excel(file)
        # Perform some basic analysis (e.g., add a column with filename)
        df=get_relevant_data(file)
        combined_data = pd.concat([combined_data, df], ignore_index=True)

    return combined_data



def convert_df_to_excel(df):
    """Converts a DataFrame to an Excel file and returns it as a downloadable object."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed_Data')
    processed_file = output.getvalue()
    return processed_file


def save_workbook_to_bytes(workbook):
    output = BytesIO()
    workbook.save(output)
    output.seek(0)  # Reset the pointer to the beginning of the stream
    return output



ctf='Mindsets PMS_Comprehensive Talent Assessment Form_v03.xlsx'

# Streamlit App
st.title("Talent Assessment")

st.write("Upload all the EAF and Baseline Assessments for the talent.")

# Upload multiple files
uploaded_files = st.file_uploader("Upload Excel Files", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    st.success(f"{len(uploaded_files)} file(s) uploaded successfully!")
    
    # Process the files
    with st.spinner("Processing files..."):
        
        eafs=[]
        bas=""
        bool_bas=False
        
        for name in uploaded_files:
            print(name.name)
            if 'Engagement' in name.name:
                eafs.append(name)
            elif 'Baselining' in name.name:
                bas=name
                bool_bas=True
                

        
        combined_data = process_files(eafs)
        print(combined_data.head())
        if  bool_bas:
            combined_data=combine_baseline_data(bas,combined_data)
        

        matrix_data=calculate_weights(combined_data)
        
        workbook = load_workbook(ctf)
        sheet = workbook["Assessment"]

        start_row=7
        start_col=3

        for i, row in enumerate(matrix_data):
            for j, value in enumerate(row):
                sheet.cell(row=start_row + i, column=start_col + j).value = value



        
    
    # Display processed data
    st.write("Processed Data:")
    st.dataframe(combined_data)
    workbook_data = save_workbook_to_bytes(workbook)

    # Allow user to download the processed data
    processed_file = convert_df_to_excel(combined_data)

    st.download_button(
        label="Download CTF",
        data=workbook_data,
        file_name="CTF.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
