# -*- coding: utf-8 -*-
"""
Enhanced Employee Data Entry System
"""

import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO

# Define file paths
FILE_PATH = "employees.csv"
UPLOAD_FOLDER = "uploads"
ADMIN_PASSWORD = "admin123"  # Change this to your preferred password

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function to load employee data
def load_data():
    if os.path.exists(FILE_PATH):
        return pd.read_csv(FILE_PATH)
    else:
        return pd.DataFrame(columns=["Name", "Email", "Position"])

# Function to save data
def save_data(df):
    df.to_csv(FILE_PATH, index=False)

# Load existing data
df = load_data()

st.title("Employee Data Entry System")

# Select mode: Employee or Admin
mode = st.radio("Select Mode:", ("Employee", "Admin"))

if mode == "Employee":
    st.write("Enter employee details below:")
    
    # Employee input fields
    name = st.text_input("Employee Name")
    email = st.text_input("Email")
    position = st.text_input("Position")

    # Upload profile photo
    uploaded_file = st.file_uploader("Upload Profile Photo (JPG, PNG)", type=["jpg", "jpeg", "png"])

    # Button to add employee
    if st.button("Add Employee"):
        if name and email and position and uploaded_file:
            # Save uploaded file
            file_path = os.path.join(UPLOAD_FOLDER, f"{name.replace(' ', '_')}.jpg")
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Save employee data
            new_entry = pd.DataFrame([[name, email, position]], columns=df.columns)
            df = pd.concat([df, new_entry], ignore_index=True)
            save_data(df)
            
            st.success("Employee added successfully!")
        else:
            st.error("Please fill in all fields and upload a photo.")

elif mode == "Admin":
    st.write("### Admin Access (Restricted)")
    admin_password = st.text_input("Enter Admin Password", type="password")

    if st.button("Login"):
        if admin_password == ADMIN_PASSWORD:
            st.success("Access granted! You can now download employee data and photos.")

            # Create a zip file of all uploaded photos
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for file in os.listdir(UPLOAD_FOLDER):
                    file_path = os.path.join(UPLOAD_FOLDER, file)
                    zipf.write(file_path, os.path.basename(file_path))
            zip_buffer.seek(0)

            # Provide download links
            st.download_button(
                label="Download Employee Photos (ZIP)",
                data=zip_buffer,
                file_name="employee_photos.zip",
                mime="application/zip"
            )
            
            st.download_button(
                label="Download Employee Data (CSV)",
                data=df.to_csv(index=False).encode('utf-8'),
                file_name="employees.csv",
                mime="text/csv"
            )
        else:
            st.error("Incorrect password! Access denied.")

