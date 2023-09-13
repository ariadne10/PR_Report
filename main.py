import streamlit as st
import pandas as pd
import io
import base64

# Upload button for Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, skiprows=1)
    
    # Remove columns: AH, AI, AJ, AC, and AD
    columns_to_remove = ['AH', 'AI', 'AJ', 'AC', 'AD']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    
    # Delete all rows where values in column E contain "PAAS"
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    
    # Remove column E
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    
    # Delete all rows where value in column H contain "06-LE/FC-N"
    df = df[df['Value'] != '06-LE/FC-N']
    
    # Show the resulting DataFrame
    st.write('Cleansed Data')
    st.write(df)
    
    # Export to CSV (as a Download Link)
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
