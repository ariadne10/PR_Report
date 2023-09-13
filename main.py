import streamlit as st
import pandas as pd
import io

# Upload button for Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, skiprows=1)
    
    # Remove columns: AH, AI, AJ, AC, and AD
    columns_to_remove = ['AH', 'AI', 'AJ', 'AC', 'AD']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    
    # Delete all rows where values in column E contain "PAAS"
    df = df[df['E'] != 'PAAS']
    
    # Remove column E
    df.drop(columns=['E'], errors='ignore', inplace=True)
    
    # Delete all rows where value in column H contain "06-LE/FC-N"
    df = df[df['H'] != '06-LE/FC-N']
    
    # Show the resulting DataFrame
    st.write('Cleansed Data')
    st.write(df)

    # Download Button for cleansed DataFrame
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}">Download CSV File</a>'
    st.markdown(href, unsafe_allow_html=True)
