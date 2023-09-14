import streamlit as st
import pandas as pd
import base64
import datetime

def get_table_download_link(df):
    current_date = datetime.datetime.now().strftime('%m-%d-%y')
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="PPV_{current_date}.csv">Download csv file</a>'
    return href

# File upload for the main file
uploaded_file = st.file_uploader("Choose the main Excel file", type="xlsx")

# File upload for the 'S72 Sites and PICs' file
uploaded_file2 = st.file_uploader("Choose the 'S72 Sites and PICs' Excel file", type="xlsx")

if uploaded_file and uploaded_file2:
    # Reading the main file
    df = pd.read_excel(uploaded_file, skiprows=1)
    
    # Initial data cleansing
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']

    # Additional Cleansing Steps
    df.loc[df['Des'] == 'Purch Req', 'Supply Source'] = 'Purch Req'
    df.loc[df['Des'] == 'Sched Agrmt', 'Supply Source'] = 'Sched Agrmt'
    df.loc[df['Des'] == 'Firm Planned Order', 'Supply Source'] = 'PlannedOrder'
    df = df[df['Supply Source'] != 'SubstituteSupply']
    df.drop(columns=['Des', 'Value', 'Action', 'Indep Dmnd'], errors='ignore', inplace=True)

    df = df.loc[pd.to_datetime(df['Date Release'], errors='coerce').notna()]
    df['Date Release'] = pd.to_datetime(df['Date Release'])
    df['Date Release1'] = df['Date Release'] - pd.to_timedelta(df['GRPT'], unit='D')
    df.drop(columns=['GRPT', 'Date Release'], errors='ignore', inplace=True)
    df.rename(columns={'Date Release1': 'Date Release'}, inplace=True)

    # Add CONC column
    df['CONC'] = df['Site Code'].astype(str) + df['BU Name'].astype(str)
    
    # Reading the 'S72 Sites and PICs' file
    df2 = pd.read_excel(uploaded_file2)
    
    # Get CONC values that need to be removed
    remove_values = df2.loc[df2['Action'] == '** Remove **', 'CONC']
    
    # Remove rows from the main dataframe
    df = df[~df['CONC'].isin(remove_values)]
    
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
