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

uploaded_file = st.file_uploader("Choose the main Excel file", type="xlsx")
uploaded_file2 = st.file_uploader("Choose the 'S72 Sites and PICs' Excel file", type="xlsx")

if uploaded_file and uploaded_file2:
    df = pd.read_excel(uploaded_file, skiprows=1)
    df2 = pd.read_excel(uploaded_file2, sheet_name='Sites VLkp')

    # Debug: Show column names to the user to debug
    st.write(f"Debug: Column names in the uploaded main file: {df.columns.tolist()}")
    st.write(f"Debug: Column names in the uploaded 'S72 Sites and PICs' file: {df2.columns.tolist()}")

    # Create the 'CONC' column by combining 'BU Name' and 'Site Code'
    df['CONC'] = df['BU Name'] + df['Site Code']

    # Initial data cleansing for df
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']
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

    # Additional calculations
    df['1'] = df['Total Dmnd'] - df['Net OH']
    df['2'] = df['1'] >= df['PR Qty']
    df['3'] = df['1'] * df['Std Price']

    # Remove values based on the second file
    remove_values = df2[df2['Action'] == '** Remove **']['CONC']
    df = df[~df['CONC'].isin(remove_values)]

    # Reordering columns to place 'Date Release' to the right of 'Supply Source' and 'CONC' to the right of 'BU Name'
    column_order = [col for col in df.columns if col not in ['Date Release', 'CONC']]
    index_date_release = column_order.index('Supply Source') + 1
    index_conc = column_order.index('BU Name') + 1
    column_order.insert(index_date_release, 'Date Release')
    column_order.insert(index_conc, 'CONC')
    df = df[column_order]

    # Display DataFrame and download link
    st.write(df)
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
