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
    df2 = pd.read_excel(uploaded_file2)

    # Initial data cleansing
    df.drop(columns=['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req'], errors='ignore', inplace=True)
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']
    df.loc[df['Des'] == 'Purch Req', 'Supply Source'] = 'Purch Req'
    df.loc[df['Des'] == 'Sched Agrmt', 'Supply Source'] = 'Sched Agrmt'
    df.loc[df['Des'] == 'Firm Planned Order', 'Supply Source'] = 'PlannedOrder'
    df = df[df['Supply Source'] != 'SubstituteSupply']
    df.drop(columns=['Des', 'Value', 'Action', 'Indep Dmnd'], errors='ignore', inplace=True)
    df = df.loc[pd.to_datetime(df['Date Release'], errors='coerce').notna()]
    df['Date Release1'] = df['Date Release'] - pd.to_timedelta(df['GRPT'], unit='D')
    df.drop(columns=['GRPT', 'Date Release'], errors='ignore', inplace=True)
    df['1'] = df['Total Dmnd'] - df['Net OH']
    df['2'] = df['1'] >= df['PR Qty']
    df['3'] = df['1'] * df['Std Price']

    # Reorder columns
    cols = df.columns.tolist()
    df['Date Release'] = df['Date Release1']
    cols.remove('Date Release1')
    cols.insert(cols.index('Supply Source') + 1, 'Date Release')
    df = df[cols]

    # Additional calculations
    df['3'] = df['1'] * df['Std Price']
    df['Delivery Date'] = pd.to_datetime(df['Delivery Date']).dt.date
    df = df[df['3'] >= 500]
    df.loc[df['2'] == False, 'PR Qty'] = df['1']
    df.drop(columns=['1', '2', '3'], inplace=True)
    df['20%'] = df['Std Price'] * 0.2
    df['Difference'] = df['Std Price'] - df['20%']
    df['Std Price'] = df['Difference']
    df.drop(columns=['Difference', '20%'], inplace=True)
    df.rename(columns={'Std Price': 'Target Price'}, inplace=True)

    # Adding CONC column next to 'BU Name'
    df['CONC'] = df['Site Code'].astype(str) + df['BU Name'].astype(str)
    cols = df.columns.tolist()
    cols.remove('CONC')
    cols.insert(cols.index('BU Name') + 1, 'CONC')
    df = df[cols]

    # Read second file to get rows to remove
    remove_values = df2.iloc[:, 12][df2.iloc[:, 13] == '** Remove **']
    
    # Remove the corresponding rows from the first DataFrame
    df = df[~df['CONC'].isin(remove_values)]

    # Generate and display download link
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
