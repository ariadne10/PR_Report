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

# Upload the first file
uploaded_file = st.file_uploader("Choose the first Excel file", type="xlsx")

# Upload the second file
uploaded_file2 = st.file_uploader("Choose the second Excel file", type="xlsx", key='file2')

if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=1)
    
    # Initial data cleansing
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
    
    # Position 'Date Release' next to 'Supply Source'
    cols = df.columns.tolist()
    cols.insert(cols.index('Supply Source') + 1, cols.pop(cols.index('Date Release')))
    df = df[cols]

    # Additional data processing
    df['1'] = df['Total Dmnd'] - df['Net OH']
    df['2'] = df['1'] >= df['PR Qty']
    df['3'] = df['1'] * df['Std Price']
    df = df[df['3'] >= 500]
    df.loc[df['2'] == False, 'PR Qty'] = df['1']
    df.drop(columns=['1', '2', '3'], inplace=True)
    df['20%'] = df['Std Price'] * 0.2
    df['Difference'] = df['Std Price'] - df['20%']
    df['Std Price'] = df['Difference']
    df.drop(columns=['Difference', '20%'], inplace=True)
    df.rename(columns={'Std Price': 'Target Price'}, inplace=True)

    # Add CONC column to the right of 'BU Name'
    df.insert(df.columns.get_loc('BU Name') + 1, 'CONC', df['Site Code'].astype(str) + df['BU Name'].astype(str))
    
    # Reorder 'Date Release' column to the right of 'Supply Source'
    date_release_idx = df.columns.get_loc('Date Release')
    supply_source_idx = df.columns.get_loc('Supply Source')
    cols = list(df.columns)
    cols.insert(supply_source_idx + 1, cols.pop(date_release_idx))
    df = df[cols]
    
    if uploaded_file2:
        df2 = pd.read_excel(uploaded_file2)
        
        # Debug: Show column names and shape of the DataFrame
        if df2 is not None:
            st.write(f"Debug: Column names in the second uploaded file: {df2.columns.tolist()}")
            st.write(f"Debug: Shape of the second uploaded file: {df2.shape}")
        
        try:
            remove_values = df2.iloc[:, 12][df2.iloc[:, 13] == '** Remove **']
            df = df[~df['CONC'].isin(remove_values)]
        except IndexError:
            st.write("Index Error: The second uploaded file may not have enough columns.")
    
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
