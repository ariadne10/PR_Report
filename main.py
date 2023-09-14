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

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
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
    
    # Check if required columns exist
    required_columns = ['Total Dmnd', 'Net OH', 'PR Qty', 'Std Price', 'Delivery Date']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if len(missing_columns) == 0:
        # All required columns are present
        # Rest of your code
        col_index = df.columns.get_loc('Total Dmnd') + 1
        df.insert(col_index, "1", df['Total Dmnd'] - df['Net OH'])
        df.insert(col_index + 1, "2", df['1'] >= df['PR Qty'])
        df.insert(col_index + 2, "3", df['1'] * df['Std Price'])
        
        col_names = df.columns.tolist()
        date_rel_index = col_names.index('Date Release')
        deliv_date_index = col_names.index('Delivery Date')
        col_names.insert(deliv_date_index, col_names.pop(date_rel_index))
        df = df[col_names]
        
        df['Delivery Date'] = pd.to_datetime(df['Delivery Date']).dt.date
        
        st.markdown(get_table_download_link(df), unsafe_allow_html=True)
        
    else:
        # Some required columns are missing
        st.write(f"The following required columns are missing: {', '.join(missing_columns)}")
