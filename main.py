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

    # Debug: Print available columns in main DataFrame
    st.write(f"Debug: Available columns in main DataFrame: {df.columns.tolist()}")

    # Additional Data Cleansing Steps
    df['Site Code'] = df['Site Code'].str.replace('_', '')
    df.drop(columns=['Priority', 'Part Type'], errors='ignore', inplace=True)

    # Create 'CONC' column
    df['CONC'] = df['Site Code'] + df['BU Name']

    # Remove rows based on 'S72 Sites and PICs' file
    remove_values = df2[df2['Action.1'] == '** Remove **']['CONC'].tolist()
    df = df[~df['CONC'].isin(remove_values)]

    # More data cleansing steps
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)

 # New requirements

# First, create a filtered DataFrame based on 'BU Name' and 'Part Description'
filtered_df = df[(df['BU Name'] == 'Crestron') & df['Part Description'].str.contains("PROG", na=False)]

# Identify rows in the filtered DataFrame that have "(" in 'Mfr Part Code'
rows_to_remove = filtered_df[filtered_df['Mfr Part Code'].str.contains(r"\(", regex=True, na=False)].index

# Remove those rows from the original DataFrame if any such rows exist
if len(rows_to_remove) > 0:
    df.drop(rows_to_remove, inplace=True)

# Remove rows where 'Manufacturer' is "A & J PROGRAMMING" or "MEXSER"
df = df[~df['Manufacturer'].isin(['A & J PROGRAMMING', 'MEXSER'])]

    # Additional data cleansing
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']

    # Additional cleansing
    df.loc[df['Des'] == 'Purch Req', 'Supply Source'] = 'Purch Req'
    df.loc[df['Des'] == 'Sched Agrmt', 'Supply Source'] = 'Sched Agrmt'
    df.loc[df['Des'] == 'Firm Planned Order', 'Supply Source'] = 'PlannedOrder'
    df = df[df['Supply Source'] != 'SubstituteSupply']
    df.drop(columns=['Des', 'Value', 'Action', 'Indep Dmnd'], errors='ignore', inplace=True)

    # Convert to datetime and create new columns
    df = df.loc[pd.to_datetime(df['Date Release'], errors='coerce').notna()]
    df['Date Release'] = pd.to_datetime(df['Date Release'])
    df['Date Release1'] = df['Date Release'] - pd.to_timedelta(df['GRPT'], unit='D')
    df.drop(columns=['GRPT', 'Date Release'], errors='ignore', inplace=True)
    df.rename(columns={'Date Release1': 'Date Release'}, inplace=True)
    
    # Additional calculations
    df['1'] = df['Total Dmnd'] - df['Net OH']
    df['2'] = df['1'] >= df['PR Qty']
    df['3'] = df['1'] * df['Std Price']

    # Finalize the DataFrame
    df = df[df['3'] >= 500]
    df.loc[df['2'] == False, 'PR Qty'] = df['1']

    # Create and remove temporary columns
    df['20%'] = df['Std Price'] * 0.2
    df['Difference'] = df['Std Price'] - df['20%']
    df['Std Price'] = df['Difference']
    df.drop(columns=['Difference', '20%'], inplace=True)
    df.rename(columns={'Std Price': 'Target Price'}, inplace=True)

    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
