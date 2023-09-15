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

# Check if 'CONC' exists in df before attempting to filter
if 'CONC' in df.columns:
    remove_values = df2['CONC'][df2['Action'] == '** Remove **']
    df = df[~df['CONC'].isin(remove_values)]
else:
    st.write(f"Warning: 'CONC' column not found in the main file. Skipping removal step.")

    # Remove rows where 'CONC' value is in remove_values
    df = df[~df['CONC'].isin(remove_values)]

    # Reorder columns
    reordered_columns = [
        'SomeColumn', 'Supply Source', 'Date Release',
        'AnotherColumn', 'BU Name', 'CONC'
    ]  # Replace 'SomeColumn' and 'AnotherColumn' with actual column names
    df = df[reordered_columns]

    # Debug: Show reordered column names to the user to debug
    st.write(f"Debug: Reordered column names in the main file: {df.columns.tolist()}")

    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
