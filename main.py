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

# File uploaders for the two different Excel files
uploaded_file = st.file_uploader("Choose the main Excel file", type="xlsx")
uploaded_file2 = st.file_uploader("Choose the 'S72 Sites and PICs' Excel file", type="xlsx")

if uploaded_file and uploaded_file2:
    df = pd.read_excel(uploaded_file, skiprows=1)
    df2 = pd.read_excel(uploaded_file2)

    # Initial data cleansing for the main file
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']
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

    # Reorder the columns
    col_names = df.columns.tolist()
    date_rel_index = col_names.index('Date Release')
    deliv_date_index = col_names.index('Delivery Date')
    col_names.insert(deliv_date_index, col_names.pop(date_rel_index))
    df = df[col_names]

    # More data manipulations
    df['Delivery Date'] = pd.to_datetime(df['Delivery Date']).dt.date
    df = df[df['3'] >= 500]
    df.loc[df['2'] == False, 'PR Qty'] = df['1']
    df.drop(columns=['1', '2', '3'], inplace=True)

    # Add the 'CONC' column
    df['CONC'] = df['Site Code'].astype(str) + df['BU Name'].astype(str)

    # Check for columns 12 and 13 in the 'S72 Sites and PICs' file
    try:
        remove_values = df2.loc[df2.iloc[:, 13] == '** Remove **', df2.columns[12]]
    except IndexError:
        st.write("Error: Columns 12 and 13 not found in 'S72 Sites and PICs' file.")
    else:
        # Remove rows from the main DataFrame based on CONC column
        df = df[~df['CONC'].isin(remove_values)]

    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
