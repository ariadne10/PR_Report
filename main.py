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
    if not set(required_columns).issubset(df.columns):
        missing_columns = set(required_columns) - set(df.columns)
        st.write(f"The following required columns are missing: {', '.join(missing_columns)}")
    else:
        # Data manipulations
        df['1'] = df['Total Dmnd'] - df['Net OH']
        df['2'] = df['1'] >= df['PR Qty']
        df['3'] = df['1'] * df['Std Price']

        # Reorder the columns
        col_names = df.columns.tolist()
        date_rel_index = col_names.index('Date Release')
        deliv_date_index = col_names.index('Delivery Date')
        col_names.insert(deliv_date_index, col_names.pop(date_rel_index))
        df = df[col_names]
        
        # Remove time from "Delivery Date"
        df['Delivery Date'] = pd.to_datetime(df['Delivery Date']).dt.date

        # Remove rows where column "3" is less than 500
        df = df[df['3'] >= 500]
        
        # Update "PR Qty" where column "2" is FALSE
        df.loc[df['2'] == False, 'PR Qty'] = df['1']

        # Remove columns "1", "2", "3"
        df.drop(columns=['1', '2', '3'], inplace=True)

        # Add new columns "20%" and "Difference"
        df['20%'] = df['Std Price'] * 0.2
        df['Difference'] = df['Std Price'] - df['20%']

        # Replace "Std Price" with "Difference"
        df['Std Price'] = df['Difference']
        
        # Drop the "Difference" and "20%" columns
        df.drop(columns=['Difference', '20%'], inplace=True)

        # Rename "Std Price" to "Target Price"
        df.rename(columns={'Std Price': 'Target Price'}, inplace=True)

# Inserting two new columns "CONC" and "VLOOKUP" to the right of "BU Name"
try:
    bu_name_index = df.columns.get_loc('BU Name') + 1
    df.insert(bu_name_index, "CONC", df['Site Code'].astype(str) + df['BU Name'].astype(str))
    df.insert(bu_name_index + 1, "VLOOKUP", "")
except KeyError:
    st.write("Column 'BU Name' not found in the DataFrame.")

# Display DataFrame to debug
st.dataframe(df)

        # Generate download link
        st.markdown(get_table_download_link(df), unsafe_allow_html=True)
