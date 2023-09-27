import streamlit as st
import pandas as pd
import base64
import datetime
import re
import os

def get_excel_download_link(df):
    current_date = datetime.datetime.now().strftime('%m-%d-%y')
    
    # Ensure the directory exists
    directory = "/mnt/data/"
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    file_path = os.path.join(directory, "PPV_{}.xlsx".format(current_date))
    
    # Save DataFrame to Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    # ... [rest of the function remains unchanged]

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

   # Debug: Print available columns in main DataFrame
    st.write(f"Debug: Available columns in main DataFrame: {df.columns.tolist()}")
    st.write(f"Debug: Available columns in df2 DataFrame: {df2.columns.tolist()}")

     # Debug: Show unique values in the 'Action Part Number' column of df2
    st.write(f"Debug: Unique values in 'Action Part Number': {df2['Action Part Number'].unique()}")

    # Filter and remove rows based on 'BU Name', 'Part Description', and 'Mfr Part Code'
    df = df[~((df['BU Name'] == 'CRESTRON') 
              & df['Part Description'].str.contains("PROG", na=False) 
              & df['Mfr Part Code'].str.contains(r'\([^)]*\)', regex=True, na=False))]

    # Debug: Print the first few rows of the DataFrame after filtering
    st.write(f"Debug: First few rows of the filtered DataFrame based on 'BU Name' and 'Part Description':")
    st.write(df.head())

    # Remove rows if 'A & J PROGRAMMING' or 'MEXSER' is present under 'Manufacturer' column
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

    # Remove "Part Description" column
    df.drop(columns=['Part Description'], errors='ignore', inplace=True)
   
    # Debug: Show unique values in 'BU Name' and 'Manufacturer'
    st.write(f"Debug: Unique values in 'BU Name': {df['BU Name'].unique()}")
    st.write(f"Debug: Unique values in 'Manufacturer': {df['Manufacturer'].unique()}")

    # Remove "Part Description" column
    df.drop(columns=['Part Description'], errors='ignore', inplace=True)

    # Filter rows where 'BU Name' is 'EMERSON' or 'EMERSONPM' and 'Manufacturer' is 'INTEGRATED SILICON SOLUTIONS (ISSI)'
    rows_to_remove = df[
        (df['BU Name'].isin(['EMERSON', 'EMERSONPM'])) 
        & (df['Manufacturer'].str.contains('INTEGRATED SILICON SOLUTIONS \(ISSI\)', case=False, na=False))
    ].index
    
    # Debug: Show index values to be removed
    st.write(f"Debug: Index values to be removed: {rows_to_remove}")

    # Remove rows if they exist in the DataFrame
    if len(rows_to_remove) > 0:
        df.drop(rows_to_remove.intersection(df.index), inplace=True)

     # Debug: Show the number of rows removed
    st.write(f"Debug: Number of rows removed: {len(rows_to_remove)}")

        # Debug: Show unique values in the 'BU Name' and 'Manufacturer' columns
    st.write(f"Debug: Unique values in 'BU Name': {df['BU Name'].unique()}")
    st.write(f"Debug: Unique values in 'Manufacturer': {df['Manufacturer'].unique()}")

    # Remove rows where "BU Name" is "CADENCE" and "Manufacturer" is "TEXAS INSTRUMENT"
    rows_to_remove_cadence = df[
        (df['BU Name'] == 'CADENCE') 
        & (df['Manufacturer'] == 'TEXAS INSTRUMENT')
    ].index

    if len(rows_to_remove_cadence) > 0:
        df.drop(rows_to_remove_cadence, inplace=True)

    # Remove rows where "BU Name" is "GIGAMON" and "Manufacturer" is "BROADCOM"
    rows_to_remove_gigamon = df[
        (df['BU Name'] == 'GIGAMON') 
        & (df['Manufacturer'] == 'BROADCOM')
    ].index

    if len(rows_to_remove_gigamon) > 0:
        df.drop(rows_to_remove_gigamon, inplace=True)

    # Remove rows where "Part Number" starts with "FB"
    rows_to_remove_fb = df[df['Part Number'].str.startswith('FB', na=False)].index

    if len(rows_to_remove_fb) > 0:
        df.drop(rows_to_remove_fb, inplace=True)

    # Remove rows where "BU Name" is "ARISTA" and "Part Number" contains 'B'
    rows_to_remove_arista = df[
        (df['BU Name'] == 'ARISTA') 
        & df['Part Number'].str.contains('B', na=False)
    ].index

    if len(rows_to_remove_arista) > 0:
        df.drop(rows_to_remove_arista, inplace=True)

    # Debug: Show the number of rows removed for each condition
    st.write(f"Debug: Number of rows removed for CADENCE and TEXAS INSTRUMENT: {len(rows_to_remove_cadence)}")
    st.write(f"Debug: Number of rows removed for GIGAMON and BROADCOM: {len(rows_to_remove_gigamon)}")
    st.write(f"Debug: Number of rows removed for Part Number starting with FB: {len(rows_to_remove_fb)}")
    st.write(f"Debug: Number of rows removed for ARISTA and Part Number containing B: {len(rows_to_remove_arista)}")

    st.write(f"Debug: Available columns in the DataFrame: {df.columns.tolist()}")

    # Debug: Print available columns right after reading the Excel file
    st.write(f"Debug: Available columns right after reading Excel: {df.columns.tolist()}")

    # Additional Filtering
    remove_part_numbers = df2[df2['Action Part Number'] == 'Remove']['Eliminar de los PR Report'].tolist()

    # Debug: Show part numbers to remove
    st.write(f"Debug: Part numbers to remove: {remove_part_numbers}")

    df = df[~df['Part Number'].isin(remove_part_numbers)]

    # Debug: Show the number of rows removed
    st.write(f"Debug: Number of rows removed based on 'Eliminar de los PR Report': {len(remove_part_numbers)}")

    # Additional Filtering based on 'BU Name', 'Site Code' and 'Commodity'
    rows_to_remove_hp = df[
        (df['BU Name'] == 'HP') 
        & (df['Site Code'].str.contains('CN02', case=False, na=False))
        & (df['Commodity'].isin(['MEMVOL', 'MEMNONVOL']))
    ].index
    
    # Debug: Show index values to be removed for HP case
    st.write(f"Debug: Index values to be removed for HP: {rows_to_remove_hp}")

    # Remove rows if they exist in the DataFrame
    if len(rows_to_remove_hp) > 0:
        df.drop(rows_to_remove_hp, inplace=True)

    # Debug: Show the number of rows removed for HP case
    st.write(f"Debug: Number of rows removed for HP: {len(rows_to_remove_hp)}")

    # Debug: Print available columns right before the issue
st.write(f"Debug: Available columns right before the issue: {df.columns.tolist()}")

    # Remove rows where "BU Name" is "APBU" and "Commodity" is "SOLID STATE DRIVE"
rows_to_remove_apbu = df[
    (df['BU Name'] == 'APBU') 
    & (df['Commodity'] == 'SOLID STATE DRIVE')
].index

    # Debug: Show the number of rows to be removed for this condition
st.write(f"Debug: Number of rows to be removed for APBU and SOLID STATE DRIVE: {len(rows_to_remove_apbu)}")

if len(rows_to_remove_apbu) > 0:
    df.drop(rows_to_remove_apbu, inplace=True)
# Identify rows with the given conditions

rows_to_remove_netapp = df[
    df['BU Name'].isin(['NETAPP', 'NETAPPCTO', 'NETAPPFJ', 'NETAPPSMT']) 
    & df['Commodity'].isin(['HDD', 'SOLID STATE DRIVE'])
].index

# Debug: Show the number of rows to be removed for this condition
st.write(f"Debug: Number of rows to be removed for NETAPP conditions: {len(rows_to_remove_netapp)}")

# Remove rows if they exist in the DataFrame
if len(rows_to_remove_netapp) > 0:
    df.drop(rows_to_remove_netapp, inplace=True)

st.write(f"Debug: Available columns in df2 DataFrame: {df2.columns.tolist()}")

action_part_values = df2.get('Action Part Number')
if action_part_values is not None:
    st.write(f"Debug: Unique values in 'Action Part Number': {action_part_values.unique()}")
else:
    st.write("Column 'Action Part Number' not found in df2.")

if 'Action Part Number' in df2.columns:
    st.write(f"Debug: Unique values in 'Action Part Number': {df2['Action Part Number'].unique()}")
else:
    st.write("Column 'Action Part Number' not found in df2.")

# Identify rows with CISCO in BU Name and INTEL in Manufacturer
rows_to_remove_cisco_intel = df[
    (df['BU Name'] == 'CISCO') 
    & (df['Manufacturer'] == 'INTEL')
].index

# Debug: Show the number of rows to be removed for this condition
st.write(f"Debug: Number of rows to be removed for CISCO and INTEL conditions: {len(rows_to_remove_cisco_intel)}")

# Remove rows if they exist in the DataFrame
if len(rows_to_remove_cisco_intel) > 0:
    df.drop(rows_to_remove_cisco_intel, inplace=True)

# Identify rows with CISCO in BU Name and Part Number starting with 17-
rows_to_remove_cisco_part = df[
    (df['BU Name'] == 'CISCO') 
    & df['Part Number'].str.startswith('17-', na=False)
].index

# Debug: Show the number of rows to be removed for this condition
st.write(f"Debug: Number of rows to be removed for CISCO and Part Number starting with 17-: {len(rows_to_remove_cisco_part)}")

# Remove rows if they exist in the DataFrame
if len(rows_to_remove_cisco_part) > 0:
    df.drop(rows_to_remove_cisco_part, inplace=True)

# Remove specified columns
columns_to_remove = ['Net OH', 'On Order', 'Priority']
df.drop(columns=columns_to_remove, inplace=True, errors='ignore')

# If you want to provide an Excel download link with font and column width adjustments:
import openpyxl
from openpyxl.styles import Font

def get_excel_download_link(df):
    current_date = datetime.datetime.now().strftime('%m-%d-%y')
    
    # Save DataFrame to Excel
    with pd.ExcelWriter("/mnt/data/PPV_{}.xlsx".format(current_date), engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        
        # Get the workbook and the sheet names
        workbook  = writer.book
        worksheet = workbook.active
        
        # Set the font and size for all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(name='Calibri', size=8)
        
        # Adjust columns width
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    # Create download link
    b64 = base64.b64encode(open("/mnt/data/PPV_{}.xlsx".format(current_date), "rb").read()).decode()
    href = f'<a href="data:file/xlsx;base64,{b64}" download="PPV_{current_date}.xlsx">Download xlsx file</a>'
    return href

# Replace the CSV download link with Excel download link
st.markdown(get_excel_download_link(df), unsafe_allow_html=True)


    # Continue with the existing code to generate download link ...
st.markdown(get_table_download_link(df), unsafe_allow_html=True)
