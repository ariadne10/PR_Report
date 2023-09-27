import streamlit as st
import pandas as pd
import base64
import datetime
import openpyxl
from openpyxl.styles import Font

def get_excel_download_link(df):
    current_date = datetime.datetime.now().strftime('%m-%d-%y')
    
    # Save DataFrame to Excel with styling
    file_path = "/mnt/data/PPV_{}.xlsx".format(current_date)
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        
        # Styling the Excel output
        workbook  = writer.book
        worksheet = workbook.active
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(name='Calibri', size=8)
        
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
    
    with open(file_path, "rb") as file:
        b64 = base64.b64encode(file.read()).decode()
    href = f'<a href="data:file/xlsx;base64,{b64}" download="PPV_{current_date}.xlsx">Download xlsx file</a>'
    return href

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

    # Data Cleansing for df
    df['Site Code'] = df['Site Code'].str.replace('_', '')
    df.drop(columns=['Priority', 'Part Type'], errors='ignore', inplace=True)
    
    # Create 'CONC' column
    df['CONC'] = df['Site Code'] + df['BU Name']

    # Remove rows based on 'S72 Sites and PICs' file
    remove_values = df2[df2['Action.1'] == '** Remove **']['CONC'].tolist()
    df = df[~df['CONC'].isin(remove_values)]

    # More data cleansing for df
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)

    # Debugging Information
    st.write(f"Debug: Available columns in main DataFrame: {df.columns.tolist()}")
    st.write(f"Debug: Available columns in df2 DataFrame: {df2.columns.tolist()}")

    # Display unique values in 'Action Part Number' column of df2
    if 'Action Part Number' in df2.columns:
        st.write(f"Debug: Unique values in 'Action Part Number': {df2['Action Part Number'].unique()}")

    # Filtering conditions for df
    df = df[~((df['BU Name'] == 'CRESTRON') 
              & df['Part Description'].str.contains("PROG", na=False) 
              & df['Mfr Part Code'].str.contains(r'\([^)]*\)', regex=True, na=False))]

    # Debugging Information
    st.write(f"Debug: First few rows of the filtered DataFrame based on 'BU Name' and 'Part Description':")
    st.write(df.head())

    # Further filtering
    df = df[~df['Manufacturer'].isin(['A & J PROGRAMMING', 'MEXSER'])]
    
    # Additional data cleansing
    df = df[df.get('Part Profit Center Profit Center') != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df.get('Value') != '06-LE/FC-N']

    # Additional cleansing
    des_mapping = {
        'Purch Req': 'Purch Req',
        'Sched Agrmt': 'Sched Agrmt',
        'Firm Planned Order': 'PlannedOrder'
    }
    for des, value in des_mapping.items():
        df.loc[df['Des'] == des, 'Supply Source'] = value
    df = df[df['Supply Source'] != 'SubstituteSupply']
    df.drop(columns=['Des', 'Value', 'Action', 'Indep Dmnd'], errors='ignore', inplace=True)

    # Date manipulation
    df = df.loc[pd.to_datetime(df.get('Date Release'), errors='coerce').notna()]
    df['Date Release'] = pd.to_datetime(df['Date Release'])
    df['Date Release1'] = df['Date Release'] - pd.to_timedelta(df.get('GRPT'), unit='D')
    df.drop(columns=['GRPT', 'Date Release'], errors='ignore', inplace=True)
    df.rename(columns={'Date Release1': 'Date Release'}, inplace=True)

    # Additional calculations
    df['1'] = df['Total Dmnd'] - df['Net OH']
    df['2'] = df['1'] >= df['PR Qty']
    df['3'] = df['1'] * df.get('Std Price')

    # Finalize the DataFrame
    df = df[df['3'] >= 500]
    df.loc[df['2'] == False, 'PR Qty'] = df['1']

    # Create and remove temporary columns
    df['20%'] = df.get('Std Price') * 0.2
    df['Difference'] = df.get('Std Price') - df['20%']
    df['Std Price'] = df['Difference']
    df.drop(columns=['Difference', '20%'], inplace=True, errors='ignore')
    df.rename(columns={'Std Price': 'Target Price'}, inplace=True)

    # Further column removal
    df.drop(columns=['Part Description'], errors='ignore', inplace=True)

    # Debugging Information
    st.write(f"Debug: Unique values in 'BU Name': {df['BU Name'].unique()}")
    st.write(f"Debug: Unique values in 'Manufacturer': {df['Manufacturer'].unique()}")

    # Filtering based on specific conditions
    rows_to_remove = df[
        (df['BU Name'].isin(['EMERSON', 'EMERSONPM'])) 
        & df['Manufacturer'].str.contains('INTEGRATED SILICON SOLUTIONS \(ISSI\)', case=False, na=False)
    ].index
    if len(rows_to_remove) > 0:
        df.drop(rows_to_remove, inplace=True)

    # Debugging Information
    st.write(f"Debug: Number of rows removed: {len(rows_to_remove)}")
    st.write(f"Debug: Unique values in 'BU Name': {df['BU Name'].unique()}")
    st.write(f"Debug: Unique values in 'Manufacturer': {df['Manufacturer'].unique()}")

    # Further filtering based on different conditions
    rows_to_remove_cadence = df[
        (df['BU Name'] == 'CADENCE') 
        & (df['Manufacturer'] == 'TEXAS INSTRUMENTS')
    ].index
    if len(rows_to_remove_cadence) > 0:
        df.drop(rows_to_remove_cadence, inplace=True)

    rows_to_remove_gigamon = df[
        (df['BU Name'] == 'GIGAMON') 
        & (df['Manufacturer'] == 'BROADCOM')
    ].index
    if len(rows_to_remove_gigamon) > 0:
        df.drop(rows_to_remove_gigamon, inplace=True)

    rows_to_remove_fb = df[df.get('Part Number').str.startswith('FB', na=False)].index
    if len(rows_to_remove_fb) > 0:
        df.drop(rows_to_remove_fb, inplace=True)

    rows_to_remove_arista = df[
        (df['BU Name'] == 'ARISTA') 
        & df.get('Part Number').str.contains('B', na=False)
    ].index
    if len(rows_to_remove_arista) > 0:
        df.drop(rows_to_remove_arista, inplace=True)

    # Debugging Information
    st.write(f"Debug: Number of rows removed for CADENCE and TEXAS INSTRUMENT: {len(rows_to_remove_cadence)}")
    st.write(f"Debug: Number of rows removed for GIGAMON and BROADCOM: {len(rows_to_remove_gigamon)}")
    st.write(f"Debug: Number of rows removed for Part Number starting with FB: {len(rows_to_remove_fb)}")
    st.write(f"Debug: Number of rows removed for ARISTA and Part Number containing B: {len(rows_to_remove_arista)}")

    # Further data cleansing steps
    remove_part_numbers = df2[df2.get('Action Part Number') == 'Remove'].get('Eliminar de los PR Report').tolist()
    df = df[~df.get('Part Number').isin(remove_part_numbers)]

    # Debugging Information
    st.write(f"Debug: Number of rows removed based on 'Eliminar de los PR Report': {len(remove_part_numbers)}")

    # Additional Filtering based on 'BU Name', 'Site Code' and 'Commodity'
    rows_to_remove_hp = df[
        (df['BU Name'] == 'HP') 
        & df.get('Site Code').str.contains('CN02', case=False, na=False)
        & df.get('Commodity').isin(['MEMVOL', 'MEMNONVOL'])
    ].index
    if len(rows_to_remove_hp) > 0:
        df.drop(rows_to_remove_hp, inplace=True)

    # Debugging Information
    st.write(f"Debug: Number of rows removed for HP: {len(rows_to_remove_hp)}")

    # Filtering based on specific conditions

# Debugging steps:
try:
    st.write(f"Columns in DataFrame df: {df.columns.tolist()}")
    st.write(f"Unique values in 'BU Name': {df['BU Name'].unique()}")
    st.write(f"Unique values in 'Manufacturer': {df['Manufacturer'].unique()}")
    
    # Now, the previous filtering condition:
    rows_to_remove_cisco_intel = df[
        (df['BU Name'] == 'CISCO') 
        & (df['Manufacturer'] == 'INTEL')
    ].index
    if len(rows_to_remove_cisco_intel) > 0:
        df.drop(rows_to_remove_cisco_intel, inplace=True)

except Exception as e:
    st.write(f"An error occurred: {e}")

# Remove rows where "BU Name" is "CISCO" and "Part Number" starts with "17-".
rows_to_remove_cisco_part = df[
    (df['BU Name'] == 'CISCO') 
    & df['Part Number'].str.startswith('17-', na=False)
].index
if len(rows_to_remove_cisco_part) > 0:
    df.drop(rows_to_remove_cisco_part, inplace=True)

# Remove rows where "BU Name" is one of ["NETAPP", "NETAPPCTO", "NETAPPFJ", "NETAPPSMT"] and "Commodity" is either "HDD" or "SOLID STATE DRIVE".
rows_to_remove_netapp = df[
    df['BU Name'].isin(['NETAPP', 'NETAPPCTO', 'NETAPPFJ', 'NETAPPSMT']) 
    & df['Commodity'].isin(['HDD', 'SOLID STATE DRIVE'])
].index
if len(rows_to_remove_netapp) > 0:
    df.drop(rows_to_remove_netapp, inplace=True)

    # Display download links for Excel and CSV
    st.markdown(get_excel_download_link(df), unsafe_allow_html=True)
    st.markdown(get_table_download_link(df), unsafe_allow_html=True)
