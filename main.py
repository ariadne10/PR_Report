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
    columns_to_remove = ['Buyer Name', 'Global Name', 'Supplier Name', 'Total Nettable On Hand', 'Net Req']
    df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    df = df[df['Part Profit Center Profit Center'] != 'PAAS']
    df.drop(columns=['Part Profit Center Profit Center'], errors='ignore', inplace=True)
    df = df[df['Value'] != '06-LE/FC-N']
    df.loc[df['Des'] == 'Purch Req', 'Supply Source'] = 'Purch Req'
    df.loc[df['Des'] == 'Sched Agrmt', 'Supply Source'] = 'Sched Agrmt'
    df.loc[df['Des'] == 'Firm Planned Order', 'Supply Source'] = 'PlannedOrder'
    df = df[df['Supply Source'] != 'SubstituteSupply']
    df.drop(columns=['Des', 'Value', 'Action'], errors='ignore', inplace=True)

    try:
        df['Date Release'] = pd.to_datetime(df['Date Release'])
        df['Date Release1'] = df['Date Release'] - pd.to_timedelta(df['GRPT'], unit='D')
    except Exception as e:
        problematic_values = df.loc[pd.to_datetime(df['Date Release'], errors='coerce').isna(), 'Date Release']
        st.write("Found problematic values in 'Date Release':", problematic_values)
        st.write("Error:", str(e))
        raise e

    st.markdown(get_table_download_link(df), unsafe_allow_html=True)

