import pandas as pd

# File paths
path_to_excel1 = '/Users/3590080/OneDrive - Jabil/PR Report/Rep - Var 1.xlsx'
path_to_excel2 = '/Users/3590080/OneDrive - Jabil/PR Report/S72 Sites and PICs.xlsx'
sheet_name = 'Sites VLkp'

# Read Excel files
df1 = pd.read_excel(path_to_excel1)
df2 = pd.read_excel(path_to_excel2, sheet_name=sheet_name)

# Data cleaning and transformation steps

# Remove unnecessary columns
columns_to_remove = [
    "Total Nettable On Hand", "Buyer Name", "Global Name", "Supplier Name",
    "Priority", "Action", "Part Type", "Part Profit Center Profit Center", "Supply Source", "Value"
]
df1 = df1.drop(columns=columns_to_remove)

# Remove rows where "paas" is in "Part Profit Center Profit Center"
df1 = df1[~df1['Part Profit Center Profit Center'].str.contains("paas", case=False)]

# Remove rows with specific conditions
df1 = df1[~((df1['Des'] == "Sched Agrmt") & (df1['Value'] == "06-LE/FC-N"))]
df1 = df1[df1['Supply Source'] != "SubstituteSupply"]

# Clean and transform the "Des" column
df1['Des'] = df1['Des'].apply(lambda x: "PlannedOrder" if ("Firm Planned Order" in x or pd.isna(x)) else x)

# Calculate the "T-N" column
df1['T-N'] = df1['Total Dmnd'] - df1['Net OH']

# Calculate the "500" column
df1['500'] = df1['T-N'] * df1['Target Price']

# Remove rows where "500" is less than 500
df1 = df1[df1['500'] >= 500]

# Update "PR Qty" and "Boolean" columns based on conditions
df1['Boolean'] = df1['T-N'] >= df1['PR Qty']
df1['Boolean'] = df1['Boolean'].fillna(False)

# Remove the '_' from the prefix of values in the "Site Code" column
df1['Site Code'] = df1['Site Code'].str.lstrip('_')

# Create a new "Concat" column by concatenating "Site Code" and "BU Name"
df1['Concat'] = df1['Site Code'] + df1['BU Name']

# Data merging and filtering
df1 = df1[~((df1['Concat'].isin(df2['CONC'])) & (df2['Action BU'] == "** Remove **"))]

# Add missing values from df1 to df2
df2 = pd.concat([df2, df1[['Site Code', 'BU Name']])

# Data filtering based on conditions
df1 = df1[~(df1['Part Number'].isin(df2['Delete']))]

# More data filtering based on conditions
df1 = df1[~((df1['BU Name'] == "CRESTRON") & (df1['Part Description'].str.contains("PROG")) & (df1['Mfr Part Code'].str.contains("(")))]
df1 = df1[~(df1['Manufacturer'].isin(["A & J PROGRAMMING", "MEXSER"]))]

# Drop "Part Description" column
df1 = df1.drop(columns=["Part Description"])

# More data filtering based on conditions
df1 = df1[~((df1['BU Name'] == "EMERSON") & (df1['Manufacturer'] == "ISSI"))]
df1 = df1[~((df1['BU Name'] == "EMERSONPM") & (df1['Manufacturer'] == "ISSI"))]
df1 = df1[~((df1['BU Name'] == "CADENCE") & (df1['Manufacturer'] == "TEXAS INSTRUMENTS"))]
df1 = df1[~((df1['BU Name'] == "GIGAMON") & (df1['Manufacturer'] == "BROADCOM"))]
df1 = df1[~(df1['Part Number'].str.startswith("FB"))]
df1 = df1[~((df1['BU Name'] == "ARISTA") & df1['Part Number'].str.endswith("B"))]
df1 = df1[~((df1['BU Name'].str.contains("HP")) & (df1['Site Code'] == "CN02") & ((df1['Commodity'] == "MEMNONVOL") | (df1['Commodity'] == "MEMVOL")))]
df1 = df1[~((df1['BU Name'] == "APBU") & (df1['Commodity'] == "SOLID STATE DRIVES"))]
df1 = df1[~((df1['BU Name'].str.contains("NETAPP")) & ((df1['Commodity'] == "HDD") | (df1['Commodity'] == "SOLID STATE DRIVES")))]
df1 = df1[~((df1['BU Name'].str.contains("CISCO")) & (df1['Manufacturer'] == "INTEL"))]
df1 = df1[~((df1['BU Name'].str.contains("CISCO")) & (df1['Part Number'].str.startswith("17-")))]
df1 = df1[~(df1['Site Code'] == "MX16")]
df1 = df1[~((df1['BU Name'] == "ADVANTEST") & (df1['Part Number'].str.contains("PP|EP")))]

# Apply formatting to specific columns
format_columns = ["Commodity", "Part Number"]
df1[format_columns] = df1[format_columns].style.set_properties(**{'background-color': 'orange'})

# Drop unnecessary columns
df1 = df1.drop(columns=["Net OH", "On Order"])

# Merge "Region" and "Site Name" columns based on matching "Site Code"
df1 = df1.merge(df2[['Site Code', 'Region', 'Site Name']], on="Site Code", how="left")

# Set font to Calibri size 8
# Note: This formatting is for display purposes only and won't be saved in the output file
df1 = df1.style.set_properties(**{'font-size': '8pt', 'font-family': 'Calibri'})

# Save the cleaned data as Internal file
# Note: You'll need to specify the desired file format and path for saving the file
# Example: df1.to_excel("internal_file.xlsx", index=False)

# Drop unnecessary columns for External file
df1 = df1.drop(columns=["Region", "BU Name", "Site Name", "Part Number"])

# Save the cleaned data as External file
# Note: You'll need to specify the desired file format and path for saving the file
# Example: df1.to_excel("external_file.xlsx", index=False)
