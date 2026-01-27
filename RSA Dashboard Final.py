import pandas as pd
# Set option once at the top of your script (after importing pandas)
pd.set_option('future.no_silent_downcasting', True)

#Cleaning Section

# Read all rows
df = pd.read_excel('(Test) RSA Report.xlsx', header=None)

# Find header row (look for 'Policy No.' or another unique column name)
for idx in range(len(df)):
    if 'Policy No.' in df.iloc[idx].values:
        # Set this row as header
        df.columns = df.iloc[idx]
        # Drop all rows before and including header
        df = df.iloc[idx+1:].reset_index(drop=True)
        break
    
# Convert to datetime first
df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d/%m/%Y', errors='coerce')

# Find the position of "วันที่" column
date_col_index = df.columns.get_loc('วันที่')

# Extract day, month, year
df.insert(date_col_index + 1, 'Day', df['วันที่'].dt.day.astype('Int64'))
df.insert(date_col_index + 2, 'Month', df['วันที่'].dt.month.astype('Int64'))
df.insert(date_col_index + 3, 'Year', df['วันที่'].dt.year.astype('Int64'))

# Convert to date only (removes timestamp)
df['วันที่'] = df['วันที่'].dt.date

# Delete completely blank rows
df = df.dropna(how='all')

# Reset index after deletion
df = df.reset_index(drop=True)

# Then use replace normally
df = df.replace('-', pd.NA)


# Extract from Policy No. and overwrite Policy Type completely
df['Policy Type'] = df['Policy No.'].str.extract(r'(A[CV]\d)', expand=False)

# Extract all letters after the last digit
df['จังหวัด ทะเบียนรถ'] = df['ทะเบียนรถ'].str.extract(r'\d([ก-๙a-zA-Z]+)$', expand=False)
df['จังหวัด ทะเบียนรถ'] = df['จังหวัด ทะเบียนรถ'].replace('กรุงเทพ', 'กรุงเทพมหานคร')