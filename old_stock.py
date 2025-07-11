
import pandas as pd
from datetime import datetime

# Read excel files
mb52 = pd.read_excel('MB52.xlsx')
mb51 = pd.read_excel('MB51.xlsx')

# Format dates in MB51
mb51['Posting Date'] = pd.to_datetime(mb51['Posting Date'], format='%d.%m.%Y')
mb51 = mb51.dropna(subset=['Posting Date'])

# Today's date
today = pd.Timestamp.now()

# Last movement date
last_movement = mb51.groupby('Material')['Posting Date'].max().reset_index()
last_movement.rename(columns={'Posting Date': 'Last Movement Date'}, inplace=True)

# Merge last movement date with MB52
df = pd.merge(mb52, last_movement, on='Material', how='left')

# Calculate days since last movement
df['Last Movement Date'] = pd.to_datetime(df['Last Movement Date'], errors='coerce')
df['Aging(Days)'] = (today - df['Last Movement Date']).dt.days

# Categorize aging
def aging_category(days):
    if pd.isna(days):
        return 'No Movement'
    elif days <= 30:
        return '0-30 Days'
    elif days <= 60:
        return '31-60 Days'
    elif days <= 90:
        return '61-90 Days'
    else:
        return '91+ Days'
    
df['Aging Category'] = df['Aging(Days)'].apply(aging_category)

# Grouped by aging category
summary = df.groupby('Aging Category')['Unrestricted'].sum().reset_index()

# Export to Excel
df.to_excel('Aging_Report.xlsx', index=False)
summary.to_excel('Aging_Summary.xlsx', index=False)

