import pandas as pd

# Load the Excel file
file_path = "Volunteers_Hist.xlsx"  # Replace with your file path
sheet_name = "Sheet1"  # Replace with your sheet name
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Normalize names: convert to lowercase and remove trailing periods
df['Name'] = df['Name'].str.lower().str.replace('.', '', regex=False)
df['Name'] = df['Name'].str.title()

# Group by normalized names and aggregate years into a list
consolidated_df = df.groupby('Name')['Year'].apply(list).reset_index()

# Save the consolidated data to a new sheet in the same workbook
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    consolidated_df.to_excel(writer, sheet_name='Consolidated', index=False)

print("Consolidation complete. Check the 'Consolidated' sheet in your Excel file.")