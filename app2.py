import pandas as pd

# Load the Excel file
file_path = "Volunteers_Hist.xlsx"  # Replace with your file path
sheet_name = "Sheet1"  # Replace with your sheet name
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Define the range of academic years
years = [f"{year}-{year+1}" for year in range(2011, 2025)]

# Create a new DataFrame to store consolidated results
consolidated_df = pd.DataFrame(columns=["Name"] + years)

# Group by name and aggregate years
for name, group in df.groupby("Name"):
    # Initialize a row with "No" for all years
    row = {"Name": name}
    for year in years:
        row[year] = "No"
    
    # Mark "Yes" for years the volunteer is present
    for year in group["Year"]:
        row[year] = "Yes"
    
    # Append the row to the consolidated DataFrame
    #consolidated_df = consolidated_df.append(row, ignore_index=True)
    consolidated_df = pd.concat([consolidated_df, pd.DataFrame([row])], ignore_index=True)

# Save the consolidated data to a new sheet in the same workbook
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    consolidated_df.to_excel(writer, sheet_name='Consolidated', index=False)

print("Consolidation complete. Check the 'Consolidated' sheet in your Excel file.")