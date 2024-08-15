import pandas as pd

# Load the original Excel file
xls = pd.ExcelFile('Airport_data.xls')

# Load data from the first sheet, skipping the first three rows
Airport_data = pd.read_excel(xls, sheet_name=0, skiprows=3)

# Iterate over each column
for column in Airport_data.columns:
    # Count non-empty values in the column
    non_empty_count = Airport_data[column].count()
    # Check if only one row is filled and others are empty
    if non_empty_count == 0:
        # Drop the column
        Airport_data.drop(column, axis=1, inplace=True)

# Iterate over each row
for index, row in Airport_data.iterrows():
    # Count non-empty values in the row
    non_empty_count = row.count()
    # Check if only three columns are filled and others are empty
    if non_empty_count == 4:
        # Drop the row
        Airport_data.drop(index, inplace=True)

# Create a writer object for the new Excel file
writer = pd.ExcelWriter('Airport_data_cleaned.xlsx', engine='xlsxwriter')

# Write the first cleaned DataFrame to the new Excel file
Airport_data.to_excel(writer, index=False, sheet_name='Sheet1')

# Iterate over each remaining sheet in the original Excel file
for sheet_name in xls.sheet_names[1:]:
    # Read the data from the current sheet
    data = pd.read_excel(xls, sheet_name=sheet_name)
    # Write the data to the new Excel file without any modifications
    data.to_excel(writer, index=False, sheet_name=sheet_name)

# Save the new Excel file
writer._save()

# Confirming the save
print("Data saved successfully as 'Airport_data_cleaned.xlsx'")
