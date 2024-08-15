import pandas as pd

# Load the data, skipping the first three rows
Airport_data = pd.read_excel('Airport_data.xls', skiprows=3)

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
    if non_empty_count == 3:
        # Drop the row
        Airport_data.drop(index, inplace=True)

# Save the modified DataFrame as a new Excel file (.xlsx format)
Airport_data.to_excel('Airport_data_cleaned.xlsx', index=False)

# Confirming the save
print("Data saved successfully as 'Airport_data_cleaned.xlsx'")
