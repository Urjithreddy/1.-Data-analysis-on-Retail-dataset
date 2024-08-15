import pandas as pd

# Load the original Excel file
xlsx = pd.ExcelFile('Airport_data_cleaned.xlsx')

# Read data from the first sheet
df1 = pd.read_excel(xlsx, sheet_name=0)

# Read data from the second sheet
df2 = pd.read_excel(xlsx, sheet_name=1)

# Merge the two dataframes on 'Country Code' column
merged_df = pd.merge(df1, df2[['Country Code', 'Region', 'IncomeGroup']], on='Country Code', how='left')

# Replace 'Indicator Name' and 'Indicator Code' with 'Region' and 'IncomeGroup'
merged_df['Indicator Name'] = merged_df['Region']
merged_df['Indicator Code'] = merged_df['IncomeGroup']

# Drop unnecessary columns
merged_df.drop(['Region', 'IncomeGroup'], axis=1, inplace=True)

# Save the modified dataframe to a new Excel file
writer = pd.ExcelWriter('Airport_data_modified.xlsx', engine='xlsxwriter')
merged_df.to_excel(writer, index=False, sheet_name='Sheet1')
writer._save()

# Confirming the save
print("Data saved successfully as 'Airport_data_modified.xlsx'")
