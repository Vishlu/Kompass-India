import pandas as pd


# Reload the original Excel file
file_path = r'Test_Upload_HunterIO.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1', engine='calamine')

# Explode the 'Executives and Designation' column into separate rows while keeping all other data intact
df = df.copy()
df['Executives and Designation'] = df['Executives and Designation'].fillna('')
df = df.assign(
    Executive_Designation=df['Executives and Designation'].str.split(', ')
).explode('Executive_Designation')

# Split the 'Executive_Designation' column into 'Executive' and 'Designation'
df[['Executive', 'Designation']] = df['Executive_Designation'].str.split(' - ', expand=True)

# Drop the temporary column
df = df.drop(columns=['Executives and Designation', 'Executive_Designation'])

# Save the result to a new Excel file
df.to_excel(r'stacked_executives_with_all_data_fixed.xlsx', index=False)

print("The data has been stacked correctly, retaining all other data, and saved to 'stacked_executives_with_all_data_fixed.xlsx'.")