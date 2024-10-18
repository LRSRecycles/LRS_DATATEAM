#Copy and Paste in Visual studio .py or your python interpreter. 
#Make sure you have pandas and os packages installed before running the code.

import pandas as pd
import os

# Dynamically set the file path using os.path.join for platform independence
file_name = 'Dec 1 2024 PI Fix.xlsx'  # Excel file name. Change file_name to your file name.
directory = os.path.expanduser('~Downloads')  # Directory path (change if needed)
file_path = os.path.join(directory, file_name)  # Full file path

# Read all sheets from the Excel file into a dictionary where keys are sheet names and values are DataFrames
excel_data = pd.read_excel(file_path, sheet_name=None)

# Create a list of DataFrames after adding a 'SheetName' and 'SourceSheet' column for each sheet
df_list = [
    df.assign(SheetName=sheet_name, SourceSheet=sheet_name)  # Adding columns for identifying the source sheet
    for sheet_name, df in excel_data.items()  # Iterating over each sheet
    if not df.empty and not df.isna().all().all()  # Exclude empty or all-NA DataFrames
]

# Concatenate all DataFrames from the list into a single DataFrame, resetting the index
merged_df = pd.concat(df_list, ignore_index=True)

# Display the first few rows of the merged DataFrame to verify
print(merged_df)

# Save the merged DataFrame to a new Excel file in the same directory as the original file
output_excel = os.path.join(directory, 'Dec 1 2024 PI Fix Combined.xlsx')
merged_df.to_excel(output_excel, index=False)

# Save the merged DataFrame to a CSV file in the same directory as the original file
output_csv = os.path.join(directory, 'Dec 1 2024 PI Fix Combined.csv')
merged_df.to_csv(output_csv, index=False)
