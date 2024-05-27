import pandas as pd
import os

# Folder containing the Excel files
folder_path = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Totals\3 Unpivoted Tables'

# List to store DataFrames from each file and sheet
dfs = []

# Iterate through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):  # Consider only Excel files
        file_path = os.path.join(folder_path, filename)
        
        # Load the workbook
        excel_file = pd.ExcelFile(file_path)
        
        # Iterate through each sheet in the workbook
        for sheet_name in excel_file.sheet_names:
            # Read the sheet into a DataFrame without header
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            # Remove the file extension from the filename
            filename_without_extension = os.path.splitext(filename)[0]
            # Add a column with the source file name (without extension) and sheet name
            df['Source_File'] = filename_without_extension
            #df['Sheet_Name'] = sheet_name
            # Append the DataFrame to the list
            dfs.append(df)

# Concatenate all DataFrames into a single DataFrame
combined_df = pd.concat(dfs, ignore_index=True)

# Write the combined DataFrame to an Excel file without header
output_path = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Totals\4 Totals\Totals.xlsx'
combined_df.to_excel(output_path, index=False, header=False)
