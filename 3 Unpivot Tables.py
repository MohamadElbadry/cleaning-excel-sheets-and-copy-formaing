import os
import pandas as pd

# Function to unpivot Excel sheets and store results in a single workbook
def unpivot_excel_workbook(input_path, output_folder):
    # Read Excel file
    xls = pd.ExcelFile(input_path)
    
    # Create a new Excel writer object for the output workbook
    output_file = os.path.join(output_folder, os.path.splitext(os.path.basename(input_path))[0] + '_1.xlsx')
    with pd.ExcelWriter(output_file) as writer:
        # Iterate over each sheet in the Excel file
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            
            # Unpivot DataFrame
            df_unpivoted = pd.melt(df, id_vars=df.columns[0], var_name='Attribute', value_name='Value')
            
            # Write unpivoted DataFrame to the output workbook as a separate sheet
            df_unpivoted.to_excel(writer, index=False, sheet_name=sheet_name)

# Folder containing input Excel files
input_folder = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Totals\2 Tables Without Hearder'

# Folder to store output Excel files
output_folder = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Totals\3 Unpivoted Tables'

# Iterate over all Excel files in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx"):
        input_path = os.path.join(input_folder, filename)
        # Unpivot Excel sheets and store result in a single workbook
        unpivot_excel_workbook(input_path, output_folder)

print("Unpivoting complete.")
