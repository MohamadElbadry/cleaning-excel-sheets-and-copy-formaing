import pandas as pd
import os

def split_excel_workbook(file_path, split_word, output_folder):
    # Load the workbook
    xl = pd.ExcelFile(file_path)

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Iterate over each sheet in the workbook
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)

        # Initialize variables
        table_dict = {}
        current_table_data = None
        sheet_counter = 1
        max_sheet_name_length = 31

        # Iterate over the rows of the dataframe
        for index, row in df.iterrows():
            if row.iloc[0] == split_word:  # Check if the specific word is found using iloc for positional indexing
                if current_table_data is not None:  # Save the current table before starting a new one
                    table_dict[f"Table_{sheet_counter}"] = current_table_data
                    sheet_counter += 1
                current_table_data = pd.DataFrame([row], columns=df.columns)  # Create a new DataFrame for the next table
            elif current_table_data is not None:  # If we are in a table
                current_table_data = pd.concat([current_table_data, pd.DataFrame([row], columns=df.columns)], ignore_index=True)

        # Add the last table
        if current_table_data is not None:
            table_dict[f"Table_{sheet_counter}"] = current_table_data

        # Write the tables to new Excel files with subsheets
        output_file = f"{output_folder}/{sheet_name}.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            # Ensure at least one sheet is visible
            if not table_dict:
                writer.book.create_sheet()
            for table_name, table_data in table_dict.items():
                # Truncate sheet names if they exceed the maximum length
                truncated_sheet_name = table_name[:max_sheet_name_length]
                table_data.to_excel(writer, sheet_name=truncated_sheet_name, index=False)
            # Set the first sheet as active
            writer.book.active = 0

# Usage
file_path = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Final 2022_1.xlsx'
split_word = 'column'  # Specific word that marks the beginning of a new table
output_folder = r'E:\DH\My Work\نشرة اسعار اهم السلعة الغذائيه\Totals\2022\Totals\1 Tables from sheets'

split_excel_workbook(file_path, split_word, output_folder)