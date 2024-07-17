import pandas as pd
import easygui
import os

# Open a file dialog to select two Excel files
files = easygui.fileopenbox(title="Select two Excel files", filetypes=["*.xlsx"], multiple=True)
if not files or len(files) < 2:
    raise ValueError("Two files must be selected.")

file1, file2 = files

# Load the two Excel files
df1 = pd.read_excel(file1, skiprows=2)  # Skip the first two rows
df2 = pd.read_excel(file2, skiprows=2)  # Skip the first two rows

# Assign a temporary name to the first unnamed column
df1.columns.values[0] = 'Time'
df2.columns.values[0] = 'Time'

# Merge dataframes on 'Time' with an outer join
merged_df = pd.merge(df1, df2, on='Time', how='outer')

# Fill NaN values forward and then backward
merged_df = merged_df.ffill().bfill()

# Generate the default output file name based on the first selected file's name
file1_name = os.path.splitext(os.path.basename(file1))[0]
default_output_file = f"{file1_name}_merged.xlsx"

# Open a file dialog to save the merged DataFrame to a new Excel file
output_file = easygui.filesavebox(title="Save merged Excel file", default=default_output_file, filetypes=["*.xlsx"])
if not output_file:
    raise ValueError("No file selected to save the merged output.")

# Ensure the file has the correct extension
if not output_file.endswith(".xlsx"):
    output_file += ".xlsx"

# Save the merged DataFrame
merged_df.to_excel(output_file, index=False)

print(f'Merged file saved as {output_file}')