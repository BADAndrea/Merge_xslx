import pandas as pd
import easygui
import os

# Open a file dialog to select multiple Excel files
files = easygui.fileopenbox(title="Select Excel files", filetypes=["*.xlsx"], multiple=True)
if not files or len(files) < 2:
    raise ValueError("At least two files must be selected.")

# Initialize an empty list to store DataFrames
dfs = []

# Load each selected Excel file into a DataFrame
for file in files:
    df = pd.read_excel(file, skiprows=2)  # Skip the first two rows
    df.columns.values[0] = 'Time'  # Assign a temporary name to the first unnamed column
    dfs.append(df)

# Merge all DataFrames on 'Time' with an outer join
merged_df = dfs[0]
for df in dfs[1:]:
    merged_df = pd.merge(merged_df, df, on='Time', how='outer')

# Fill NaN values forward and then backward
merged_df = merged_df.ffill().bfill()

# Generate the default output file name based on the first selected file's name
file1_name = os.path.splitext(os.path.basename(files[0]))[0]
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