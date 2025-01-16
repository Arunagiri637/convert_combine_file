import os
import pandas as pd
from openpyxl import Workbook

# Specify the folder where your .xls files are stored (use raw string to handle backslashes)
folder_path = r"C:\Users\PCM\Desktop\ssi\xls"  # Path where your .xls files are

# Specify the folder where you want to save the converted .xlsx files
converted_folder_path = r"C:\Users\PCM\Desktop\ssi\xlsx"  # Path for saving converted .xlsx files

# Specify the location for saving the combined file on the Desktop
desktop_path = r"C:\Users\PCM\Desktop\combined_data.xlsx"  # Combined data file on the Desktop

# Create the folder for converted files if it doesn't exist
if not os.path.exists(converted_folder_path):
    os.makedirs(converted_folder_path)

# Create a new empty list to store individual DataFrames
all_data_frames = []

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xls'):  # Only process .xls files
        file_path = os.path.join(folder_path, filename)

        # Read the .xls file into a pandas DataFrame
        data = pd.read_excel(file_path, engine='xlrd')

        # Append the DataFrame to the list
        all_data_frames.append(data)

        # Convert .xls file to .xlsx and save it in the converted folder
        xlsx_filename = filename.replace('.xls', '.xlsx')
        xlsx_file_path = os.path.join(converted_folder_path, xlsx_filename)

        # Save the DataFrame to an .xlsx file in the converted folder
        data.to_excel(xlsx_file_path, index=False, engine='openpyxl')

# Concatenate all DataFrames into one
combined_data = pd.concat(all_data_frames, ignore_index=True)

# Save the combined data to a new .xlsx file on the Desktop
combined_data.to_excel(desktop_path, index=False, engine='openpyxl')

print(
    "Conversion and combination completed successfully! Converted files saved in 'converted_xlsx_files' folder and combined data saved on Desktop.")
