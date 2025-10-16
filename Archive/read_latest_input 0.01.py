import os
import pandas as pd

# 1. Path to Input folder
input_folder = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input"

# 2. Get all .xlsx files in the folder
excel_files = [
    os.path.join(input_folder, f)
    for f in os.listdir(input_folder)
    if f.endswith('.xlsx')
]

# 3. Sort by last modified (newest first)
excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

# 4. Take only the 2 most recent
latest_files = excel_files[:2]

# 5. Read and preview each file
for file_path in latest_files:
    print(f"\n--- Reading file: {os.path.basename(file_path)} ---")
    try:
        df = pd.read_excel(file_path)
        print(df.head())  # Preview the top 5 rows
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
