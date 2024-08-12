import pandas as pd

# Source file path
source_file = 'origin.xlsx'  # Replace with your source file name
# Destination file path
destination_file = 'musachi.xlsx'  # Replace with your destination file name

# Step 1: Read the data from the source file using pandas
df = pd.read_excel(source_file, sheet_name=None)  # Read all sheets into a dictionary

# Step 2: Create a new Excel file using xlsxwriter
with pd.ExcelWriter(destination_file, engine='xlsxwriter') as writer:
    for sheet_name, data in df.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data copied from {source_file} to {destination_file}")