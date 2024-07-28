#This program does exactly what attack navigator does, but better.
#It takes an excel file as input, extracts the first column from all the sheets
#and displayes TTPs in descending order of their frequency, along with actor names.

import pandas as pd
from openpyxl import load_workbook

# Provide the path to your Excel file
excel_file_path = "PS, Financian, Govt targeting actors TTPs.xlsx"

# Load the Excel file into a Pandas DataFrame
excel_data = pd.read_excel(excel_file_path, header=None, sheet_name=None, engine='openpyxl')

# Initialize a dictionary to store entry frequencies and associated sheets
entry_freq = {}

# Iterate over each sheet in the Excel file
for sheet_name, sheet_data in excel_data.items():
    # Extract the first column (column index 0) as a Series
    first_column = sheet_data.iloc[:, 0]

    # Count the frequency of each entry in the first column
    freq = first_column.value_counts().to_dict()

    # Update the entry_freq dictionary with the frequencies and associated sheets from the current sheet
    for entry, frequency in freq.items():
        if entry not in entry_freq:
            entry_freq[entry] = {'frequency': frequency, 'sheets': [sheet_name]}
        else:
            entry_freq[entry]['frequency'] += frequency
            entry_freq[entry]['sheets'].append(sheet_name)

# Sort the entries based on their frequencies in descending order
sorted_entries = sorted(entry_freq.items(), key=lambda x: x[1]['frequency'], reverse=True)

# Create a list to hold the data for the new sheet
new_sheet_data = []

# Populate the new_sheet_data list with the sorted entries
for entry, data in sorted_entries:
    frequency = data['frequency']
    sheets = ', '.join(data['sheets'])
    new_sheet_data.append([entry, frequency, sheets])

# Convert the new_sheet_data list into a DataFrame
new_sheet_df = pd.DataFrame(new_sheet_data, columns=['Technique ID', 'Score', 'Actor name'])

# Load the existing workbook
book = load_workbook(excel_file_path)

# Add the new sheet with the DataFrame data
with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
    writer._book = book
    new_sheet_df.to_excel(writer, sheet_name='Summary', index=False)

# Reorder the sheets to make the "Summary" sheet the first one
book._sheets.insert(0, book._sheets.pop(-1))
book.save(excel_file_path)

print("New sheet 'Summary' added as the first sheet in the Excel workbook successfully.")
