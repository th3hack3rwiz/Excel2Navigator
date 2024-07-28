import pandas as pd
import json
import copy
from collections import Counter
import matplotlib.colors as mcolors
from openpyxl import load_workbook

def clean_TTPs(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Initialize an empty list to store DataFrames for each sheet
    all_dfs = []

    # Iterate over each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        # Read the sheet into a DataFrame without header
        df = pd.read_excel(xls, sheet_name, header=None)

        # Group by the first column (TTPs) and aggregate the second column (Source) as a comma-separated string
        df[1] = df.groupby(0)[1].transform(lambda x: ', '.join(x))
        df.drop_duplicates(inplace=True)

        # Append the DataFrame to the list
        all_dfs.append(df)

    # Save the result back to the same Excel file with separate sheets
    with pd.ExcelWriter(file_path) as writer:
        for i, df in enumerate(all_dfs):
            df.to_excel(writer, sheet_name=xls.sheet_names[i], index=False, header=None)

    print(f"Processed and saved unique TTPs to {file_path}")

def create_summary_sheet(excel_file_path):
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

def clubjson(file_path):
    # Load the master JSON
    master_json = {}
    with open("layer.json", "r") as json_file:
        master_json = json.load(json_file)
    master_json2 = copy.deepcopy(master_json)

    # Load the Excel file into a Pandas DataFrame
    excel_data = pd.read_excel(file_path, header=None, sheet_name=None, skiprows=None)

    # Initialize a dictionary to store entry frequencies and associated sheets
    entry_freq = {}

    for sheet_name, sheet_data in excel_data.items():
        known_techniques = sheet_data.iloc[:, 0].tolist()

        freq = Counter(known_techniques)

        for k, v in freq.items():
            if k not in entry_freq:
                entry_freq[k] = 0

            entry_freq[k] += v

    # Determine the max value for color gradient
    max_value = max(entry_freq.values(), default=1)  # Avoid division by zero if empty

    # Generate the color gradient
    def generate_gradient(start_color, end_color, num_steps):
        if num_steps <= 1:
            return [start_color]
        
        cmap = mcolors.LinearSegmentedColormap.from_list("gradient", [start_color, end_color], N=num_steps)
        return [mcolors.to_hex(cmap(i / (num_steps - 1))) for i in range(num_steps)]

    num_colors = max_value  # Equal number of colors as maxValue
    colors = generate_gradient("#8ec843", "#ff6666", num_colors)  # Light green to medium red

    # Filter techniques and add scores
    known_techs = entry_freq.keys()
    filtered_techs = [t for t in master_json2["techniques"] if t["techniqueID"].lower() in known_techs]

    for item in filtered_techs:
        item["score"] = entry_freq[item["techniqueID"].lower()]

    # Update master_json2 with filtered techniques and new gradient
    master_json2["techniques"] = filtered_techs
    master_json2["gradient"] = {
        "colors": colors,
        "minValue": 1,
        "maxValue": max_value
    }

    # Write the updated JSON to a new file
    with open("allActorsClubbed.json", "w") as op:
        json.dump(master_json2, op, indent=4)

    print("Updated allActorsClubbed.json with new color gradient and scores.")

def main():
    file_path = 'ActorTTPs.xlsx'
    
    # Clean the TTPs
    clean_TTPs(file_path)
    
    # Update allActorsClubbed.json
    clubjson(file_path)
    
    # Create the summary sheet
    create_summary_sheet(file_path)

if __name__ == "__main__":
    main()
