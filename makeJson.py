import pandas as pd
import json
import copy

master_json = {}
with open("layer.json", "r") as json_file:
    master_json = json.load(json_file)
# Provide the path to your Excel file
excel_file_path = "PS, Financian, Govt targeting actors TTPs.xlsx"

# Load the Excel file into a Pandas DataFrame
excel_data = pd.read_excel(excel_file_path, header=None, sheet_name=None, skiprows=None)

# Initialize a dictionary to store entry frequencies and associated sheets
entry_freq = {}

# Iterate over each sheet in the Excel file
for sheet_name, sheet_data in excel_data.items():
    # Extract the first column (column index 0) as a Series
    copy_json = copy.deepcopy(master_json)
    known_techniques = sheet_data.iloc[:, 0].tolist()
    #print(known_techniques)
    filtered_techs = [t for t in copy_json["techniques"] if t["techniqueID"].lower() in known_techniques]
    copy_json["techniques"] = filtered_techs
    #print(filtered_techs)
    print(sheet_name)
    file_path = sheet_name + ".json"
    with open(file_path, "w") as op:
        json.dump(copy_json, op, indent=4)
