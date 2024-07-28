import pandas as pd
import json
import copy
from collections import Counter

master_json = {}
with open("layer.json", "r") as json_file:
    master_json = json.load(json_file)
master_json2 = copy.deepcopy(master_json)
# Provide the path to your Excel file
excel_file_path = "PS, Financian, Govt targeting actors TTPs.xlsx"

# Load the Excel file into a Pandas DataFrame
excel_data = pd.read_excel(excel_file_path, header=None, sheet_name=None, skiprows=None)

# Initialize a dictionary to store entry frequencies and associated sheets
entry_freq = {}

for sheet_name, sheet_data in excel_data.items():
    known_techniques = sheet_data.iloc[:, 0].tolist()

    freq = Counter(known_techniques)

    for k, v in freq.items():
        if k not in entry_freq:
            entry_freq[k] = 0

        entry_freq[k] += v


known_techs = entry_freq.keys()

filtered = filtered_techs = [t for t in master_json2["techniques"] if t["techniqueID"].lower() in known_techs]
print(filtered)
for item in filtered:
    item["score"] = entry_freq[item["techniqueID"].lower()]

print(filtered)
master_json2["techniques"]= filtered

with open ("layered.json", "w") as op:
    json.dump(master_json2, op, indent=4)
