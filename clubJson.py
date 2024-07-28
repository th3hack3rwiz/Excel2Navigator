import pandas as pd
import json
import copy
from collections import Counter
import matplotlib.colors as mcolors

# Load the master JSON
master_json = {}
with open("layer.json", "r") as json_file:
    master_json = json.load(json_file)
master_json2 = copy.deepcopy(master_json)

# Provide the path to your Excel file
excel_file_path = "ActorTTPs.xlsx"

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
with open("layered.json", "w") as op:
    json.dump(master_json2, op, indent=4)
