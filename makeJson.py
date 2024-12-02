import pandas as pd
import json
import copy

# Load the master JSON (layer.json)
master_json = {}
with open("layer.json", "r") as json_file:
    master_json = json.load(json_file)

# Provide the path to your Excel file
excel_file_path = "TTPs.xlsx"

# Load the Excel file into a Pandas DataFrame (assuming only one sheet)
excel_data = pd.read_excel(excel_file_path, header=None, sheet_name=None)

# Initialize a dictionary to store entry frequencies and associated sheets
entry_freq = {}
ttp_sources = {}

# Iterate over each sheet in the Excel file
for sheet_name, sheet_data in excel_data.items():
    # Extract the first column (technique IDs) and second column (sources)
    known_techniques = sheet_data.iloc[:, 0].tolist()
    sources = sheet_data.iloc[:, 1].tolist()

    # Count the frequency of each technique (this part remains unchanged)
    for ttp, source in zip(known_techniques, sources):
        if ttp not in entry_freq:
            entry_freq[ttp] = 0
        entry_freq[ttp] += 1

        # Store the sources for each technique
        if ttp not in ttp_sources:
            ttp_sources[ttp] = []
        ttp_sources[ttp].append(f"{sheet_name}: {source}")

    # Filter techniques that are in the known techniques list
    copy_json = copy.deepcopy(master_json)
    filtered_techs = [t for t in copy_json["techniques"] if t["techniqueID"].lower() in known_techniques]

    # Iterate over the filtered techniques and assign scores
    for item in filtered_techs:
        technique_id = item["techniqueID"].lower()

        # Assign the score based on the frequency of the technique (1 or higher)
        if technique_id in entry_freq:
            item["score"] = 1

            # Add a comment only if the technique has a corresponding source (and a score of 1)
            if technique_id in ttp_sources:
                item["comment"] = '\n'.join(ttp_sources[technique_id])

    # Update the techniques in the copy of the master JSON
    copy_json["techniques"] = filtered_techs

    # Save the modified JSON to a file named after the sheet
    file_path = f"{sheet_name}.json"
    with open(file_path, "w") as op:
        json.dump(copy_json, op, indent=4)

    print(f"Processed sheet: {sheet_name}, saved to {file_path}")
