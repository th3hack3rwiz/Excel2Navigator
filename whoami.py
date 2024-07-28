import json
import pandas as pd
from openpyxl import load_workbook

# Define the path to the Excel file and the sheet name
excel_file_path = "Actors, Result, Navigator Output.xlsx"
sheet_name = "Result"

# Read the Excel file
excel_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Create a dictionary to store variable-value pairs from the test.txt file
variable_dict = {}

# Read the test.txt file and parse variable-value pairs
with open("test.txt", "r") as test_file:
    test_data = test_file.readlines()
    for line in test_data:
        parts = line.strip().split(":")
        if len(parts) == 2:
            variable = parts[0].strip().upper()
            value = parts[1].strip()
            variable_dict[variable] = value

# Initialize a list to store the updated values for the "Techniques" column
updated_techniques = []
updated_tactics = []

# Parse JSON to get map from techniqueID to tactic:
tidToTacticMap = {}
with open("layer.json", "r") as json_file:
    parsed_json = json.load(json_file)
    techniques = parsed_json["techniques"]

    for technique in techniques:
        tidToTacticMap[technique["techniqueID"]] = " ".join(word.capitalize() for word in technique["tactic"].replace('-', ' ').split())


# Iterate through the rows of the Excel data
for index, row in excel_data.iterrows():
    # Get the value from "Technique ID" (Column A)
    value = row["Technique ID"]  # Replace "Technique ID" with the actual name of your Column A
    
    # Capitalize the value
    capitalized_value = value.upper()

    # Check if the capitalized value is in the dictionary (case-insensitive)
    if capitalized_value in variable_dict:
        updated_techniques.append(variable_dict[capitalized_value])
    else:
        updated_techniques.append("TTP Outdated In MITRE")

    # Check if this TechniqueID has a tactic
    if capitalized_value in tidToTacticMap:
        updated_tactics.append(tidToTacticMap[capitalized_value])
    else:
        updated_tactics.append("TTP Outdated In MITRE")
        

# Insert a new column "Techniques" (Column B) with the updated values
excel_data.insert(loc=1, column="Techniques", value=updated_techniques)
excel_data.insert(loc=1, column="Tactics", value=updated_tactics)


# Save the updated data to the same Excel file (overwrite it)
with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    excel_data.to_excel(writer, sheet_name=sheet_name, index=False)

print("Data has been updated in the Excel file.")
