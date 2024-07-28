import pandas as pd

# Load the Excel file
file_path = 'test.xlsx'
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
