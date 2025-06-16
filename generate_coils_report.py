import pandas as pd

# --- 1. Load the Data ---
# Assumes your data is in 'CoilData.xlsx' in a sheet named 'Data'
try:
    df = pd.read_excel('Coils data.xlsx', sheet_name='Coils')
except FileNotFoundError:
    print("Error: 'CoilData.xlsx' not found. Make sure the file is in the same directory.")
    exit()

# --- 2. First Grouping: Count coils per customer for each product size ---
# This is equivalent to the first 'Group By' in Power Query
counts_df = df.groupby(['Thickness', 'Width', 'Steel Type', 'Customer Name']).size().reset_index(name='CoilCount')

# --- 3. Create the formatted string ---
# This is like the 'Add Custom Column' step
# It creates strings like "Customer A (4)"
counts_df['CustomerSummary'] = counts_df['Customer Name'] + ' (' + counts_df['CoilCount'].astype(str) + ')'

# --- 4. Second Grouping: Combine the strings for each product size ---
# This is the final aggregation step, where we join the text strings
# The .agg() function is very powerful for this
final_report = counts_df.groupby(['Thickness', 'Width', 'Steel Type']).agg(
    CustomerSummary = ('CustomerSummary', lambda x: ', '.join(x))
).reset_index()


# --- 5. Display and Save the Report ---
print("--- Final Summary Report ---")
print(final_report.to_string())

# Save the final report to a new Excel file
# index=False prevents pandas from writing the DataFrame index as a column
final_report.to_excel('StainlessSteel_Summary_Report.xlsx', index=False, sheet_name='Summary')

print("\nReport successfully saved to 'StainlessSteel_Summary_Report.xlsx'")
