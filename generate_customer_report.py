import pandas as pd

# --- 1. Load the Data ---
try:
    df = pd.read_excel('CoilData.xlsx', sheet_name='Data')
except FileNotFoundError:
    print("Error: 'CoilData.xlsx' not found.")
    exit()

# --- 2. Group by Customer and Aggregate ---
# This is a single, powerful groupby operation.
# We calculate the sum of mass and the count of rows for each customer.
customer_summary = df.groupby('Customer Name').agg(
    TotalMass=('Mass (tons)', 'sum'),
    TotalCoils=('Order Number', 'size')  # 'size' efficiently counts all rows in the group
).reset_index()

# --- 3. Sort the report by Total Mass ---
customer_summary = customer_summary.sort_values(by='TotalMass', ascending=False)

# --- 4. Format and Save ---
# Rename columns for the final report
customer_summary.rename(columns={'TotalMass': 'Total Mass (tons)', 'TotalCoils': 'Total Coils'}, inplace=True)

print("--- Simple Summary Report by Customer (Sorted by Mass) ---")
print(customer_summary.to_string(index=False))

# Save to a new sheet named 'Simple_Customer_Summary'
with pd.ExcelWriter('StainlessSteel_Summary_Report.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    customer_summary.to_excel(writer, index=False, sheet_name='Simple_Customer_Summary')

print("\nReport successfully saved to sheet 'Simple_Customer_Summary' in 'StainlessSteel_Summary_Report.xlsx'")
