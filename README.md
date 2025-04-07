import pandas as pd

# Load the 'Dirty 1' sheet from the input file
input_file = "Structured-Sales-Data-dirty.xlsx"  # Input file path
output_file = "Structured-Sales-clean.xlsx"      # Output file path

# Read the data
df_dirty = pd.read_excel(input_file, sheet_name="Dirty 1", header=None)

# Forward-fill hierarchical categories (Segment and Ship Mode)
df_dirty['Segment'] = df_dirty[0].where(df_dirty[0].str.contains('Segment>>', na=False)).ffill()
df_dirty['Ship Mode'] = df_dirty[0].where(df_dirty[0].str.contains('Ship Mode>>', na=False)).ffill()

# Remove irrelevant rows (headers and totals)
df_cleaned = df_dirty.dropna(subset=[2])  # Retain rows with valid Order IDs
df_cleaned = df_cleaned[~df_cleaned[0].str.contains('>>|Grand Total', na=False)]

# Rename columns to match the clean structure
df_cleaned.columns = ['Order ID', 'Sales', 'Segment', 'Ship Mode']

# Convert the 'Sales' column to numeric for consistency
df_cleaned['Sales'] = pd.to_numeric(df_cleaned['Sales'], errors="coerce")

# Rearrange columns to match the structure of 'Clean 1'
df_cleaned = df_cleaned[['Segment', 'Ship Mode', 'Order ID', 'Sales']]

# Save the cleaned data to a new file
df_cleaned.to_excel(output_file, sheet_name="Clean 1", index=False)

# Display the cleaned DataFrame
print("Cleaned Data:")
print(df_cleaned)
