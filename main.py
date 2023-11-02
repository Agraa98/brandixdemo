import pandas as pd
import os
import numpy as np

# Read the Excel file
input_file = 'G:/300 Level/MT325/Brandix/Quantity.xlsx'
df = pd.read_excel(input_file)

# Remove leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

# Replace spaces with underscores in column names
df.columns = df.columns.str.replace(' ', '_')

# Print the list of column names to verify
#print(df.columns)

# Extract size columns
#sizes = ['3XS', '2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL']

# Convert relevant columns to numeric, handling percentages
def convert_to_float(value):
    try:
        return float(value.rstrip('%'))
    except (ValueError, AttributeError):
        return value

def convert_percentage(value):
    if isinstance(value, str) and '%' in value:
        return float(value.rstrip('%')) / 100
    return value

df['CA_UNITS'] = df['CA_UNITS'].apply(convert_to_float)
df['US_UNITS'] = df['US_UNITS'].apply(convert_to_float)

# List of size columns
sizes_CA = ['3XS', '2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL']
sizes_US=['3XS.1','2XS.1', 'XS.1', 'S.1', 'M.1', 'L.1', 'XL.1', '2XL.1']

# Generate new columns for CA_UNITS
for size in sizes_CA:
    # Calculate new column name for CA
    ca_column_name = f'CA_{size}'
    result_df = np.ceil(df['CA_UNITS'] * df[size].apply(convert_percentage)).replace([np.inf, -np.inf, np.nan], 0).astype(int)
    df[ca_column_name] = result_df

# Generate new columns for US_UNITS
for size in sizes_US:
    # Extract the size without ".1"
    size_without_dot = size[:-2]

    # Calculate new column name for US
    us_column_name = f'US_{size_without_dot}'
    result_df = np.ceil(df['US_UNITS'] * df[size].apply(convert_percentage)).replace([np.inf, -np.inf, np.nan],
                                                                                     0).astype(int)
    df[us_column_name] = result_df

# Generate a new column for each size containing the sum of 'CA_' and 'US_' columns
for size in sizes_CA:
    sum_column_name = f'{size}_Sum'
    df[sum_column_name] = df[f'CA_{size}'] + df[f'US_{size}']

# Generate new columns for tolerance with "_Tolerance" suffix
tolerance_percentage = 2 / 100  # 2% tolerance
for size in sizes_CA:
    tolerance_column_name = f'{size}_Tolerance'
    df[tolerance_column_name] = np.ceil(df[f'{size}_Sum'] * (1 + tolerance_percentage))

# Determine the path of the input file
input_path = os.path.dirname(os.path.abspath(input_file))

# Overwrite the existing Excel file with the new data
output_file = os.path.join(input_path, 'output_file.xlsx')
df.to_excel(output_file, index=False, header=True)

