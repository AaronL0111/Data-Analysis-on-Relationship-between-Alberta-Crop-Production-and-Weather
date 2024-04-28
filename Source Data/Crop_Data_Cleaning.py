import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from sklearn.linear_model import LinearRegression

# Local Path to the Excel file for crop data
crop_file_path = r"C:\Users\Aaron\Desktop\Data Project\Report_G002_20240218_201419.xlsx"

# Read the crop Excel file into a pandas DataFrame
crop_df = pd.read_excel(crop_file_path, index_col=0)

# Convert float values in the second and third rows to strings and concatenate
crop_df.iloc[0] = crop_df.iloc[0].astype(str) + ' ' + crop_df.iloc[1].astype(str).fillna('').astype(str)

# Drop the third row
crop_df = pd.concat([crop_df.iloc[[0]], crop_df.iloc[2:]])

# Delete columns with 'nan nan' in the first row
for col in crop_df.columns:
    if crop_df.loc[crop_df.index[0], col] == 'nan nan':
        crop_df.drop(col, axis=1, inplace=True)

# Rename columns for crop data
for col_index in crop_df.columns:
    if col_index[:7] != "Unnamed":
        column_name = col_index[:4]
        next_row_value = crop_df.loc[crop_df.index[0], col_index]  # Use loc instead of iloc to access row by label
        new_col_index = column_name + ' ' + str(next_row_value)  # Convert next_row_value to string
        crop_df.rename(columns={col_index: new_col_index}, inplace=True)
    else:
        next_row_value = crop_df.loc[crop_df.index[0], col_index]  # Use loc instead of iloc to access row by label
        new_col_index = column_name + ' ' + str(next_row_value)  # Convert next_row_value to string
        crop_df.rename(columns={col_index: new_col_index}, inplace=True)

# Drop the first row (column names)
crop_df = crop_df.drop(crop_df.index[0])

# Define the output CSV file path
output_csv_path = r"C:\Users\Aaron\Desktop\Data Project\crop_data.csv"

# Output the crop DataFrame to a CSV file
crop_df.to_csv(output_csv_path, index_label='Index')