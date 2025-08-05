import os
import pandas as pd

parent_dir = os.path.dirname(os.path.dirname(__file__))
src_path = os.path.join(parent_dir, 'Timesheet_Set', 'approved_codes_result.csv')

df = pd.read_csv(src_path)

# Group by unauthorized value and create separate columns for each file
grouped = df.groupby(['Column', 'Unauthorized Value', 'Description']).agg({
    'File': lambda x: list(x),
    'Row': lambda x: list(x)
}).reset_index()

# Create a clean dataframe with files as separate columns
result_rows = []
for _, row in grouped.iterrows():
    files = row['File']
    rows = row['Row']

    # Create base row data
    base_data = {
        'Column': row['Column'],
        'Unauthorized Value': row['Unauthorized Value'],
        'Description': row['Description']
    }

    # Add each file as a separate column with its corresponding row number
    for i, (file, row_num) in enumerate(zip(files, rows)):
        file_col = f'*File_{i+1}'
        row_col = f'Rows_{i+1}'
        base_data[file_col] = file
        base_data[row_col] = row_num

    result_rows.append(base_data)

# Convert to DataFrame
result_df = pd.DataFrame(result_rows)

# Fill NaN values with empty strings for cleaner display
result_df = result_df.fillna('')

# Save the grouped results to a new CSV file
output_path = os.path.join(parent_dir, 'Timesheet_Set', 'approved_codes_result_GRP.csv')
result_df.to_csv(output_path, index=False)

print(f"Original rows: {len(df)}")
print(f"Reduced rows: {len(result_df)}")
print(f"Grouped data saved to: {output_path}")
print("\nCleaned data preview:")
print(result_df.head())