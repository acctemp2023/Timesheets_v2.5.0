import time
import os
import pandas as pd
#pylint: disable=C0413 W0611

parent_dir = os.path.dirname(os.path.dirname(__file__))

# Used sequence numbering instead of per 'script_pasths.xlsx' file
sheet1 = pd.read_csv(os.path.join(parent_dir, 'Editorial Garth Group.csv'))
sheet2 = pd.read_csv(os.path.join(parent_dir, 'Editorial Luis Group.csv'))
sheet3 = pd.read_csv(os.path.join(parent_dir, 'Photo Timesheets.csv'))
sheet4 = pd.read_csv(os.path.join(parent_dir, 'Pre-Press.csv'))
sheet5 = pd.read_csv(os.path.join(parent_dir, 'Entertainment.csv'))
sheet6 = pd.read_csv(os.path.join(parent_dir, 'Prod-Dev-Baseball.csv'))
sheet7 = pd.read_csv(os.path.join(parent_dir, 'Prod-Dev-Basketball.csv'))
sheet8 = pd.read_csv(os.path.join(parent_dir, 'Prod-Dev-Soccer.csv'))
sheet9 = pd.read_csv(os.path.join(parent_dir, 'Prod-Dev-Football.csv'))

df_sheet1 = pd.DataFrame(sheet1)
df_sheet2 = pd.DataFrame(sheet2)
df_sheet3 = pd.DataFrame(sheet3)
df_sheet4 = pd.DataFrame(sheet4)
df_sheet5 = pd.DataFrame(sheet5)
df_sheet6 = pd.DataFrame(sheet6)
df_sheet7 = pd.DataFrame(sheet7)
df_sheet8 = pd.DataFrame(sheet8)
df_sheet9 = pd.DataFrame(sheet9)

suffixes = ['_Sh1', '_Sh2', '_Sh3', '_Sh4', '_Sh5', '_Sh6', '_Sh7', '_Sh8', '_Sh9']

# redefine apply_suffixes to keep the original 'Collection Code' column and rename others
def apply_suffixes(df, suffix):
    df[f'Collection Code{suffix}'] = df['Collection Code']
    df = df.rename(columns={
        'AX Project Code': f'AX Project Code{suffix}',
        'Collection Description': f'Collection Description{suffix}',
        'Sport': f'Sport{suffix}',
        'Total': f'Total{suffix}'
    })
    return df

# Apply suffixes to each DataFrame (remove .iloc[:450] to include all rows)
df_sheet1 = apply_suffixes(df_sheet1, suffixes[0])
df_sheet2 = apply_suffixes(df_sheet2, suffixes[1])
df_sheet3 = apply_suffixes(df_sheet3, suffixes[2])
df_sheet4 = apply_suffixes(df_sheet4, suffixes[3])
df_sheet5 = apply_suffixes(df_sheet5, suffixes[4])
df_sheet6 = apply_suffixes(df_sheet6, suffixes[5])
df_sheet7 = apply_suffixes(df_sheet7, suffixes[6])
df_sheet8 = apply_suffixes(df_sheet8, suffixes[7])
df_sheet9 = apply_suffixes(df_sheet9, suffixes[8])

# stack all DataFrames vertically
all_dfs = [df_sheet1, df_sheet2, df_sheet3, df_sheet4, df_sheet5, df_sheet6, df_sheet7, df_sheet8, df_sheet9]
# Concatenate all DataFrames into one
combined_df = pd.concat(all_dfs, axis=0, ignore_index=True)

# group by 'Collection Code' to consolidate rows (ensure all rows are included)
merged_df = combined_df.groupby('Collection Code', as_index=False).first()

# drop unwanted columns early to reduce memory usage
unwanted_cols = ['_Assembly', '_Imaging', '_Baseball', '_Basketball', '_Soccer', '_Football']
merged_df.drop(columns=[col for col in unwanted_cols if col in merged_df.columns], inplace=True, errors='ignore')

# sum the 'Total' columns across all suffixes
total_cols = [f'Total{suffix}' for suffix in suffixes]
merged_df['Total_Sum'] = merged_df[total_cols].sum(axis=1, numeric_only=True)
merged_df['Total_Sum'] = merged_df['Total_Sum'].apply(lambda x: f"{x:,.2f}")

# drop suffixed 'Collection Code' columns if not needed in output
collection_code_cols = [f'Collection Code{suffix}' for suffix in suffixes]
merged_df.drop(columns=collection_code_cols, inplace=True, errors='ignore')

# keep only the first 450 rows if needed
merged_df = merged_df.iloc[:450]  # Optional: Remove this line if all rows should be included

# define the base columns to consolidate
base_columns = ['AX Project Code', 'Collection Code', 'Collection Description', 'Sport']

# function to consolidate columns with suffixes into a single col
def consolidate_columns(df, column_name):
    # Find all columns that start with the base col name
    related_columns = [col for col in df.columns if col.startswith(column_name)]
    # Use coalesce to take the first non-NaN value across these columns
    df[column_name] = df[related_columns].bfill(axis=1).iloc[:, 0]
    # Drop the suffixed columns
    df.drop(columns=related_columns, inplace=True)

# Consolidate columns before selecting final columns
consolidate_columns(merged_df, 'AX Project Code')
consolidate_columns(merged_df, 'Collection Description')
consolidate_columns(merged_df, 'Sport')

# select final columns
final_columns = base_columns + ['Total_Sum']
merged_df = merged_df[final_columns]

# find & move the row for "Total"
total_row_index = merged_df[(merged_df.iloc[:, 0] == 'Total') | (merged_df.iloc[:, 1] == 'Total')].index
if not total_row_index.empty:
    # get row data
    total_row = merged_df.loc[total_row_index].copy()

    # drop row from df
    merged_df = merged_df.drop(total_row_index)

    # calculate new index placement
    new_index = total_row_index[0] + 3

    # bring row back into df at the new position
    upper_half = merged_df.iloc[:new_index]
    lower_half = merged_df.iloc[new_index:]
    merged_df = pd.concat([upper_half, total_row, lower_half]).reset_index(drop=True)

# write the final DataFrame to a CSV file
merged_df.to_csv(os.path.join(parent_dir, 'Omni_bookmerge_output.csv'), index=False)
time.sleep(0.25)

import fom_development_3 #noqa E402 F401

if __name__ == '__main__':
    pass
    # This script handles only the concat of all like kind row-wise 'Collection Codes'
    # before producing final excel workbook. It's not intended to be run standalone


