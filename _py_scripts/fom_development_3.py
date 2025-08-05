import os
import time
import pandas as pd
# pylint: disable=C0413 W0611

##########################################
# FULL OUTER MERGE - PRODUCT DEVELOPMENT #
##########################################

parent_dir = os.path.dirname(os.path.dirname(__file__))

df_srcF6 = pd.read_csv(os.path.join(parent_dir, 'Entertainment.csv'))
df_srcF7 = pd.read_csv(os.path.join(parent_dir,'Prod-Dev-Baseball.csv'))
df_srcF8 = pd.read_csv(os.path.join(parent_dir,'Prod-Dev-Basketball.csv'))
df_srcF9 = pd.read_csv(os.path.join(parent_dir,'Prod-Dev-Soccer.csv'))
df_srcF10 = pd.read_csv(os.path.join(parent_dir,'Prod-Dev-Football.csv'))

# Start Merge operations
#▽――――――――――――――――――――――――――――――――――――――――---――――――――――――――――――――――――――――――――――▽

#region

# Prepare DataFrames by renaming Total columns
df_srcF6 = df_srcF6.rename(columns={'Total': 'Total_F6'})
df_srcF7 = df_srcF7.rename(columns={'Total': 'Total_F7'})
df_srcF8 = df_srcF8.rename(columns={'Total': 'Total_F8'})
df_srcF9 = df_srcF9.rename(columns={'Total': 'Total_F9'})
df_srcF10 = df_srcF10.rename(columns={'Total': 'Total_F10'})

# Merge DataFrames while preserving metadata
merge_cols = ['Collection Code', 'AX Project Code', 'Collection Description', 'Sport']
merged_df = df_srcF6.merge(df_srcF7, on=merge_cols, how='outer')
merged_df = merged_df.merge(df_srcF8, on=merge_cols, how='outer')
merged_df = merged_df.merge(df_srcF9, on=merge_cols, how='outer')
merged_df = merged_df.merge(df_srcF10, on=merge_cols, how='outer')

merged_df = merged_df.iloc[:350]
time.sleep(0.25)

# Sum the 'Total' columns
merged_df['Total'] = merged_df[['Total_F6', 'Total_F7', 'Total_F8', 'Total_F9', 'Total_F10']].sum(axis=1, numeric_only=True)
time.sleep(0.25)

# Fill NaN values in Total columns with 0 for summation
total_cols = ['Total_F6', 'Total_F7', 'Total_F8', 'Total_F9', 'Total_F10']
merged_df[total_cols] = merged_df[total_cols].fillna(0)

# Sum the 'Total' columns
merged_df['Total'] = merged_df[total_cols].sum(axis=1)

# Rename columns for final output
merged_df = merged_df.rename(columns={
    'Total_F6': '_Entertainment',
    'Total_F7': '_Baseball',
    'Total_F8': '_Basketball',
    'Total_F9': '_Soccer',
    'Total_F10': '_Football'
})
sport = merged_df.pop('Sport')
total_sum = merged_df.pop('Total')
merged_df.insert(3, 'Sport', sport)
merged_df.insert(4, 'Total', total_sum)


total_row_index = merged_df[(merged_df.iloc[:, 0] == 'Total') | (merged_df.iloc[:, 1] == 'Total')].index
if not total_row_index.empty:
    # Get row data
    total_row = merged_df.loc[total_row_index].copy()
    # Drop row from df
    merged_df = merged_df.drop(total_row_index)
    # Calc new index place
    new_index = total_row_index[0] + 3
    # Bring row back at the new position
    upper_half = merged_df.iloc[:new_index]
    lower_half = merged_df.iloc[new_index:]
    merged_df = pd.concat([upper_half, total_row, lower_half]).reset_index(drop=True)

#endregion

#△―――――――――――――――――――――――――――――――――――――――――――---―――――――――――――――――――――――――――――――△
                                                        # End Merge operations



# Formatting
#▽――――――――――――――――――――――――――――――――――――――――――――――---――――――――――――――――――――――――――――▽
# Drop rows where all col values are NaN/null
merged_df = merged_df.dropna(how='all')

#  - hard code output to 300 rows
merged_df = merged_df.iloc[:350]

# Write to excel
merged_df.to_csv(os.path.join(parent_dir, 'ProdDev_Merged.csv'), index=False)

# Local .py module & latency mitigation
time.sleep(0.25)

import make_wbxl

#△―――――――――――――――――――――――――――――――――――――――――――――――――---―――――――――――――――――――――――――△
                                                            # End Formatting



if __name__ == '__main__':
    pass
