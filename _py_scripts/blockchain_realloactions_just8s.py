import os
import csv
import pandas as pd
#pylint: disable=W0718

#####################################################################
# Digital Ralloacations (aka "8 Codes") for blockchain related hours #
#####################################################################

parent_dir = os.path.dirname(os.path.dirname(__file__))

# Define department names to match the reference file exactly
timesheet_dir = os.path.join(parent_dir, 'Timesheet_Set')
departments = {
    'Editorial Garth': os.path.join(timesheet_dir, 'Editorial Garth Group.csv'),
    'Editorial Luis': os.path.join(timesheet_dir, 'Editorial Luis Group.csv'),
    'Photo': os.path.join(timesheet_dir, 'Photo Timesheets.csv'),
    'Pre-Press': os.path.join(timesheet_dir, 'Pre-Press.csv'),
    'PD-BK': os.path.join(timesheet_dir, 'Prod-Dev-Basketball.csv'),
    'PD-BB': os.path.join(timesheet_dir, 'Prod-Dev-Baseball.csv'),
    'PD-SC': os.path.join(timesheet_dir, 'Prod-Dev-Soccer.csv'),
    'PD-FB': os.path.join(timesheet_dir, 'Prod-Dev-Football.csv'),
    'Entertainment': os.path.join(timesheet_dir, 'Entertainment.csv')
}

path_tgt = os.path.join(parent_dir, 'Timesheet_Set', 'Blockchain_8Codes.csv')

# Get all unique Collection Codes that start with 8 from all departments
all_codes: list[str] = []

for dept, path in departments.items():
    try:
        print(f"Checking {dept} at {path}")
        if not os.path.exists(path):
            print(f"  File does not exist: {path}")
            continue

        df = pd.read_csv(path)
        print(f"  File loaded successfully. Shape: {df.shape}")
        print(f"  Columns: {list(df.columns)}")

        if 'Collection Code' in df.columns:
            # First, let's see how many codes contain '8'
            codes_with_8 = df[df['Collection Code'].astype(str).str.contains('8')]
            print(f"  Codes containing '8': {len(codes_with_8)}")

            # Filter codes that start with '8'
            codes = df[df['Collection Code'].astype(str).str.startswith('8')]
            print(f"  Found {len(codes)} codes starting with 8")

            if len(codes) > 0:
                print(f"  Sample codes starting with 8: {codes['Collection Code'].head().tolist()}")

            # Add all matching codes to the list (avoiding duplicates)
            for code in codes['Collection Code'].astype(str).values:
                if code not in all_codes:
                    all_codes.append(code)
        else:
            print(f"  'Collection Code' column not found in {path}")
    except Exception as e:
        print(f"Error reading codes from {path}: {e}")

print(f"Total unique codes found: {len(all_codes)}")
print(f"First few codes: {all_codes[:5]}")

# Sort the codes for consistent output
all_codes.sort()

# Create a dictionary to store department dataframes
dept_dfs = {}

# For each department, create a dataframe with the codes and their totals
for dept, path in departments.items():
    try:
        # Create a dataframe for this department with all standardized codes
        dept_df = pd.DataFrame({'Collection Code': all_codes, 'Total': 0})

        # Read the source file
        df = pd.read_csv(path)

        # If this department's file has both required columns
        if 'Collection Code' in df.columns and 'Total' in df.columns:
            # Process each code in our standardized code list
            for idx, code in enumerate(all_codes):
                # Try to find matching row in the source data
                # First try exact match
                matches = df[df['Collection Code'] == code]

                # If no match, try matching the number part before the CO
                if matches.empty and code.endswith('CO'):
                    code_num = code.replace('CO', '')
                    # Look for codes that start with the number part (like '850105-')
                    matches = df[df['Collection Code'].astype(str).str.startswith(code_num + '-')]

                # If we found a match, get the Total
                if not matches.empty:
                    total_value = matches['Total'].iloc[0]
                    dept_df.loc[idx, 'Total'] = total_value
                    # print(f"  Matched {code} with total: {total_value}")

        # Store this department's dataframe
        dept_dfs[dept] = dept_df
    except Exception as e:
        print(f"Error processing {path}: {e}")
        # Create an empty dataframe for this department
        dept_df = pd.DataFrame({'Collection Code': all_codes, 'Total': 0})
        dept_dfs[dept] = dept_df

# Create CSV with the exact format of the reference file
with open(path_tgt, 'w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f)

    # Write header row with department names and empty cells
    header_row = []
    for dept in departments:
        header_row.extend([dept, '', ''])
    header_row.extend(['Processing columns with 8 codes...', '', ''])
    writer.writerow(header_row)

    # Write subheader row with 'Collection Code' and 'Total' repeated
    subheader_row = []
    for _ in departments:
        subheader_row.extend(['Collection Code', 'Total', ''])
    subheader_row.extend(['Collection Code', 'Total', ''])
    writer.writerow(subheader_row)

    # Write data rows
    for code in all_codes:
        row = []
        total = 0
        for dept in departments:
            # Find the total for this department and code
            dept_df = dept_dfs[dept]
            dept_total = dept_df.loc[dept_df['Collection Code'] == code, 'Total'].values[0]
            row.extend([code, dept_total, ''])
            total += dept_total

        # Add the total column
        row.extend([code, total, ''])
        writer.writerow(row)

print(f"Saved result to {path_tgt}")
