import os
import shutil
# pylint: disable=W0105

"""
After the main script initailly scrapes & proceses network files for use in consolidated Excel WB;
- Move required files to a new directory
    note: the directory will overwrite any existing versons of itself each time it is generated

- Once move operation is complete, script will check for blank/missing cells in all target .csv files
  and then display offending rows in each .csv to the user

- The script will also call local module to validate existance of duplicate 'Collection Codes'
  if duplicates are found; mgnt will manually adjust source files on network.

- Re-run until no duplicates and no blanks exist.

- Lastly, Master_Codes_timesheets.csv will never have duplicate ['Collection Code'] entries, as it is the controlling document.
    Functionality of the Master collection code audit:
    1. Master_Codes_timesheets.csv file as the source of truth for approved ['Collection Code'] data for all departments .csv exports
    2. Report all instances of unauthorized collection codes to the manager in the terminal
    3. Show exact row locations and descriptions for each violation
"""

parent_dir = os.path.dirname(os.path.dirname(__file__))

# Define files to move && move them
file_2move = [
    os.path.join(parent_dir, 'Editorial Garth Group.csv'),
    os.path.join(parent_dir, 'Editorial Luis Group.csv'),
    os.path.join(parent_dir, 'Excel WB - Dept Hrs.xlsx'),
    os.path.join(parent_dir, 'Omni_bookmerge_output.csv'),
    os.path.join(parent_dir, 'Photo Timesheets.csv'),
    os.path.join(parent_dir, 'Pre-Press.csv'),
    os.path.join(parent_dir, 'Entertainment.csv'),
    os.path.join(parent_dir, 'Prod-Dev-Baseball.csv'),
    os.path.join(parent_dir, 'Prod-Dev-Basketball.csv'),
    os.path.join(parent_dir, 'Prod-Dev-Football.csv'),
    os.path.join(parent_dir, 'Prod-Dev-Soccer.csv'),
    os.path.join(parent_dir, 'ProdDev_Merged.csv'),
    os.path.join(parent_dir, 'Master_Codes_timesheets.csv')
]

# Norm target paths
target_path = os.path.normpath(os.path.join(parent_dir, 'Timesheet_Set'))
# chk target directory exists
os.makedirs(target_path, exist_ok=True)

# Move files to new folder 'Timesheet_Set' (emulating cut/paste)
for file in file_2move:
    # Normalize path & isolate needed file(s)
    file = os.path.normpath(file)
    file_name = os.path.basename(file)
    # set dest
    target_file = os.path.join(target_path, file_name)

    try:
        # Check source file(s) exists
        if not os.path.exists(file):
            print(f"Source file {file_name} does not exist. Skipping.")
            continue
        # Check source is a file (not dir)
        if not os.path.isfile(file):
            print(f"Source path {file_name} is not a file. Skipping.")
            continue
        # Check if target file exists
        if os.path.isfile(target_file):
            print(f"File {file_name} already exists in {target_path}. Overwriting.")
        # Mooov da file(s)
        shutil.move(file, target_file)
        print(f"File {file_name} moved to {target_path}")
    except (PermissionError, OSError, shutil.Error) as e:
        print(f"\n❌Error accessing or moving file '{file_name}'.\nPossible reasons:\nFile is open\nPermission denied\nFile system error\nMore Details: {e}")

print("▶ File move operation completed.")

file_to_validate = [
    os.path.join(target_path, 'Editorial Garth Group.csv'),
    os.path.join(target_path, 'Editorial Luis Group.csv'),
    os.path.join(target_path, 'Photo Timesheets.csv'),
    os.path.join(target_path, 'Pre-Press.csv'),
    os.path.join(target_path, 'Entertainment.csv'),
    os.path.join(target_path, 'Prod-Dev-Baseball.csv'),
    os.path.join(target_path, 'Prod-Dev-Basketball.csv'),
    os.path.join(target_path, 'Prod-Dev-Football.csv'),
    os.path.join(target_path, 'Prod-Dev-Soccer.csv')
]

# Check all sheets in `file_to_validate` for blank rows and show what we needs corrected
def check_for_blanks():
    import pandas as pd

    print('\n▢▢▢ Checking if Blank Cells in: CSV Files')
    print('.' * 55)

    all_issues = {}

    for file_path in file_to_validate:
        file_name = os.path.basename(file_path)
        print(f"Processing: {file_name}")
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            continue

        try:
            # Read the CSV file
            df = pd.read_csv(file_path)

            # Focus on specific columns if they exist
            columns_to_check = ['AX Project Code', 'Collection Code', 'Collection Description', 'Sport']
            available_cols = [col for col in columns_to_check if col in df.columns]

            if not available_cols:
                # If none of our desired columns exist, check all columns
                available_cols = df.columns.tolist()
            # Replace various forms of empty values with NA
            df = df.replace(r'^\s*$', pd.NA, regex=True)
            df = df.replace([0, '0', '', ' '], pd.NA)

            # Process each row
            file_issues = []
            for idx, row in df.iterrows():
                # Skip rows where all values are NaN
                if row[available_cols].isnull().all():
                    continue
                # Check each column in the row
                missing_cols = [col for col in available_cols if pd.isnull(row[col])]
                if missing_cols:
                    issue = {
                        # 'row': idx + 2,  # +2 for row locaiton per local snapshot excel workbook
                        'row': idx + 9,  # +9 is for row location per original network .xlsx
                        'invalid_cells': missing_cols
                    }
                    file_issues.append(issue)
            if file_issues:
                all_issues[file_name] = file_issues
        except (FileNotFoundError, pd.errors.EmptyDataError, pd.errors.ParserError, PermissionError, OSError) as e:
            print(f"\n❌Error processing file '{file_name}'.\nPossible reasons:\n- File not found\n- File is empty or corrupted\n- Permission denied\n- File is open or locked\n- File system error\nMore Details: {e}")

    # Display results
    if all_issues:
        print("\n■ Rows with BLANK cells:")
        for file_name, file_issues in all_issues.items():
            print(f"\nFile: {file_name}")
            for issue in file_issues:
                print(f"  Row {issue['row']}:")
                print(f"    Empty/Invalid cells: {', '.join(issue['invalid_cells'])}")
                print('.'*80)
                print('')
    else:
        print('')
        print(" ■ NO BLANK cells found in any files!")
        print('')
    return all_issues

def audit_dupes(file_path):
    import pandas as pd

    try:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return None
        file_name = os.path.basename(file_path)
        print('_'*86)
        print(f"■ {file_name}")

        # Read the CSV file
        df = pd.read_csv(file_path)

        if 'Collection Code' not in df.columns:
            print(f"Collection Code column not found in {file_name}")
            return None

        # Get the Collection Code column and drop NaN values
        base_col = df['Collection Code'].dropna()

        # Find duplicates and their indices, keeping first occurrence as valid
        duplicates_mask = base_col.duplicated(keep='first')
        duplicate_values = base_col[duplicates_mask]

        if not duplicate_values.empty:
            print('  DUPLICATES>>')
            # Group duplicates and show their locations
            dupe_info = []
            for value in duplicate_values.unique():
                rows = df[df['Collection Code'] == value].index.tolist()
                dupe_info.append({
                    'value': value,
                    # 'rows': [r + 2 for r in rows]  # +2 for row locaiton per local snapshot excel workbook
                    'rows': [r + 9 for r in rows]  # +9 is for row location per original network .xlsx
                })
            # Print detailed information
            for info in dupe_info:

                print(f"  Collection Code: {info['value']}")
                print(f"  Rows: {', '.join(map(str, info['rows']))}")
                print('')
            return dupe_info
        print("    NO DUPLICATES")
        print('')
        return []
    except (FileNotFoundError, pd.errors.EmptyDataError, pd.errors.ParserError, PermissionError, OSError) as e:
        print(f"\n❌Error processing file '{file_name}'.\nPossible reasons:\n- File not found\n- File is empty or corrupted\n- Permission denied\n- File is open or locked\n- File system error\nMore Details: {e}")
        return None

def check_for_duplicates():
    import pandas as pd
    print('\n▶ Checking for: DUPLICATE COLLECTION CODES')
    print('.'*86)
    all_dupes = {}

    for file_path in file_to_validate:
        dupes = audit_dupes(file_path)
        if dupes:
            file_name = os.path.basename(file_path)
            all_dupes[file_name] = dupes
    if all_dupes:
        print("\n>> Duplicate check complete.")
    else:
        print("\n>> Duplicate check complete.")
    return all_dupes

def get_offending_codes(master_file_path, review_file_path, columns_to_check=None):
    """
    Check if values in specified columns of the review files exist in the master file.

    Args:
        master_file_path: Path to the master codes file (source of truth)
        review_file_path: Path to the department file being validated
        columns_to_check: Dictionary of column names to check with their data types
    """
    import pandas as pd

    # Default to just checking Collection Code if no columns specified
    if columns_to_check is None:
        columns_to_check = {'Collection Code': str}

    try:
        print(f"Checking codes in: {os.path.basename(review_file_path)}")

        # Read master file with authorized codes
        df_mstr = pd.read_csv(master_file_path, dtype=columns_to_check)
        df_mstr = df_mstr.replace(r'^\s*$', pd.NA, regex=True)
        df_mstr = df_mstr.replace([0, '0', '', ' '], pd.NA)

        # Read the department file to review
        df_review = pd.read_csv(review_file_path, dtype=columns_to_check)
        df_review = df_review.replace(r'^\s*$', pd.NA, regex=True)
        df_review = df_review.replace([0, '0', '', ' '], pd.NA)

        unauthorized_entries = []

        # Check each column specified
        for column_name in columns_to_check:
            if column_name not in df_mstr.columns or column_name not in df_review.columns:
                print(f"Column '{column_name}' not found in one of the files. Skipping this column.")
                continue

            # Get authorized values from master file
            master_values = set(df_mstr[column_name].dropna())

            # Get review values with their row indices
            review_values = df_review[column_name].dropna()

            # Check each review value against master
            for idx, value in review_values.items():
                if value not in master_values:
                    unauthorized_entries.append({
                        'column': column_name,
                        'code': value,
                        'row': idx + 9,  # Add 2 for 0-indexed to 1-indexed conversion plus header row
                        'description': df_review.loc[idx, 'Collection Description'] if 'Collection Description' in df_review.columns else 'N/A'
                    })

        return unauthorized_entries
    except (FileNotFoundError, pd.errors.EmptyDataError, pd.errors.ParserError, PermissionError, OSError) as e:
        print(f"\n❌Error processing file '{os.path.basename(review_file_path)}'.\nPossible reasons:\n- File not found\n- File is empty or corrupted\n- Permission denied\n- File is open or locked\n- File system error\nMore Details: {e}")
        return None

def verify_collection_codes():
    import pandas as pd
    print('\n▪▪▪ Verifying Collection Codes Against Master List')
    print('_'*86)

    # Set up paths
    master_file_path = os.path.join(parent_dir, 'Timesheet_Set', 'Master_Codes_timesheets.csv')
    results_file_path = os.path.join(parent_dir, 'Timesheet_Set', 'approved_codes_result.csv')

    if not os.path.exists(master_file_path):
        print(f"Master codes file not found: {master_file_path}")
        return

    # Define all columns to check against master file
    columns_to_check = {
        'AX Project Code': str,
        'Collection Code': str,
        'Collection Description': str,
        'Sport': str
    }

    all_unauthorized = {}
    total_unauthorized = 0

    # Check each department file against the master codes
    for file_path in file_to_validate:
        file_name = os.path.basename(file_path)
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            continue

        unauthorized_entries = get_offending_codes(master_file_path, file_path, columns_to_check)

        if unauthorized_entries:
            all_unauthorized[file_name] = unauthorized_entries
            total_unauthorized += len(unauthorized_entries)    # Display results
    # Prepare data for CSV export
    results_data = []
    if all_unauthorized:
        print(f"\n>>> FOUND {total_unauthorized} unauthorized entries across all files:")
        print('_'*86)

        for file_name, entries in all_unauthorized.items():
            print(f"\n▶ File: {file_name}")
            columns_seen = {}
            for entry in entries:
                column = entry['column']
                code = entry['code']

                if column not in columns_seen:
                    columns_seen[column] = {}
                if code not in columns_seen[column]:
                    columns_seen[column][code] = []

                columns_seen[column][code].append(entry)
                # Add to results data
                results_data.append({
                    'File': file_name,
                    'Column': column,
                    'Unauthorized Value': code,
                    'Row': entry['row'],
                    'Description': entry['description']
                })

            for column in sorted(columns_seen.keys()):
                print(f"\n  - {column}")
                for code, code_entries in sorted(columns_seen[column].items()):
                    print(f"     Review Value: {code}")
                    for entry in code_entries:
                        print(f"      Row {entry['row']}: {entry['description']}")
                        print('')
    else:
        print("\n>>> No unauthorized entries found in any files")
        results_data.append({
            'File': 'All Files',
            'Column': 'N/A',
            'Unauthorized Value': 'N/A',
            'Row': 'N/A',
            'Description': 'No unauthorized entries found'
        })

    print('_'*86)

    # Export results to CSV
    if results_data:
        df_results = pd.DataFrame(results_data)
        df_results.to_csv(results_file_path, index=False)
        print(f"\nResults exported to: {results_file_path}")

    return all_unauthorized

# Call validation funcs when script is run
if __name__ == "__main__":
    # First run the blank check automatically after files are moved
    print("\n▶ Starting validation check: Blank Fields...")
    initial_issues = check_for_blanks()

    # Display summary of initial check
    if initial_issues:
        print("\n▶ Initial validation FOUND BLANK CELLS that need to be addressed.")
    else:
        print("\nInitial validation complete: NO BLANK cells found.")

    # Then show menu for user choices
    while True:
        print('▂' * 86)
        print("\n\n--- VALIDATION OPTIONS (per Dept) ---")
        # print("1: Continue if all blanks have been resolved")
        print("1: Check for blank fields again")
        print("2: Check Collection Code Duplications")
        print("3: Verify Collection Codes Against Master List")
        print("4: Exit")
        print('▂' * 86)

        choice = input("\nEnter your choice (1-4): ")

        if choice == '1':
            # Check for blanks
            print("\nStarting validation checks for blank fields...")
            issues = check_for_blanks()
            # Summary
            if issues:
                print("\nValidation complete. Issues were found.")
            else:
                print("\nValidation complete. No issues found.")
        elif choice == '2':
            # Check for duplicates
            dupes = check_for_duplicates()
        elif choice == '3':
            print("\n▶ Verify: APPROVED COLLECTION CODES BY DEPT")
            verify_collection_codes()
        elif choice == '4':
            print("\nExiting program...")
            break
        else:
            print("\nInvalid choice. Please select 1, 2, 3, or 4.")


