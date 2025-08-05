import os
import openpyxl # noqa 401 //ruff linter
import pandas as pd

#############################
# SUPPLEMENTAL PATHS SCRIPT #
#############################
parent_dir = os.path.dirname(os.path.dirname(__file__))

# Read the user defined source files from: script_paths.xlsx file
path_src = os.path.join(parent_dir, 'PATHS_Template', 'script_paths.xlsx')

path_rd = pd.read_excel(path_src, sheet_name='Sheet1', engine='openpyxl')

# Transpose script_paths.xlsx dataFrame to dictionary for dynamic path retrieval
path_dict = path_rd.set_index('INDEX').T.to_dict('list')
# print(path_dict.keys())

# Function to create usable full path from the dictionary for subsequent dataframes
def get_full_path(index):
    path, file = path_dict[index]
    return os.path.join(path, file)

path_indexes = [
    'src_file_strF1',
    'src_file_strF2',
    'src_file_strF3',
    'src_file_strF4',
    # 'src_file_strF5', # disabled, Prepress now consolidated to one sheet
    'src_file_strF6',
    'src_file_strF7',
    'src_file_strF8',
    'src_file_strF9',
    'src_file_strF10',
    'src_file_strF11'
]

df_names = {
    'src_file_strF1':'df_srcF1',
    'src_file_strF2':'df_srcF2',
    'src_file_strF3':'df_srcF3',
    'src_file_strF4':'df_srcF4',
    # 'src_file_strF5':'df_srcF5',
    'src_file_strF6':'df_srcF6',
    'src_file_strF7':'df_srcF7',
    'src_file_strF8':'df_srcF8',
    'src_file_strF9':'df_srcF9',
    'src_file_strF10':'df_srcF10',
    'src_file_strF11':'df_srcF11'
}

for paths in path_indexes:
    try:
        src_file = get_full_path(paths)
        df = pd.read_excel(src_file, sheet_name='Master Timesheet', usecols='A:E', skiprows=7, nrows=350, engine='openpyxl')
        print(f"DataFrame for {paths}:")
        print(df) # Dynamically assign DataFrame to corresponding variable in df_names
        if paths in df_names:
            globals()[df_names[paths]] = df
    except KeyError as e:
        print(f"\n❌Error!!! Check 'INDEX' column names in the script_paths.xlsx file❌ {e}")
        print('')
    except FileNotFoundError as e:
        print("\n❌ Can not find file at path! ❌:")
        print(f"FileNotFoundError: {e}")
        print("=" * 71)
    except PermissionError as e:
        print("\n❌ Permission denied accessing file - check if open/locked by someone else ❌:")
        print(f"PermissionError: {e}")
        print("=" * 71)
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        print("\n❌ Error reading Excel file - may be empty or corrupted - check source and try again ❌:")
        print(f"{type(e).__name__}: {e}")
        print("=" * 71)
    except (ValueError, TypeError) as e:
        print("\n❌ Invalid parameter or data-type error: Similar to excel numbers starting with single quote acting as text ❌:")
        print(f"{type(e).__name__}: {e}")
        print("=" * 71)
    except OSError as e:
        print("\n❌ OS/File system error - Network drive or Windows PC problem ❌:")
        print(f"OSError: {e}")
        print("=" * 71)
    except Exception as e:  # Still keep this as a last resort, but log it differently
        print("\n❌ Unexpected error occurred - please report this to developer ❌:")
        print(f"{type(e).__name__}: {e}")
        print("Stack trace may be needed to diagnose this issue - developer")
        print("=" * 71)

#▽―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――---―――――――▽

if __name__ == "__main__":
    pass
    # supp_paths.py is imported to the other scripts to manage path & datafreame variables
    # not intended to be run standalone


#△――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――---――――――――――――――△