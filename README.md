## Overview:

The Timesheets application is a Python script designed to streamline the process of manually auditing, consolidating, and distributing labor of various departments
which will ultimately have the hours redistributed appropriately per the Payroll Allocation Report Journal Entries.

#### Key Functionality:

The script scrapes and processes the timesheet data from the various network locations without altering the source .xlsx files in any way. Then "snapshots" of the relevant data
are stored to a new local directory for further aggregation.

#### Error Detection:

  - Identifies blank cells and shows the offending rows to the user

  - Detects duplicate entries and shows the offending rows to the user

  - Validates the Cost Codes & descriptions - Once the above two error findings are resolved, the script will take all the acquired timesheet data & cross-reference it against a "Master Code" list set up by management. If any item is found which is not on the approved list; the script will show the offending line items to the user for further review/editing

#### Usage Notes:

The user must manually correct the various errors found by the script in the individual live excel files. The script will need to be run as many times as needed until zero errors are found. Upon each run a new folder will pop-up names "Timesheet_Set". This folder must be deleted after each re-run. When there are no errors to be reported -- the final "Timesheet_Set" must be retained as the "@Run It_4_FINAL_Allocs.bat" launcher will reference its contents.

The user must also manually update & maintain the 'script_paths.xlsx' file inside the "PATHS_Template" folder. 

The user will also need to drag/paste the Payroll Allocation Report excel file into the "READ_Payroll_File" **BEFORE** running any scripts. Be sure to check that only **1** file is in the "READ_Payroll_File" directory at runtime.
 
