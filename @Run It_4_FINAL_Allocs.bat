REM #START
@echo off
ECHO Starting Final Allocations Process...
ECHO.

REM #create & map temp drive to R:\
ECHO Mapping network drive...
net use R: \\public.playoff.com\Public /persistent:no
IF ERRORLEVEL 1 (
    ECHO Failed to map drive. Drive may already be in use.
    ECHO Press any key to continue...
    pause >nul
    EXIT /B 1
)
ECHO Drive mapped successfully.

REM #change the working dir to the script's dir
ECHO Changing to script directory...
cd /d R:\Dept-Accounting\z_scriptBeta\2.5_Timesheets\_py_scripts
IF ERRORLEVEL 1 (
    ECHO Failed to change directory. Ensure the path exists.
    ECHO Current directory: %CD%
    ECHO Press any key to continue...
    pause >nul
    net use R: /delete
    EXIT /B 1
)
ECHO Directory changed successfully.
ECHO Current directory: %CD%
ECHO.

REM #run it
ECHO Running Python script...
python final_alloc_main.py 2>&1
SET SCRIPT_EXIT_CODE=%ERRORLEVEL%
IF %SCRIPT_EXIT_CODE% NEQ 0 (
    ECHO.
    ECHO ================================
    ECHO Script execution failed with error code: %SCRIPT_EXIT_CODE%
    ECHO ================================
    ECHO Please check the error messages above.
    ECHO Press any key to continue...
    pause >nul
) ELSE (
    ECHO.
    ECHO Script completed successfully!
)

REM #disconnect & remove temp drive R:\
ECHO Disconnecting network drive...
net use R: /delete
IF ERRORLEVEL 1 (
    ECHO Failed to disconnect drive. Please check manually.
    ECHO Press any key to continue...
    pause >nul
)

REM #END
ECHO.
ECHO Process completed. Press any key to close...
pause >nul