@echo off
echo Running Pipe Class Validation...
echo.

REM Check if virtual environment exists
if not exist ..\..\.venv\Scripts\activate.bat (
    echo Virtual environment not found. Please create a virtual environment first.
    echo See README.md for instructions.
    exit /b 1
)

REM Activate virtual environment and run the script
call ..\..\.venv\Scripts\activate.bat
python line_list_check_pipe_class_refactored.py
pause
