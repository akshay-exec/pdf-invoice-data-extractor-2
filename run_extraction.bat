@echo off

REM ------------------ CONFIGURATION ------------------
set PDF_FOLDER=PATH_TO_PDF_FOLDER
set OUTPUT_FOLDER=PATH_TO_OUTPUT_FOLDER

REM ------------------ USER INPUT ------------------
set /p LIMIT="Type and Enter how many PDFs to process (or press ENTER to process all PDFs): "

REM ------------------ RUN SCRIPT ------------------
if "%LIMIT%"=="" (
    python extract2.py "%PDF_FOLDER%" "%OUTPUT_FOLDER%"
) else (
    python extract2.py "%PDF_FOLDER%" "%OUTPUT_FOLDER%" %LIMIT%
)

pause
