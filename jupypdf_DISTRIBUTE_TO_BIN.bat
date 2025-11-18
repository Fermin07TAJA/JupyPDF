@echo off
setlocal enabledelayedexpansion

REM jupypdf.bat
REM Usage:
REM   jupypdf -i filename.ipynb
REM   jupypdf -a

if "%1"=="-i" goto withfile
if "%1"=="-a" goto auto

echo Usage:
echo   jupypdf -i filename.ipynb
echo   jupypdf -a
exit /b 1


:withfile
set "NB=%2"
if "%NB%"=="" (
    echo No filename given.
    exit /b 1
)
if not exist "%NB%" (
    echo File not found: %NB%
    exit /b 1
)
call :run "%NB%"
exit /b 0


:auto
set "NB="
set /a count=0
for %%F in (*.ipynb) do (
    set "NB=%%F"
    set /a count+=1
)

if %count%==0 (
    echo No .ipynb files found in "%CD%".
    exit /b 1
)

if %count% GTR 1 (
    echo Multiple .ipynb files found. Use -i to pick one:
    for %%F in (*.ipynb) do echo   %%F
    exit /b 1
)

call :run "%NB%"
exit /b 0


:run
set "NBFILE=%~1"
set "NBPATH=%CD%\%NBFILE%"
set "PDFPATH=%CD%\%~n1.pdf"

echo.
echo ========================================
echo Jupyter to PDF Conversion
echo Notebook: %NBFILE%
echo Working Directory: %CD%
echo Tool Directory: %~dp0
echo ========================================
echo.

python "U:\repos\JupyPDF\jupypdf.py" "%NBPATH%"

echo.
echo ========================================
echo Conversion finished.
echo PDF in: %CD%
echo.
echo Clickable PDF link:
echo %PDFPATH%
echo ========================================
echo.

goto :eof
