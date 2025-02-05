@echo off
cd /d "%~dp0"
echo Running script from: %CD%\jupypdf.py
pythonw -B "%CD%\jupypdf.py"
exit