@echo off
setlocal
cd /d "%~dp0"
title Statement Finder

REM ---- 1. find Python (ignoring the blank Windows Store shortcut) ----------------------
set "PY="
where python >nul 2>&1 && (python -c "import sys" >nul 2>&1 && set "PY=python")
if not defined PY where py >nul 2>&1 && (py -c "import sys" >nul 2>&1 && set "PY=py")
if not defined PY goto :nopython

REM ---- 2. the window library that comes with Python ------------------------------------
%PY% -c "import tkinter" >nul 2>&1
if errorlevel 1 goto :notk

REM ---- 3. install what is missing (first run only; needs internet) ---------------------
%PY% -c "import pymupdf, docx" >nul 2>&1
if not errorlevel 1 goto :start
echo Installing the packages Statement Finder needs. This is done once, and needs internet.
echo.
%PY% -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :pipfailed
%PY% -c "import pymupdf, docx" >nul 2>&1
if errorlevel 1 goto :pipfailed
echo.
echo Done.

REM ---- 4. start the window (no black console window) ------------------------------------
:start
if defined STATEMENT_FINDER_NO_LAUNCH (
    echo Everything is ready. The window is not started because STATEMENT_FINDER_NO_LAUNCH is set.
    exit /b 0
)
if "%PY%"=="python" where pythonw >nul 2>&1 && (start "" pythonw statement_finder.py & exit /b 0)
if "%PY%"=="py" where pyw >nul 2>&1 && (start "" pyw statement_finder.py & exit /b 0)
where pythonw >nul 2>&1 && (start "" pythonw statement_finder.py & exit /b 0)
where pyw >nul 2>&1 && (start "" pyw statement_finder.py & exit /b 0)
REM no pythonw found: run with python in this window, so an error can be read
%PY% statement_finder.py
if errorlevel 1 pause
exit /b 0

:nopython
echo Python was not found on this PC.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run this file again.
echo.
pause
exit /b 1

:notk
echo This Python has no tkinter, which the window needs.
echo Run the Python installer again, choose Modify, and make sure "tcl/tk and IDLE" is ticked.
echo.
pause
exit /b 1

:pipfailed
echo.
echo The packages could not be installed. Check that this PC is online, then run this file again.
echo Or install them by hand:   %PY% -m pip install "pymupdf>=1.24,<2" python-docx
echo.
pause
exit /b 1
