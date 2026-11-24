@echo off
cd /d "%~dp0"

where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw -3 "zoom_auto_sender.pyw" --detached-gui
    exit
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "zoom_auto_sender.pyw" --detached-gui
    exit
)

for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%D\pythonw.exe" (
        start "" "%%D\pythonw.exe" "zoom_auto_sender.pyw" --detached-gui
        exit
    )
)

where py >nul 2>nul
if %errorlevel%==0 (
    start "" py -3 "zoom_auto_sender.py"
    exit
)

where python >nul 2>nul
if %errorlevel%==0 (
    start "" python "zoom_auto_sender.py"
    exit
)

exit
