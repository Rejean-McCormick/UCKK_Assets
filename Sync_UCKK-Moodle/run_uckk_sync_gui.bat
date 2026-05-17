@echo off
setlocal
set SCRIPT_DIR=%~dp0

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    py "%SCRIPT_DIR%uckk_sync_gui.py"
    goto :eof
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    python "%SCRIPT_DIR%uckk_sync_gui.py"
    goto :eof
)

echo Python n'a pas ete trouve dans le PATH.
pause
