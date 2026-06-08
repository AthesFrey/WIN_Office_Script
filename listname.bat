@echo off
setlocal EnableExtensions DisableDelayedExpansion

set "OUT_FILE=list_names.txt"
set "DIR_TMP=%TEMP%\list_dirs_%RANDOM%%RANDOM%.tmp"
set "FILE_TMP=%TEMP%\list_files_%RANDOM%%RANDOM%.tmp"

rem Collect names first. Do not echo each file name as a command.
dir /b /a:d > "%DIR_TMP%" 2>nul
dir /b /a:-d > "%FILE_TMP%" 2>nul

for /f %%C in ('find /v /c "" ^< "%DIR_TMP%"') do set "DC=%%C"
for /f %%C in ('find /v /c "" ^< "%FILE_TMP%"') do set "FC=%%C"

> "%OUT_FILE%" (
    echo [Folders]
    if "%DC%"=="0" (
        echo ^(none^)
    ) else (
        type "%DIR_TMP%"
    )
    echo.
    echo [Files]
    if "%FC%"=="0" (
        echo ^(none^)
    ) else (
        type "%FILE_TMP%"
    )
    echo.
    echo Total: %DC% folders, %FC% files
)

del /f /q "%DIR_TMP%" "%FILE_TMP%" >nul 2>nul

echo Created: "%OUT_FILE%"
type "%OUT_FILE%"
pause
endlocal
