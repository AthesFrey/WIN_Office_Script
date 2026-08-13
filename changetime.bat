@echo off
setlocal EnableExtensions DisableDelayedExpansion

:: Fixed root: script folder
set "ROOT_DIR=%~dp0"
pushd "%ROOT_DIR%" >nul 2>&1 || (
  echo [ERR] Cannot switch to script folder.
  echo %ROOT_DIR%
  goto END
)

echo ===============================
echo Bulk change file/folder timestamps
echo Date format: YYYY-MM-DD / YYYYMMDD / YYYY/MM/DD
echo Time format: 6 digits = HHMMSS, 4 digits = HHMM (seconds=00), blank = current time
echo Digits only. Do not use colon.
echo Note: access denied / locked / special objects will be skipped
echo Working folder:
echo %ROOT_DIR%
echo ===============================

:: Get current date/time
for /f %%I in ('powershell -NoLogo -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set "CUR_DATE=%%I"
for /f %%I in ('powershell -NoLogo -NoProfile -Command "Get-Date -Format HHmmss"') do set "CUR_TIME=%%I"

:: Read date
:READ_DATE
set "DATE_RAW="
set /p "DATE_RAW=Enter date [default %CUR_DATE%]: "
if "%DATE_RAW%"=="" set "DATE_RAW=%CUR_DATE%"

for /f "tokens=* delims= " %%A in ("%DATE_RAW%") do set "DATE_RAW=%%A"
set "DATE_STD=%DATE_RAW:/=-%"
set "DATE_STD=%DATE_STD:\=-%"

call :AllDigits "%DATE_STD%" _ALL
if "%_ALL%"=="1" if "%DATE_STD:~8,1%"=="" (
  set "DATE_STD=%DATE_STD:~0,4%-%DATE_STD:~4,2%-%DATE_STD:~6,2%"
)

if not "%DATE_STD:~4,1%"=="-" goto BAD_DATE
if not "%DATE_STD:~7,1%"=="-" goto BAD_DATE

set "Y=%DATE_STD:~0,4%"
set "M=%DATE_STD:~5,2%"
set "D=%DATE_STD:~8,2%"

call :AllDigits "%Y%" _Y
call :AllDigits "%M%" _M
call :AllDigits "%D%" _D
if "%_Y%%_M%%_D%" neq "111" goto BAD_DATE

if "%M:~0,1%"=="0" (set /a MV=%M:~1,1%) else set /a MV=%M%
if "%D:~0,1%"=="0" (set /a DV=%D:~1,1%) else set /a DV=%D%

if %MV% LSS 1  goto BAD_DATE
if %MV% GTR 12 goto BAD_DATE
if %DV% LSS 1  goto BAD_DATE
if %DV% GTR 31 goto BAD_DATE

echo [OK] Date = %DATE_STD%
goto DATE_OK

:BAD_DATE
echo [ERR] Invalid date
goto READ_DATE

:DATE_OK

:: Read time
:READ_TIME
set "TIME_RAW="
set /p "TIME_RAW=Enter time (4 or 6 digits, default current time %CUR_TIME%): "
if "%TIME_RAW%"=="" set "TIME_RAW=%CUR_TIME%"

call :AllDigits "%TIME_RAW%" _TD
if "%_TD%"=="0" (
  echo [ERR] Digits only
  goto READ_TIME
)

call :StrLen "%TIME_RAW%" _TL
if "%_TL%"=="6" (
  set "HH=%TIME_RAW:~0,2%"
  set "MI=%TIME_RAW:~2,2%"
  set "SS=%TIME_RAW:~4,2%"
) else if "%_TL%"=="4" (
  set "HH=%TIME_RAW:~0,2%"
  set "MI=%TIME_RAW:~2,2%"
  set "SS=00"
) else (
  echo [ERR] Length must be 4 or 6
  goto READ_TIME
)

if 1%HH% GEQ 124 (echo [ERR] Hour must be 00-23 & goto READ_TIME)
if 1%MI% GEQ 160 (echo [ERR] Minute must be 00-59 & goto READ_TIME)
if 1%SS% GEQ 160 (echo [ERR] Second must be 00-59 & goto READ_TIME)

set "TIME_STD=%HH%:%MI%:%SS%"
echo [OK] Time = %TIME_STD%

:: Confirm
set "DT_FULL=%DATE_STD% %TIME_STD%"
echo Target time: [%DT_FULL%]

:READ_CONFIRM
set "CONFIRM="
set /p "CONFIRM=Run now? [Y/n] (Enter = Y): "
if not defined CONFIRM set "CONFIRM=Y"

if /I "%CONFIRM%"=="Y" goto DO_EXEC
if /I "%CONFIRM%"=="N" (
  echo Cancelled
  goto END_AFTER_POPD
)

echo [ERR] Please enter Y or N, or press Enter
goto READ_CONFIRM

:DO_EXEC
echo Processing... (root folder itself will be skipped)

set "ROOT_DIR_PS=%ROOT_DIR%"

powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
 "$ErrorActionPreference='Stop';" ^
 "try {" ^
 "  $raw = '%DT_FULL%';" ^
 "  try { $ts = [datetime]::ParseExact($raw,'yyyy-MM-dd HH:mm:ss',[System.Globalization.CultureInfo]::InvariantCulture) } catch { Write-Host ('Parse failed: ' + $_.Exception.Message); exit 2 }" ^
 "  $root = $env:ROOT_DIR_PS;" ^
 "  $items = New-Object System.Collections.Generic.List[Object];" ^
 "  $fails = New-Object System.Collections.Generic.List[Object];" ^
 "  function Classify([string]$etype,[string]$msg){" ^
 "    if($etype -match 'UnauthorizedAccess'){ return 'Access denied / policy' }" ^
 "    elseif($etype -match 'PathTooLong'){ return 'Path too long' }" ^
 "    elseif($etype -eq 'System.IO.IOException' -and $msg -match 'being used|used by another process|sharing violation|cannot access the file'){ return 'Locked / in use' }" ^
 "    elseif($msg -match 'read-only'){ return 'Read-only' }" ^
 "    elseif($etype -eq 'System.IO.IOException'){ return 'I/O error' }" ^
 "    else { return 'Other' }" ^
 "  }" ^
 "  function AddFail([string]$path,[bool]$ro,[string]$etype,[string]$msg){" ^
 "    $cat = Classify $etype $msg;" ^
 "    $fails.Add([pscustomobject]@{Path=$path; ReadOnlyBefore=$ro; ExceptionType=$etype; Message=$msg; Category=$cat}) | Out-Null" ^
 "  }" ^
 "  function EnumSafe([string]$dir){" ^
 "    $entries = @();" ^
 "    try { $entries = Get-ChildItem -LiteralPath $dir -Force -ErrorAction Stop }" ^
 "    catch { AddFail $dir $false $_.Exception.GetType().FullName $_.Exception.Message; return }" ^
 "    foreach($e in $entries){" ^
 "      if($e.Attributes -band [IO.FileAttributes]::ReparsePoint){ continue }" ^
 "      $items.Add($e) | Out-Null;" ^
 "      if($e.PSIsContainer){ EnumSafe $e.FullName }" ^
 "    }" ^
 "  }" ^
 "  EnumSafe $root;" ^
 "  $total = $items.Count; $ok = 0;" ^
 "  foreach($f in $items){" ^
 "    $ro = $false;" ^
 "    try { if($f.Attributes -band [IO.FileAttributes]::ReadOnly){ $ro = $true; $f.Attributes = ($f.Attributes -band (-bnot [IO.FileAttributes]::ReadOnly)) } } catch {}" ^
 "    try { $f.CreationTime = $ts; $f.LastAccessTime = $ts; $f.LastWriteTime = $ts; $ok++ }" ^
 "    catch { AddFail $f.FullName $ro $_.Exception.GetType().FullName $_.Exception.Message }" ^
 "  }" ^
 "  $fail = $fails.Count;" ^
 "  Write-Host ('Total:{0}  Success:{1}  Skipped/Failed:{2}  (root skipped)' -f $total,$ok,$fail);" ^
 "  if($fail -gt 0){" ^
 "    $groups = $fails | Group-Object Category | Sort-Object Count -Descending;" ^
 "    Write-Host 'Failure summary:';" ^
 "    foreach($g in $groups){ Write-Host ('  {0} : {1}' -f $g.Name,$g.Count) }" ^
 "    Write-Host 'First 20 failure samples:';" ^
 "    $fails | Select-Object -First 20 Path,Category,ReadOnlyBefore,ExceptionType | Format-Table -AutoSize | Out-String | Write-Host;" ^
 "  }" ^
 "  exit 0" ^
 "} catch {" ^
 "  Write-Host ('Script terminated early: ' + $_.Exception.Message);" ^
 "  exit 10" ^
 "}"

set "PS_RC=%errorlevel%"
echo Done (errorlevel=%PS_RC%)

if "%PS_RC%"=="0" (
  echo Finished
) else if "%PS_RC%"=="2" (
  echo Date/time parse failed
) else if "%PS_RC%"=="10" (
  echo PowerShell main flow terminated early
) else (
  echo An error occurred. Check the output above.
)

:END_AFTER_POPD
popd >nul 2>&1

:END
echo.
echo Press any key to exit...
pause >nul
endlocal
goto :eof

:AllDigits
setlocal
set "S=%~1"
if "%S%"=="" (
  endlocal & set "%~2=0"
  goto :eof
)
for /f "delims=0123456789" %%Q in ("%S%") do (
  endlocal & set "%~2=0"
  goto :eof
)
endlocal & set "%~2=1"
goto :eof

:StrLen
setlocal EnableDelayedExpansion
set "S=%~1"
set /a L=0
:SL
set "C=!S:~%L%,1!"
if "!C!"=="" (
  endlocal & set "%~2=%L%"
  goto :eof
)
set /a L+=1
goto SL
