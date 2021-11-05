@Echo Off
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do     rem"') do (
  set "DEL=%%a"
)

CD %~dp0
If not exist .\Check-Distro.ps1 (
Title Error - Powershell Functions missing
Color 0C
Echo.
call :colorEcho C0 " ERROR________________________________________________________________________________"
Echo.
Echo.
Echo     Check-Distro.PS1 could not located in the current working directory
Echo     Please verify the "Load Check-Distro.bat" AND "Check-Distro.ps1" are 
Echo     in the same folder before running the application again.
Echo.
Echo Press any key to exit
Pause > Nul
exit
)

NET SESSION 1>NUL 2>NUL
IF %ERRORLEVEL% EQU 0 GOTO elevatecheck
CD %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1); close();"
EXIT

:elevatecheck
@echo off
color 0c
cls
echo Checking for Administrator elevation...
timeout /t 1 /nobreak > nul
echo.
echo.
openfiles /local > NUL 2>&1
if %errorlevel%==0 (
	call :colorEcho A0 " Elevation found! Proceeding..."	
	Echo.	
	goto vercheck
) else (
	call :colorEcho C0 " ERROR________________________________________________________________________________"	
	Echo.	
	echo You are not running as an Administrator...
	echo This tool may not function properly without elevation!
	echo.
	echo If you exprerience issues running this Application, close it, and Right-click
	Echo the batch file, then select ^'Run as Administrator^' to try again...
	echo.
	echo Press any key to Continue...
	pause > NUL
	goto vercheck
)

:vercheck
@echo off
CD %~dp0
CLS
@echo off
Color 0a
powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy bypass -Force}"
START /MAX Powershell -NoExit -file ".\Check-Distro.ps1"
exit

:colorEcho
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1i

