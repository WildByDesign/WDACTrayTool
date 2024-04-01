@echo off
Net session >nul 2>&1 || (PowerShell start -verb runas '%~0' &exit /b)

:: Ensure that current directory is used
pushd "%~dp0"

:: Move to relevant policy directory
cd ..\policies\EnforcedMode

:: Copy all (*.cip) policy files, overwriting existing policy files of the same filename
copy /y *.cip C:\Windows\System32\CodeIntegrity\CiPolicies\Active\
