@echo off
Net session >nul 2>&1 || (PowerShell start -verb runas '%~0' &exit /b)

:: Use CiTool to list policies

start C:\Windows\System32\CiTool.exe --list-policies
