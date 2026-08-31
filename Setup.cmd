@echo off
rem ============================================================================
rem  The only file you need to run. It opens a small window that installs,
rem  updates or removes Excel Query Trigger Manager for you.
rem
rem  Nothing here needs administrator rights: everything is written under your
rem  own user profile.
rem ============================================================================
setlocal
if not defined __COMPAT_LAYER set "__COMPAT_LAYER=GdiScaling"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0tools\Setup.ps1" %*
endlocal
