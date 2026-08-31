@echo off
rem ============================================================================
rem  Visible launcher - for the first run and for troubleshooting.
rem
rem  The console window stays open on purpose: PowerShell writes here, so if the
rem  application fails to start you can see why. Leave it minimised, or close it
rem  and use Start-Hidden.vbs for normal day-to-day use.
rem
rem  Closing this window also closes the application.
rem ============================================================================
setlocal
set "ROOT=%~dp0..\\"
rem Sharper text on a scaled display. The application is deliberately
rem DPI-unaware so that Windows scales the whole window and the fixed 96-DPI
rem layout stays correct; GdiScaling ("System (Enhanced)") then makes Windows
rem draw the GDI text at the real resolution instead of stretching it. Remove
rem the line below to turn it off. A value already set by you or IT wins.
if not defined __COMPAT_LAYER set "__COMPAT_LAYER=GdiScaling"

echo Starting Excel Query Trigger Manager...
echo.
echo   This console belongs to the application - closing it stops the app.
echo   For normal use start it from the Start menu, or with Start-Hidden.vbs
echo   or turn on "Start automatically when I log in to Windows" in Settings.
echo.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -STA -File "%ROOT%ExcelQueryTrigger.ps1" -ShowWindow %*
endlocal
