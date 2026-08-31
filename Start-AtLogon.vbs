' ==============================================================================
'  Start-AtLogon.vbs
'  This is the file the "Start automatically when I log in to Windows" registry
'  entry points at. It is identical to Start-Hidden.vbs except that it passes
'  -StartedFromLogon, which is how the application knows it may offer the logon
'  refresh prompt. Starting the application by hand never shows that prompt.
' ==============================================================================
Option Explicit

Dim shell, fso, scriptDir, target, command, env, powershellExe
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
target    = fso.BuildPath(scriptDir, "ExcelQueryTrigger.ps1")

If Not fso.FileExists(target) Then
    MsgBox "ExcelQueryTrigger.ps1 was not found next to this launcher." & vbCrLf & target, 16, "Excel Query Trigger"
    WScript.Quit 1
End If

' ------------------------------------------------------------------------------
'  Sharper text on a scaled display.
'
'  This application is deliberately DPI-unaware: its windows use fixed 96-DPI
'  coordinates, and letting Windows scale the whole window is what keeps the
'  layout correct at 150-200% (including Retina virtual machines). The cost is
'  that Windows stretches the bitmap, which makes text look soft.
'
'  GdiScaling - the "System (Enhanced)" compatibility setting - tells Windows to
'  render GDI text and lines at the real resolution instead of stretching them.
'  The layout is unchanged; only the drawing is sharper. It is set on this
'  process so the PowerShell process started below inherits it.
'
'  To turn it off, comment out the two lines below with a leading apostrophe.
'  An existing value set by you or by IT is never overwritten.
' ------------------------------------------------------------------------------
Set env = shell.Environment("PROCESS")
If Len(env("__COMPAT_LAYER")) = 0 Then env("__COMPAT_LAYER") = "GdiScaling"

powershellExe = shell.ExpandEnvironmentStrings("%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe")
command = """" & powershellExe & """ -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & target & """ -StartedFromLogon"

' 0 = hidden window, False = do not wait for it to finish
shell.Run command, 0, False
