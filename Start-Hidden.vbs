' ==============================================================================
'  Start-Hidden.vbs
'  Launches ExcelQueryTrigger.ps1 with no console window. This is the
'  recommended launcher for normal day-to-day use.
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
command = """" & powershellExe & """ -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & target & """"

' Setup uses this one-launch override after a first installation so the new
' Dashboard is visible immediately. Normal shortcuts and sign-in stay silent.
If WScript.Arguments.Count > 0 Then
    If LCase(WScript.Arguments(0)) = "-showwindow" Then command = command & " -ShowWindow"
End If

' 0 = hidden window, False = do not wait for it to finish
shell.Run command, 0, False
