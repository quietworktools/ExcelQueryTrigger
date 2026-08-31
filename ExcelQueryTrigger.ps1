# ==============================================================================
#  ExcelQueryTrigger.ps1 - application entry point
#
#  Two runspaces, one shared state object:
#
#      +-------------------+  commands   +--------------------------+
#      |  UI runspace      | ----------> |  Engine runspace (STA)   |
#      |  WinForms + tray  |             |  watchers, queue, Excel  |
#      |  250 ms timer     | <---------- |  status, activity, jobs  |
#      +-------------------+   state     +--------------------------+
#
#  The UI never blocks on Excel and the engine never touches a control.
#
#  Launch:  powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ExcelQueryTrigger.ps1
#  or double-click Start-Hidden.vbs (no console window).
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$ShowWindow,        # overrides "start minimized" for this launch only
    [switch]$StartedFromLogon,  # set only by Start-AtLogon.vbs; enables the logon prompt
    [switch]$NoHighDpi          # accepted and ignored; kept so an old shortcut still starts
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$appRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($appRoot)) { $appRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

# Common and the splash are intentionally tiny and load first. The rest of the
# application can then report honest progress while its modules are loaded.
. (Join-Path (Join-Path $appRoot 'src') 'Common.ps1')
. (Join-Path (Join-Path $appRoot 'src') 'UISplash.ps1')

# PowerShell 5.1 WinForms uses many fixed 96-DPI coordinates in this utility.
# Declaring legacy System-DPI awareness made fonts render at native Retina/high-DPI
# size while several runtime-created controls kept 96-DPI geometry, which could
# destroy the layout at 150-200% scaling (especially under Parallels). Leave the
# process DPI-unaware and let Windows scale the complete window consistently.
#
# Sharpness is handled outside the process instead: the launchers set
# __COMPAT_LAYER=GdiScaling, so Windows draws GDI text at the real resolution
# without changing any coordinate this code uses. Nothing here depends on it -
# starting the script directly simply gives the softer stretched rendering.
# Set-FormWithinWorkingArea remains the secondary guard for small work areas.
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
# This must run before the splash or any other WinForms control is created.
# Calling it later throws once a control handle exists on the UI thread.
[System.Windows.Forms.Application]::SetUnhandledExceptionMode(
    [System.Windows.Forms.UnhandledExceptionMode]::CatchException)

# ---- only one instance may own the watchers ---------------------------------
# A second launch does not just fail with an "already running" message. It
# asks the existing UI instance to restore its Dashboard from the tray. This
# also makes it obvious when a hidden PowerShell host is genuinely stuck inside
# an Excel COM call: a healthy UI acknowledges the request within 1.5 seconds.
$instanceMutexName       = 'Local\ExcelQueryTriggerSingleInstance'
$activateEventName       = 'Local\ExcelQueryTriggerActivateExisting'
$activateAckEventName    = 'Local\ExcelQueryTriggerActivateAcknowledged'
$instanceMutex = New-Object System.Threading.Mutex($false, $instanceMutexName)
$ownsInstanceMutex = $false
try {
    $ownsInstanceMutex = $instanceMutex.WaitOne(0, $false)
}
catch [System.Threading.AbandonedMutexException] {
    # The previous process died without releasing the mutex. .NET transfers
    # ownership to this thread when it raises AbandonedMutexException, so it is
    # safe to continue as the new primary instance.
    $ownsInstanceMutex = $true
}

if (-not $ownsInstanceMutex) {
    $activateEvent = $null
    $activateAckEvent = $null
    try {
        $activateEvent = New-Object System.Threading.EventWaitHandle(
            $false, [System.Threading.EventResetMode]::AutoReset, $activateEventName)
        $activateAckEvent = New-Object System.Threading.EventWaitHandle(
            $false, [System.Threading.EventResetMode]::AutoReset, $activateAckEventName)
        [void]$activateAckEvent.Reset()
        [void]$activateEvent.Set()

        if ($activateAckEvent.WaitOne(1500, $false)) { return }

        [System.Windows.Forms.MessageBox]::Show(
            ('Excel Query Trigger Manager is still running in the background, but it did not respond to a request to reopen the Dashboard.' + [Environment]::NewLine + [Environment]::NewLine +
             'It may still be starting or waiting inside an Excel operation. Check the system tray and Task Manager. Do not force-close Excel while a workbook is being saved.'),
            'Excel Query Trigger', 'OK', 'Warning') | Out-Null
    }
    finally {
        try { if ($null -ne $activateEvent) { $activateEvent.Dispose() } } catch { }
        try { if ($null -ne $activateAckEvent) { $activateAckEvent.Dispose() } } catch { }
        try { $instanceMutex.Dispose() } catch { }
    }
    return
}

$splash = $null
if (-not $StartedFromLogon) {
    $splash = Show-StartupSplash -AppRoot $appRoot
    Update-StartupSplash -Splash $splash -Message 'Loading application components...' -Percent 8
}

try {
    foreach ($file in @('LogManager.ps1', 'ConfigManager.ps1', 'StartupManager.ps1',
            'UIDialogs.ps1', 'UIDiagnostics.ps1', 'UIWorkbookInfo.ps1', 'UIQueryScope.ps1',
            'UIUpdate.ps1', 'UIManager.ps1', 'UIRuleEditor.ps1', 'UISettings.ps1',
            'UIAddFileWizard.ps1', 'UILogonPrompt.ps1')) {
        . (Join-Path (Join-Path $appRoot 'src') $file)
    }
    Update-StartupSplash -Splash $splash -Message 'Application components loaded.' -Percent 16 `
        -ActivityText ('[OK] Application components loaded' + [Environment]::NewLine + '[..] Reading settings and trigger rules')
}
catch {
    Close-StartupSplash -Splash $splash
    [System.Windows.Forms.MessageBox]::Show(
        ('Excel Query Trigger could not load its application files:' + [Environment]::NewLine + $_.Exception.Message),
        'Excel Query Trigger', 'OK', 'Error') | Out-Null
    try { $instanceMutex.ReleaseMutex() } catch { }
    try { $instanceMutex.Dispose() } catch { }
    return
}

Update-StartupSplash -Splash $splash -Message 'Preparing application folders...' -Percent 20
$paths = Get-AppPaths -AppRoot $appRoot

$runspace       = $null
$enginePipeline = $null
$shared         = $null

try {
    Update-StartupSplash -Splash $splash -Message 'Reading settings and trigger rules...' -Percent 30
    Initialize-LogManager -LogDirectory $paths.LogDir -HistoryPath $paths.HistoryPath -Shared $null `
        -RetentionDays 30 -MaxTotalMegabytes 50 -DebugLogging $false -SourceTag 'ui'

    # ---- first run: seed the configuration -----------------------------------
    if (-not (Test-Path -LiteralPath $paths.ConfigPath)) {
        $template = Join-Path (Join-Path $appRoot 'config') 'rules.json'
        if ((Test-Path -LiteralPath $template) -and ($template -ne $paths.ConfigPath)) {
            Copy-Item -LiteralPath $template -Destination $paths.ConfigPath -Force
        }
        else {
            Save-AppConfiguration -Path $paths.ConfigPath -Config @{ schemaVersion = (Get-ConfigSchemaVersion); appSettings = (Get-DefaultAppSettings); rules = @() }
        }
        Write-AppLog -Level 'INFO' -Message ('Created a new configuration file at {0}' -f $paths.ConfigPath)
    }

    $config = Import-AppConfiguration -Path $paths.ConfigPath
    if (-not [string]::IsNullOrWhiteSpace($config.loadError)) {
        [System.Windows.Forms.MessageBox]::Show(
            ('rules.json could not be read and defaults are being used:' + [Environment]::NewLine + $config.loadError),
            'Excel Query Trigger', 'OK', 'Warning') | Out-Null
    }
    if ($config.migrationRequired -and [string]::IsNullOrWhiteSpace($config.loadError)) {
        # Persist the upgraded schema before the engine starts, so retired
        # settings disappear from disk on this launch.
        Save-AppConfiguration -Path $paths.ConfigPath -Config $config
        $config.schemaVersion = Get-ConfigSchemaVersion
        $config.migrationRequired = $false
        Write-AppLog -Level 'INFO' -Message ('Updated the settings file to configuration schema {0}.' -f (Get-ConfigSchemaVersion))
    }
    if ($ShowWindow) { $config.appSettings.startMinimized = $false }
    Update-StartupSplash -Splash $splash -Message ('Loaded {0} trigger rule(s).' -f @($config.rules).Count) `
        -Detail 'Next: start monitoring and read workbook information.' -Percent 42 `
        -ActivityText (('[OK] Application components loaded' + [Environment]::NewLine +
            '[OK] Settings loaded' + [Environment]::NewLine +
            '[OK] {0} trigger rule(s) loaded' + [Environment]::NewLine +
            '[..] Starting the monitoring engine') -f @($config.rules).Count)

    # ---- shared state --------------------------------------------------------
    $shared = New-SharedState
    $shared.DebugLogging = (ConvertTo-BoolValue $config.appSettings.debugLogging $false)
    Update-StartupSplash -Splash $splash -Message 'Starting the monitoring engine...' -Percent 52

    # The UI logger writes to the same files; give it the shared activity list
    # so messages raised on the UI side also appear in the dashboard.
    Initialize-LogManager -LogDirectory $paths.LogDir -HistoryPath $paths.HistoryPath -Shared $shared `
        -RetentionDays (ConvertTo-IntValue $config.appSettings.logRetentionDays 30 1) `
        -MaxTotalMegabytes (ConvertTo-IntValue $config.appSettings.logMaxTotalMegabytes 50 1) `
        -DebugLogging ([bool]$shared.DebugLogging) -SourceTag 'ui'

    # An installation registered before Start-AtLogon.vbs existed would keep
    # starting without -StartedFromLogon, and the logon prompt would never
    # appear. Repoint it silently at the current launcher.
    if (Update-StartupRegistration -Paths $paths) {
        Write-AppLog -Level 'INFO' -Message 'Updated the Windows startup entry to the current launcher.'
    }

    # ---- start the engine ----------------------------------------------------
    $engineScript = {
        Set-StrictMode -Version 1.0
        $ErrorActionPreference = 'Stop'
        try {
            foreach ($engineFile in @('Common.ps1', 'LogManager.ps1', 'ConfigManager.ps1', 'TriggerManager.ps1',
                    'ExcelManager.ps1', 'JobManager.ps1', 'Engine.ps1')) {
                . (Join-Path (Join-Path $AppRoot 'src') $engineFile)
            }
            Start-Engine -Shared $Shared -Paths $Paths
        }
        catch {
            $Shared.FatalError = '{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message
            $Shared.Status     = 'Error'
            $Shared.EngineAlive = $false
        }
    }

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = [System.Threading.ApartmentState]::STA   # required for Excel COM
    $runspace.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('Shared', $shared)
    $runspace.SessionStateProxy.SetVariable('Paths', $paths)
    $runspace.SessionStateProxy.SetVariable('AppRoot', $appRoot)

    $enginePipeline = [powershell]::Create()
    $enginePipeline.Runspace = $runspace
    [void]$enginePipeline.AddScript($engineScript)
    $engineHandle = $enginePipeline.BeginInvoke()
    $engineHandle | Out-Null
    Update-StartupSplash -Splash $splash -Message 'Building the dashboard...' -Percent 68

    # ---- last-resort exception handling for the UI thread --------------------
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $threadEventArgs)
        try { Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' -Message ('UI exception: {0}' -f $threadEventArgs.Exception.Message) } catch { }
    })

    Write-AppLog -Level 'INFO' -Message ('--- Excel Query Trigger Manager v{0} started ({1}) ---' -f (Get-AppVersion), `
        $(if ($StartedFromLogon) { 'from Windows logon' } else { 'started by hand' }))

    # ---- run the dashboard (blocks until the user exits) ---------------------
    Show-Dashboard -Shared $shared -Paths $paths -Config $config -StartedFromLogon:$StartedFromLogon -Splash $splash
    $splash = $null
}
catch {
    Close-StartupSplash -Splash $splash
    $splash = $null
    $message = '{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message
    try { Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' -Message ('Startup failed: {0}' -f $message) } catch { }
    [System.Windows.Forms.MessageBox]::Show(
        ('Excel Query Trigger could not start:' + [Environment]::NewLine + $message),
        'Excel Query Trigger', 'OK', 'Error') | Out-Null
}
finally {
    Close-StartupSplash -Splash $splash
    try { Stop-WorkbookInfoBackgroundScan } catch { }
    # ---- ordered shutdown ----------------------------------------------------
    if ($null -ne $shared) {
        # Normal Exit is already blocked while a workbook job is active. This
        # extra guard covers an unexpected UI failure: keep the host alive until
        # the engine finishes the workbook rather than cancelling or killing it.
        if ($shared.CurrentJob.Active -and $shared.EngineAlive) {
            try { Write-AppLog -Level 'ERROR' -Message 'The UI ended unexpectedly while a workbook job was active. Waiting for the engine to finish safely before shutdown.' } catch { }
            while ($shared.CurrentJob.Active -and $shared.EngineAlive) { Start-Sleep -Milliseconds 500 }
        }

        $shared.ShouldExit = $true
        $deadline = (Get-Date).AddSeconds(20)
        while ($shared.EngineAlive -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }

        # Do not force-kill Excel here. ExcelManager owns cleanup and only
        # terminates its dedicated process after workbook closure is confirmed.
        # If engine shutdown itself is abnormal, preserving the workbook is more
        # important than guaranteeing that every EXCEL.EXE disappears.
        if ([int]$shared.OwnedExcelPid -gt 0) {
            try { Write-AppLog -Level 'WARN' -Message ('Application shutdown left dedicated Excel pid {0} for safety; no forced termination was attempted.' -f [int]$shared.OwnedExcelPid) } catch { }
        }
    }

    # If the UI died before Stop-Application ran, the tray icon would otherwise
    # stay behind as a ghost until the user hovers over it.
    try {
        if ($script:UiControls.ContainsKey('Tray') -and $null -ne $script:UiControls.Tray) {
            $script:UiControls.Tray.Visible = $false
            $script:UiControls.Tray.Dispose()
        }
    }
    catch { }

    if ($null -ne $enginePipeline) {
        try { $enginePipeline.Stop() } catch { }
        try { $enginePipeline.Dispose() } catch { }
    }
    if ($null -ne $runspace) {
        try { $runspace.Close() } catch { }
        try { $runspace.Dispose() } catch { }
    }

    try { Write-AppLog -Level 'INFO' -Message ('--- Excel Query Trigger Manager v{0} exited ---' -f (Get-AppVersion)) } catch { }

    try { $instanceMutex.ReleaseMutex() } catch { }
    try { $instanceMutex.Dispose() } catch { }
}
