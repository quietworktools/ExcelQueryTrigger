# ==============================================================================
#  Engine.ps1  (Engine runspace only)
#
#  One loop owns everything that can block: watcher events, debounce timers,
#  the job queue and the Excel refreshes themselves. The UI never calls into
#  this file - it posts commands and reads status through the shared state.
#
#  Loop order matters:
#     commands -> watcher events -> due triggers -> run one job -> health
#  Commands are read first so Pause and Exit are honoured as early as possible.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:EngineConfig       = $null
$script:EnginePaths        = $null
$script:EngineShared       = $null
$script:LastHealthCheck    = [DateTime]::MinValue
$script:LastStateSave      = [DateTime]::MinValue
$script:LastScheduleCheck  = [DateTime]::MinValue

# How late a scheduled run may still start. It exists so that starting the
# application at 07:31 for an 07:30 rule still runs it, while starting it at
# 09:00 does not silently fire a morning refresh at nine o'clock.
$script:ScheduleCatchUpSeconds = 120

function Get-EngineRules {
    <#
        Always returns an array, even an empty one. PowerShell unrolls a
        collection on return, so a bare "return @()" arrives at the caller as
        $null - which is exactly what stopped the engine when no rules existed
        yet. The comma operator wraps the array so it survives the pipeline.
    #>
    if ($null -eq $script:EngineConfig) { return ,@() }
    $rules = @($script:EngineConfig.rules)
    return ,$rules
}

function Get-EngineAppSettings {
    if ($null -eq $script:EngineConfig) { return (Get-DefaultAppSettings) }
    return $script:EngineConfig.appSettings
}

function Find-RuleById {
    param([string]$RuleId)
    foreach ($rule in (Get-EngineRules)) {
        if ([string]$rule.id -eq $RuleId) { return $rule }
    }
    return $null
}

function Update-AppStatus {
    <#  Single place that decides what the dashboard light shows.  #>
    param([string]$Detail = $null)

    $shared = $script:EngineShared
    if ($null -eq $shared) { return }

    if ($shared.ShouldExit) {
        $shared.Status = 'Stopping'
    }
    elseif ($shared.Paused) {
        $shared.Status = 'Paused'
    }
    elseif ($shared.CurrentJob.Active) {
        $shared.Status = 'Refreshing'
    }
    else {
        $summary = Get-WatcherSummary -Rules (Get-EngineRules)
        if ($summary.Expected -gt 0 -and $summary.Active -lt $summary.Expected) {
            $shared.Status = 'Degraded'
            if ($null -eq $Detail) {
                $inactiveText = @($summary.Inactive) -join ', '
                if ([string]::IsNullOrWhiteSpace($inactiveText)) { $inactiveText = 'unknown rule' }
                $Detail = '{0} of {1} monitor(s) unavailable: {2}' -f `
                    ($summary.Expected - $summary.Active), $summary.Expected, $inactiveText
            }
        }
        else {
            $shared.Status = 'Running'
            if ($null -eq $Detail) { $Detail = '{0} watcher(s) active.' -f $summary.Active }
        }
    }

    if ($null -ne $Detail) { $shared.StatusDetail = $Detail }
}

function Invoke-EngineReload {
    <#  Re-reads rules.json and re-arms the watchers without a restart.  #>
    param([switch]$Silent)

    $script:EngineShared.StartupMessage = 'Reading trigger rules...'
    $script:EngineConfig = Import-AppConfiguration -Path $script:EnginePaths.ConfigPath
    if (-not [string]::IsNullOrWhiteSpace($script:EngineConfig.loadError)) {
        Write-AppLog -Level 'ERROR' -ErrorType 'ConfigurationError' `
            -Message ('rules.json could not be read ({0}). Running with defaults until it is fixed.' -f $script:EngineConfig.loadError)
    }

    Set-LogDebugMode -Enabled (ConvertTo-BoolValue (Get-EngineAppSettings).debugLogging $false)

    $engineRules = Get-EngineRules
    $script:EngineShared.StartupTotal   = $engineRules.Count
    $script:EngineShared.StartupCurrent = 0
    foreach ($rule in $engineRules) {
        $script:EngineShared.StartupMessage = ('Preparing rule: {0}' -f [string]$rule.name)
        Set-RuleState -Shared $script:EngineShared -RuleId ([string]$rule.id) -Values @{} | Out-Null
        $script:EngineShared.StartupCurrent = [int]$script:EngineShared.StartupCurrent + 1
    }

    $watcherRules = @($engineRules | Where-Object {
        (ConvertTo-BoolValue $_.enabled $true) -and (Test-TriggerUsesWatcher $_.trigger.type)
    })
    $script:EngineShared.StartupMessage = 'Activating file and folder monitors...'
    $script:EngineShared.StartupTotal   = $watcherRules.Count
    $script:EngineShared.StartupCurrent = 0
    Sync-RuleWatchers -Rules $engineRules -Paused ([bool]$script:EngineShared.Paused)
    $script:EngineShared.ConfigVersion = [int]$script:EngineShared.ConfigVersion + 1

    if (-not $Silent) {
        Write-AppLog -Level 'INFO' -Message ('Configuration reloaded: {0} rule(s).' -f (Get-EngineRules).Count)
    }

    # Do not claim startup monitoring is ready merely because Sync returned.
    # Network shares can be unavailable for a while after sign-in, and the
    # Dashboard must not be published until the watcher objects themselves are
    # present and raising events.
    $monitorSummary = Get-WatcherSummary -Rules $engineRules
    $script:EngineShared.StartupTotal   = [int]$monitorSummary.Expected
    $script:EngineShared.StartupCurrent = [int]$monitorSummary.Active
    if ([int]$monitorSummary.Active -ge [int]$monitorSummary.Expected) {
        $script:EngineShared.StartupMessage = 'File and folder monitors are active.'
    }
    else {
        $script:EngineShared.StartupMessage = 'Waiting for file and folder monitors...'
    }
    Update-AppStatus
}

function Invoke-EngineCommands {
    <#  Drains the UI command queue. Returns $true if a reload happened.  #>
    $shared = $script:EngineShared
    $reloaded = $false

    while ($shared.Commands.Count -gt 0) {
        $command = $null
        try { $command = $shared.Commands.Dequeue() } catch { break }
        if ($null -eq $command) { continue }

        try {
            switch ([string]$command.Type) {
                'Pause' {
                    if (-not $shared.Paused) {
                        $shared.Paused = $true
                        Stop-AllRuleWatchers
                        Sync-RuleWatchers -Rules (Get-EngineRules) -Paused $true
                        Write-AppLog -Level 'WARN' -Message 'Monitoring paused. File triggers are disabled; Run Now still works.'
                    }
                }
                'Resume' {
                    if ($shared.Paused) {
                        $shared.Paused = $false
                        Clear-TriggerState
                        Sync-RuleWatchers -Rules (Get-EngineRules) -Paused $false
                        Write-AppLog -Level 'INFO' -Message 'Monitoring resumed.'
                    }
                }
                'Reload' {
                    Invoke-EngineReload
                    $reloaded = $true
                }
                'RunNow' {
                    $rule = Find-RuleById -RuleId ([string]$command.RuleId)
                    if ($null -eq $rule) {
                        Write-AppLog -Level 'WARN' -Message ('Run Now requested for an unknown rule id: {0}' -f $command.RuleId)
                    }
                    else {
                        # The logon prompt reuses this command, so the history
                        # shows where a job actually came from.
                        $source = [string]$command.Source
                        if ([string]::IsNullOrWhiteSpace($source)) { $source = 'Manual' }
                        $reason = $(if ($source -eq 'Logon') { 'Logon run requested.' } else { 'Manual run requested.' })
                        Write-AppLog -Level 'INFO' -RuleName $rule.name -Message $reason
                        # Automatic logon rules use the recent-refresh guard.
                        # Rules explicitly selected in the logon dialog have
                        # already been approved and bypass this second prompt.
                        if ($source -eq 'Logon' -and (ConvertTo-BoolValue $command.CheckRecentRefresh $false)) {
                            Queue-TriggeredRuleRun -Rule $rule -Source $source
                        }
                        else {
                            $job = New-RefreshJob -Rule $rule -TriggerSource $source
                            Add-RefreshJob -Job $job -CoalesceDuplicateWorkbooks (ConvertTo-BoolValue (Get-EngineAppSettings).coalesceDuplicateWorkbooks $true) | Out-Null
                        }
                    }
                }
                'RunWorkbook' {
                    # A one-off refresh of a workbook that has no rule. It is
                    # turned into an ordinary job so queueing, de-duplication,
                    # cancellation, logging and history all behave identically.
                    $path = [string]$command.Path
                    if ([string]::IsNullOrWhiteSpace($path)) {
                        Write-AppLog -Level 'WARN' -Message 'Manual workbook refresh requested without a path.'
                    }
                    else {
                        $adhocAction = $null
                        if ($command.ContainsKey('Action') -and $null -ne $command.Action) { $adhocAction = $command.Action }
                        $adhoc = New-ManualWorkbookRule -Path $path -Action $adhocAction `
                            -DefaultTimeoutSeconds (ConvertTo-IntValue (Get-EngineAppSettings).defaultRefreshWarningSeconds 300 5) `
                            -AllowWorkbookMacros (ConvertTo-BoolValue $command.AllowWorkbookMacros (ConvertTo-BoolValue (Get-EngineAppSettings).allowWorkbookMacrosByDefault $false))
                        Write-AppLog -Level 'INFO' -RuleName $adhoc.name -Message 'Manual workbook refresh requested.'
                        $job = New-RefreshJob -Rule $adhoc -TriggerSource 'Manual'
                        Add-RefreshJob -Job $job -CoalesceDuplicateWorkbooks (ConvertTo-BoolValue (Get-EngineAppSettings).coalesceDuplicateWorkbooks $true) | Out-Null
                    }
                }
                'RunAllManual' {
                    foreach ($rule in (Get-EngineRules)) {
                        if ((ConvertTo-BoolValue $rule.enabled $true) -and [string]$rule.trigger.type -eq 'Manual') {
                            $job = New-RefreshJob -Rule $rule -TriggerSource 'Manual'
                            Add-RefreshJob -Job $job -CoalesceDuplicateWorkbooks $true | Out-Null
                        }
                    }
                }
                'ApproveTriggeredRun' {
                    $ruleId = [string]$command.RuleId
                    if ($shared.PendingApprovalRuleIds.ContainsKey($ruleId)) { $shared.PendingApprovalRuleIds.Remove($ruleId) }
                    $rule = Find-RuleById -RuleId $ruleId
                    if ($null -eq $rule -or -not (ConvertTo-BoolValue $rule.enabled $true)) {
                        Write-AppLog -Level 'WARN' -Message ('Approved trigger no longer has an enabled rule: {0}' -f $ruleId)
                    }
                    else {
                        $source = [string]$command.Source
                        if ([string]::IsNullOrWhiteSpace($source)) { $source = 'File' }
                        $triggerFile = [string]$command.TriggerFile
                        $triggerFiles = @($command.TriggerFiles)
                        $automaticallyApproved = ConvertTo-BoolValue $command.AutomaticApproval $false
                        $approvalMessage = $(if ($automaticallyApproved) {
                            'No recent workbook refresh was found; the triggered refresh was approved automatically.'
                        } else {
                            'User approved the triggered refresh.'
                        })
                        Write-AppLog -Level 'INFO' -RuleName $rule.name -Message $approvalMessage
                        $job = New-RefreshJob -Rule $rule -TriggerSource $source -TriggerFile $triggerFile -TriggerFiles $triggerFiles
                        Add-RefreshJob -Job $job -CoalesceDuplicateWorkbooks (ConvertTo-BoolValue (Get-EngineAppSettings).coalesceDuplicateWorkbooks $true) | Out-Null
                    }
                }
                'DeclineTriggeredRun' {
                    $ruleId = [string]$command.RuleId
                    if ($shared.PendingApprovalRuleIds.ContainsKey($ruleId)) { $shared.PendingApprovalRuleIds.Remove($ruleId) }
                    $rule = Find-RuleById -RuleId $ruleId
                    $ruleName = $(if ($null -ne $rule) { [string]$rule.name } else { $ruleId })
                    Write-AppLog -Level 'INFO' -RuleName $ruleName -Message 'Triggered refresh skipped by user before Excel was opened.'
                }
                'CancelJob' {
                    $shared.CancelCurrentJob = $true
                    Write-AppLog -Level 'WARN' -Message 'Cancellation requested for the running job.'
                }
                'Exit' {
                    $shared.ShouldExit = $true
                    Write-AppLog -Level 'INFO' -Message 'Shutdown requested.'
                }
                default {
                    Write-AppLog -Level 'DEBUG' -Message ('Unknown command ignored: {0}' -f $command.Type)
                }
            }
        }
        catch {
            Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' `
                -Message ('Command "{0}" failed: {1}' -f $command.Type, $_.Exception.Message)
        }
    }

    return $reloaded
}

function Invoke-EngineEventIntake {
    <#  Watcher events -> debounce bookkeeping.  #>
    $shared = $script:EngineShared
    $events = Read-WatcherEvents
    if ($events.Count -eq 0) { return }

    foreach ($watcherEvent in $events) {
        $rule = Find-RuleById -RuleId ([string]$watcherEvent.RuleId)
        if ($null -eq $rule) { continue }

        if ($watcherEvent.EventName -eq 'Error') {
            Write-AppLog -Level 'ERROR' -RuleName $rule.name -ErrorType 'WatcherError' `
                -Message ('Watcher error: {0}. It will be recreated by the next health check.' -f $watcherEvent.Message)
            Set-RuleState -Shared $shared -RuleId ([string]$rule.id) -Values @{ WatcherStatus = 'Error' } | Out-Null
            continue
        }

        if ($shared.Paused) { continue }
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }

        Write-AppLog -Level 'DEBUG' -RuleName $rule.name `
            -Message ('Raw watcher event: {0} {1}' -f $watcherEvent.ChangeType, $watcherEvent.FullPath)

        if (Register-TriggerEvent -Rule $rule -FullPath $watcherEvent.FullPath -ChangeType $watcherEvent.ChangeType) {
            if (ConvertTo-BoolValue (Get-EngineAppSettings).showTriggerNotifications $false) {
                Request-Notification -Shared $shared -Level 'Info' -Title 'Trigger detected' `
                    -Text ('{0}: {1}' -f $rule.name, (Split-Path -Leaf $watcherEvent.FullPath))
            }
        }
    }
}

function Test-RuleNeedsPreRefreshApproval {
    param([hashtable]$Rule, [string]$Source)
    if ($null -eq $Rule) { return $false }
    if ($Source -notin @('File', 'Schedule', 'Logon')) { return $false }
    $alwaysAsk = ($Source -in @('File', 'Schedule')) -and (ConvertTo-BoolValue $Rule.askBeforeRefresh $false)
    $recentGuard = (ConvertTo-IntValue $Rule.trigger.recentRefreshPromptMinutes 0 0) -gt 0
    return ($alwaysAsk -or $recentGuard)
}

function Request-PreRefreshApproval {
    <#
        Engine -> UI hand-off for a rule that is configured to ask before an
        automatic refresh. Nothing is queued and Excel is not opened until the
        UI replies with ApproveTriggeredRun.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$Source,
        [string]$TriggerFile = '',
        [string[]]$TriggerFiles = @()
    )

    $shared = $script:EngineShared
    $ruleId = [string]$Rule.id
    if ($shared.PendingApprovalRuleIds.ContainsKey($ruleId)) {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name -Message 'Another trigger occurred while this rule is already waiting for approval; one pending approval is kept.'
        return
    }

    $requestId = New-ShortId -Prefix 'ASK'
    $shared.PendingApprovalRuleIds[$ruleId] = $requestId
    $workbookPaths = @(@($Rule.actions) | ForEach-Object { [string]$_.path } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $workbooks = @($workbookPaths | ForEach-Object { Split-Path -Leaf $_ })
    $alwaysAsk = ($Source -in @('File', 'Schedule')) -and (ConvertTo-BoolValue $Rule.askBeforeRefresh $false)
    $recentMinutes = ConvertTo-IntValue $Rule.trigger.recentRefreshPromptMinutes 0 0
    $shared.RefreshApprovalRequests.Enqueue(@{
        RequestId    = $requestId
        RuleId       = $ruleId
        RuleName     = [string]$Rule.name
        Source       = $Source
        TriggerFile  = $TriggerFile
        TriggerFiles = @($TriggerFiles)
        TriggeredAt  = (Get-Date)
        Workbooks    = $workbooks
        WorkbookPaths = $workbookPaths
        AlwaysAsk    = $alwaysAsk
        RecentRefreshPromptMinutes = $recentMinutes
    })
    $waitReason = $(if ($recentMinutes -gt 0) { 'Checking whether a workbook was refreshed recently.' } else { 'Waiting for user confirmation.' })
    Write-AppLog -Level 'INFO' -RuleName $Rule.name -Stage 'AwaitingApproval' -Message ('Trigger is ready. {0}' -f $waitReason)
}

function Queue-TriggeredRuleRun {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$Source,
        [string]$TriggerFile = '',
        [string[]]$TriggerFiles = @()
    )
    if (Test-RuleNeedsPreRefreshApproval -Rule $Rule -Source $Source) {
        Request-PreRefreshApproval -Rule $Rule -Source $Source -TriggerFile $TriggerFile -TriggerFiles $TriggerFiles
        return
    }
    $job = New-RefreshJob -Rule $Rule -TriggerSource $Source -TriggerFile $TriggerFile -TriggerFiles $TriggerFiles
    Add-RefreshJob -Job $job -CoalesceDuplicateWorkbooks (ConvertTo-BoolValue (Get-EngineAppSettings).coalesceDuplicateWorkbooks $true) | Out-Null
}

function Invoke-EngineTriggerDispatch {
    <#  Debounce elapsed -> create jobs.  #>
    $shared = $script:EngineShared
    if ($shared.Paused) { return }

    $due = Get-DueTriggers -Rules (Get-EngineRules)
    foreach ($item in $due) {
        $rule = $item.Rule
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }

        Queue-TriggeredRuleRun -Rule $rule -Source 'File' -TriggerFile ([string]$item.FilePath) -TriggerFiles @($item.FilePaths)
    }
}

function Invoke-EngineJobPump {
    <#
        Runs at most one job per loop iteration. Blocking here is fine and
        intended: the UI lives in another runspace and keeps painting, and
        watcher events keep queueing while we are inside Excel.
    #>
    $shared = $script:EngineShared
    if ((Get-JobQueueCount) -eq 0) { return }
    if ($shared.ShouldExit) { return }

    $job = Get-NextJob
    if ($null -eq $job) { return }

    $shared.CancelCurrentJob = $false
    $shared.ForcedExcelTermination = $false
    $shouldAbort = { return ([bool]$script:EngineShared.ShouldExit -or [bool]$script:EngineShared.CancelCurrentJob) }

    Update-AppStatus
    try {
        Invoke-JobExecution -Job $job -Shared $shared -AppSettings (Get-EngineAppSettings) -ShouldAbort $shouldAbort | Out-Null
    }
    catch {
        Write-AppLog -Level 'ERROR' -RuleName $job.RuleName -ErrorType 'UnexpectedError' `
            -Message ('The job runner threw an unhandled exception: {0}' -f $_.Exception.Message)
    }
    finally {
        $shared.CancelCurrentJob = $false
        $shared.CancelRequestedAt = $null
        $shared.ForcedExcelTermination = $false
        Set-CurrentJobDisplay -Job $null
        Save-AppState -Path $script:EnginePaths.StatePath -Shared $shared
        $script:LastStateSave = Get-Date
        Update-AppStatus
    }
}

function Invoke-EngineScheduleCheck {
    <#
        Time-based triggers. This is arithmetic on the clock only - no folder is
        read and no file is touched - so it is cheap enough to run on a timer
        every 30 seconds. Due schedules go through the normal job queue, exactly
        like a file event or a manual run.
    #>
    $shared = $script:EngineShared
    if ($null -eq $shared -or $shared.Paused) { return }

    $now      = Get-Date
    $interval = ConvertTo-IntValue (Get-EngineAppSettings).scheduleCheckSeconds 30 5
    if (($now - $script:LastScheduleCheck).TotalSeconds -lt $interval) { return }
    $script:LastScheduleCheck = $now

    $window = $interval + $script:ScheduleCatchUpSeconds
    $today  = $now.DayOfWeek.ToString()

    foreach ($rule in (Get-EngineRules)) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (-not (Test-TriggerIsScheduled $rule.trigger.type)) { continue }

        $ruleId = [string]$rule.id
        $time   = ConvertTo-ScheduleTime ([string]$rule.trigger.scheduleTime)
        if ($null -eq $time) {
            Set-RuleState -Shared $shared -RuleId $ruleId -Values @{ WatcherStatus = 'Misconfigured' } | Out-Null
            continue
        }
        if (@($rule.trigger.scheduleDays) -notcontains $today) { continue }

        $due = $now.Date.AddHours($time.Hour).AddMinutes($time.Minute)
        if ($now -lt $due) { continue }

        $lateBy = ($now - $due).TotalSeconds
        $key    = $due.ToString('yyyy-MM-dd HH:mm')
        $state  = Set-RuleState -Shared $shared -RuleId $ruleId
        if ([string]$state.LastScheduledRun -eq $key) { continue }

        if ($lateBy -gt $window) {
            # The application was not running at the scheduled moment. Record it
            # so the same slot is not picked up later in the day.
            Set-RuleState -Shared $shared -RuleId $ruleId -Values @{ LastScheduledRun = $key } | Out-Null
            Write-AppLog -Level 'DEBUG' -RuleName $rule.name `
                -Message ('Scheduled run for {0} was missed by {1:0} seconds and will not be started now.' -f $key, $lateBy)
            continue
        }

        Set-RuleState -Shared $shared -RuleId $ruleId -Values @{
            LastScheduledRun = $key
            LastTrigger      = $now
        } | Out-Null

        Write-AppLog -Level 'INFO' -RuleName $rule.name `
            -Message ('Scheduled run due ({0}).' -f [string]$rule.trigger.scheduleTime)

        Queue-TriggeredRuleRun -Rule $rule -Source 'Schedule'
    }
}

function Invoke-EngineHealthCheck {
    $shared = $script:EngineShared
    $settings = Get-EngineAppSettings
    $intervalSeconds = ConvertTo-IntValue $settings.watcherHealthCheckSeconds 60 10

    if (((Get-Date) - $script:LastHealthCheck).TotalSeconds -lt $intervalSeconds) { return }
    $script:LastHealthCheck = Get-Date

    try {
        $recreated = Test-WatcherHealth -Rules (Get-EngineRules) -Paused ([bool]$shared.Paused)
        if ($recreated -gt 0) {
            Write-AppLog -Level 'INFO' -Message ('Watcher health check recreated {0} watcher(s).' -f $recreated)
        }
        else {
            Write-AppLog -Level 'DEBUG' -Message 'Watcher health check completed; nothing to repair.' -NoActivity
        }
    }
    catch {
        Write-AppLog -Level 'ERROR' -ErrorType 'WatcherError' -Message ('Health check failed: {0}' -f $_.Exception.Message)
    }

    Update-AppStatus
}


function Complete-StartupWatcherInitialization {
    <#
        The splash is the only user-facing surface during startup.

        Gate 1: every enabled file/folder rule must own a live
                FileSystemWatcher with EnableRaisingEvents = $true.
        Gate 2: the UI must finish its initial workbook-information scan.

        The engine keeps revalidating Gate 1 while the UI performs Gate 2.
        StartupReady is set only after a final watcher check and only while the
        shared application status is Running. There is intentionally no time
        limit: a slow VPN or network share keeps the splash visible instead of
        exposing a Degraded Dashboard.
    #>
    param(
        [int]$DelayMilliseconds = 500,
        [int]$StableSamplesRequired = 3
    )

    $shared = $script:EngineShared
    if ($null -eq $shared) { return $false }

    $shared.StartupMonitoringReady = $false
    $shared.StartupReady = $false
    $stableSamples = 0
    $lastProgressLog = [DateTime]::MinValue

    while (-not (ConvertTo-BoolValue $shared.ShouldExit $false)) {
        $rules = Get-EngineRules

        if (ConvertTo-BoolValue $shared.Paused $false) {
            # A paused state is not a valid "ready" state for first paint.
            # This path is defensive; normal launches always begin unpaused.
            $shared.StartupMonitoringReady = $false
            $shared.StartupReady = $false
            $shared.Status = 'Starting'
            $shared.StartupMessage = 'Waiting for monitoring to resume...'
            $shared.StatusDetail = 'Startup is waiting for monitoring to be enabled.'
            Start-Sleep -Milliseconds ([Math]::Max(250, $DelayMilliseconds))
            continue
        }

        # Sync validates healthy existing watchers and creates/recreates any
        # missing watcher. QuietUnavailable avoids a warning every half-second
        # while the splash already tells the user which network location is
        # still unavailable.
        Sync-RuleWatchers -Rules $rules -Paused $false -QuietUnavailable

        $summary = Get-WatcherSummary -Rules $rules
        $expected = [int]$summary.Expected
        $active   = [int]$summary.Active
        $shared.StartupTotal   = $expected
        $shared.StartupCurrent = [Math]::Min($active, $expected)

        $allMonitorsReady = ($active -ge $expected)
        if ($allMonitorsReady) {
            $stableSamples++

            # Require several consecutive healthy observations before the UI
            # moves on to workbook metadata. This prevents a watcher that was
            # only momentarily constructible from being treated as ready.
            if ($stableSamples -ge [Math]::Max(1, $StableSamplesRequired)) {
                $shared.StartupMonitoringReady = $true

                if (ConvertTo-BoolValue $shared.StartupUiReady $false) {
                    # Final verification after workbook metadata has finished.
                    # A watcher may have disappeared while that scan was busy.
                    Sync-RuleWatchers -Rules $rules -Paused $false -QuietUnavailable
                    $finalSummary = Get-WatcherSummary -Rules $rules
                    $shared.StartupTotal   = [int]$finalSummary.Expected
                    $shared.StartupCurrent = [int]$finalSummary.Active

                    if ([int]$finalSummary.Active -ge [int]$finalSummary.Expected) {
                        Update-AppStatus
                        if ([string]$shared.Status -eq 'Running') {
                            $shared.StartupMessage = 'Monitoring and workbook information are ready.'
                            $shared.StatusDetail = '{0} watcher(s) active. Startup complete.' -f [int]$finalSummary.Active
                            $shared.StartupReady = $true
                            return $true
                        }
                    }

                    # The final check failed. Go back to the monitoring phase;
                    # the UI remains on the splash and will report what is
                    # being repaired.
                    $shared.StartupMonitoringReady = $false
                    $stableSamples = 0
                }
                else {
                    # Monitoring is ready but the UI is still loading workbook
                    # metadata. Keep Status=Starting so "Running" has one clear
                    # meaning: the Dashboard may now be shown.
                    $shared.Status = 'Starting'
                    $shared.StartupMessage = 'Monitoring ready. Reading workbook information...'
                    $shared.StatusDetail = '{0} of {1} monitor(s) ready. Waiting for workbook information.' -f $active, $expected
                }
            }
            else {
                $shared.StartupMonitoringReady = $false
                $shared.Status = 'Starting'
                $shared.StartupMessage = 'Verifying file and folder monitors...'
                $shared.StatusDetail = '{0} of {1} monitor(s) active. Verifying stability...' -f $active, $expected
            }
        }
        else {
            $stableSamples = 0
            $shared.StartupMonitoringReady = $false
            $shared.StartupReady = $false
            $shared.Status = 'Starting'

            $inactiveText = @($summary.Inactive) -join '; '
            if ([string]::IsNullOrWhiteSpace($inactiveText)) {
                $inactiveText = 'waiting for Windows/network monitoring'
            }
            $shared.StartupMessage = 'Waiting for file and folder monitors...'
            $shared.StatusDetail = '{0} of {1} monitor(s) ready. Waiting for: {2}' -f `
                $active, $expected, $inactiveText
        }

        if ($lastProgressLog -eq [DateTime]::MinValue -or
            ((Get-Date) - $lastProgressLog).TotalSeconds -ge 30) {
            Write-AppLog -Level 'INFO' -Message `
                ('Startup monitoring: {0}' -f [string]$shared.StatusDetail) -NoActivity
            $lastProgressLog = Get-Date
        }

        Start-Sleep -Milliseconds ([Math]::Max(250, $DelayMilliseconds))
    }

    return $false
}

function Start-Engine {
    <#
        Entry point invoked inside the engine runspace. Runs until the UI sets
        ShouldExit. Every iteration is wrapped so that no single failure can
        take the monitoring application down.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$Paths
    )

    $script:EngineShared = $Shared
    $script:EnginePaths  = $Paths
    $Shared.EngineAlive  = $true
    $Shared.StartupMessage = 'Restoring the previous rule state...'
    $startupPhase = 'loading configuration'

    try {
        $bootConfig = Import-AppConfiguration -Path $Paths.ConfigPath
        $startupPhase = 'initializing logging'
        Initialize-LogManager -LogDirectory $Paths.LogDir -HistoryPath $Paths.HistoryPath -Shared $Shared `
            -RetentionDays (ConvertTo-IntValue $bootConfig.appSettings.logRetentionDays 30 1) `
            -MaxTotalMegabytes (ConvertTo-IntValue $bootConfig.appSettings.logMaxTotalMegabytes 50 1) `
            -DebugLogging (ConvertTo-BoolValue $bootConfig.appSettings.debugLogging $false) `
            -SourceTag 'engine'

        Write-AppLog -Level 'INFO' -Message '--- Excel Query Trigger engine starting ---'

        $startupPhase = 'initializing monitoring and job queues'
        $Shared.StartupMessage = 'Preparing file and folder monitoring...'
        Initialize-TriggerManager -Shared $Shared
        Initialize-JobManager -Shared $Shared

        $startupPhase = 'restoring previous rule state'
        $Shared.StartupMessage = 'Restoring the previous rule state...'
        Restore-RuleStateFromDisk -Shared $Shared -State (Import-AppState -Path $Paths.StatePath)
        $startupPhase = 'loading rules and activating monitors'
        Invoke-EngineReload -Silent

        $startupPhase = 'verifying file and folder monitors and workbook information'
        $startupCompleted = Complete-StartupWatcherInitialization
        if (-not $startupCompleted -and -not (ConvertTo-BoolValue $Shared.ShouldExit $false)) {
            throw 'Startup readiness ended before monitoring reached Running state.'
        }
        if ($startupCompleted -and [string]$Shared.Status -eq 'Running') {
            Write-AppLog -Level 'INFO' -Message ('Engine ready. {0} rule(s) loaded. Dashboard may open.' -f (Get-EngineRules).Count)
        }

        $pollMilliseconds = ConvertTo-IntValue (Get-EngineAppSettings).enginePollMilliseconds 500 100
        $script:LastHealthCheck = Get-Date
        $script:LastStateSave   = Get-Date

        while (-not $Shared.ShouldExit) {
            try {
                Invoke-EngineCommands | Out-Null

                if ($Shared.ReloadRequested) {
                    $Shared.ReloadRequested = $false
                    Invoke-EngineReload
                }

                Invoke-EngineEventIntake
                Invoke-EngineTriggerDispatch
                Invoke-EngineScheduleCheck
                Invoke-EngineJobPump
                Invoke-EngineHealthCheck

                if (((Get-Date) - $script:LastStateSave).TotalSeconds -ge 120) {
                    Save-AppState -Path $Paths.StatePath -Shared $Shared
                    $script:LastStateSave = Get-Date
                }

                Update-AppStatus
            }
            catch {
                # A failure inside one iteration must never end monitoring.
                try {
                    Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' `
                        -Message ('Engine iteration failed: {0}' -f $_.Exception.Message)
                }
                catch { }
                Start-Sleep -Milliseconds 1000
            }

            Start-Sleep -Milliseconds $pollMilliseconds
        }
    }
    catch {
        $Shared.FatalError = '{0} during {1}: {2}' -f $_.Exception.GetType().Name, $startupPhase, $_.Exception.Message
        $Shared.Status     = 'Error'
        try {
            $position = [string]$_.InvocationInfo.PositionMessage
            $message = 'Engine stopped unexpectedly: {0}' -f $Shared.FatalError
            if (-not [string]::IsNullOrWhiteSpace($position)) { $message += ' ' + ($position -replace '[\r\n]+', ' ') }
            Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' -Message $message
        }
        catch { }
    }
    finally {
        try { Write-AppLog -Level 'INFO' -Message 'Engine shutting down: disposing watchers.' } catch { }
        try { Stop-AllRuleWatchers } catch { }
        try { Clear-JobQueue } catch { }
        try { Save-AppState -Path $Paths.StatePath -Shared $Shared } catch { }
        try { Write-AppLog -Level 'INFO' -Message '--- Excel Query Trigger engine stopped ---' } catch { }
        $Shared.EngineAlive = $false
        $Shared.Status      = 'Stopped'
    }
}
