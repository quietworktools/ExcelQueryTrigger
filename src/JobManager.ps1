# ==============================================================================
#  JobManager.ps1  (Engine runspace only)
#
#  A trigger never touches Excel. It creates a Job and puts it on the queue.
#  One worker drains that queue strictly sequentially - Excel COM is far more
#  reliable when only one automation instance is alive at a time.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:JobShared    = $null

function Initialize-JobManager {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:JobShared = $Shared
}

# The queue itself lives in Common.ps1 and is shared with the dashboard, so a
# manual run added while Excel is busy shows up immediately and can be
# reordered or removed. These wrappers keep the engine's call sites unchanged.

function New-RefreshJob {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [string]$TriggerSource = 'File',
        [string]$TriggerFile = '',
        [string[]]$TriggerFiles = @()
    )
    return (New-PendingJob -Shared $script:JobShared -Rule $Rule -TriggerSource $TriggerSource `
        -TriggerFile $TriggerFile -TriggerFiles $TriggerFiles)
}

function Get-JobWorkbookPaths {
    param([Parameter(Mandatory = $true)][hashtable]$Job)
    return (Get-PendingJobWorkbookPaths -Job $Job)
}

function Add-RefreshJob {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Job,
        [bool]$CoalesceDuplicateWorkbooks = $true
    )
    return (Add-PendingJob -Shared $script:JobShared -Job $Job -CoalesceDuplicateWorkbooks $CoalesceDuplicateWorkbooks)
}

function Get-NextJob      { return (Get-NextPendingJob  -Shared $script:JobShared) }
function Get-JobQueueCount { return (Get-PendingJobCount -Shared $script:JobShared) }
function Clear-JobQueue    { Clear-PendingJobs -Shared $script:JobShared }

function Set-CurrentJobDisplay {
    param(
        [hashtable]$Job,
        [string]$Stage = '',
        [string]$Workbook = ''
    )

    if ($null -eq $script:JobShared) { return }
    $current = $script:JobShared.CurrentJob

    if ($null -eq $Job) {
        $current.Active       = $false
        $current.ExcelVisible = $false
        $current.CanCancel = $false
        $current.CanForceTerminate = $false
        $current.CanInspectDialogs = $false
        $current.JobId     = ''
        $current.RuleId    = ''
        $current.RuleName  = ''
        $current.Workbook  = ''
        $current.Stage     = 'Idle'
        $current.StartedAt = $null
        $current.Source    = ''
        $current.QueryTotal          = 0
        $current.QueryCompleted      = 0
        $current.QueryPosition       = 0
        $current.QueryName           = ''
        $current.QueryElapsedSeconds = 0
        $current.QueryProgressDetail = ''
        $current.SaveStartedAt      = $null
        $current.SaveMode           = ''
        $current.SaveProgressDetail = ''
        return
    }

    $current.Active   = $true
    $current.JobId    = [string]$Job.Id
    $current.RuleId   = [string]$Job.RuleId
    $current.RuleName = [string]$Job.RuleName
    $current.Source   = $(if ([string]::IsNullOrWhiteSpace([string]$Job.TriggerFile)) { [string]$Job.TriggerSource } else { Split-Path -Leaf ([string]$Job.TriggerFile) })
    if ($null -eq $current.StartedAt) { $current.StartedAt = $Job.StartedAt }
    if (-not [string]::IsNullOrWhiteSpace($Stage)) {
        $current.Stage = $Stage
        $Job.Stage = $Stage
        $current.CanCancel = ($Stage -in @('Pending', 'WaitingForFile', 'OpeningExcel', 'Refreshing', 'RefreshingLong'))
        $current.CanForceTerminate = ($Stage -in @('OpeningExcel', 'Refreshing', 'RefreshingLong'))
        $current.CanInspectDialogs = ($Stage -in @('OpeningExcel', 'Refreshing', 'RefreshingLong', 'Saving'))
        if ($Stage -eq 'OpeningExcel') {
            $current.QueryTotal          = 0
            $current.QueryCompleted      = 0
            $current.QueryPosition       = 0
            $current.QueryName           = ''
            $current.QueryElapsedSeconds = 0
            $current.QueryProgressDetail = ''
            $current.SaveStartedAt      = $null
            $current.SaveMode           = ''
            $current.SaveProgressDetail = ''
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Workbook)) { $current.Workbook = (Split-Path -Leaf $Workbook) }
}

function Invoke-JobExecution {
    <#
        Runs one job to completion: source-file readiness, then every Excel
        action in order. Returns the finished job (Status is Completed,
        CompletedWithErrors, Failed or Cancelled).
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Job,
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$AppSettings,
        [scriptblock]$ShouldAbort
    )

    $Job.StartedAt = Get-Date
    $Job.Status    = 'Running'

    $Shared.CurrentJob.StartedAt = $Job.StartedAt
    Set-CurrentJobDisplay -Job $Job -Stage 'Pending'

    Write-AppLog -Level 'INFO' -RuleName $Job.RuleName -Stage 'JobStarted' `
        -Message ('{0} started ({1} trigger, {2} workbook action(s)).' -f $Job.Id, $Job.TriggerSource, @($Job.Actions).Count)

    $failureCount = 0
    $successCount = 0

    try {
        # ---- stage 1: is the source file finished being written? ----------
        $waitForReady = ConvertTo-BoolValue $Job.Trigger.waitForReady $true
        $triggerFiles = @(@($Job.TriggerFiles) + @($Job.TriggerFile) | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        } | Select-Object -Unique)
        if ($waitForReady -and $triggerFiles.Count -gt 0) {
            Set-CurrentJobDisplay -Job $Job -Stage 'WaitingForFile'
            foreach ($triggerFile in $triggerFiles) {
                Write-AppLog -Level 'INFO' -RuleName $Job.RuleName -Stage 'WaitingForFile' `
                    -Message ('Waiting for the source file to be released: {0}' -f (Split-Path -Leaf $triggerFile))

                $readiness = Test-SourceFileReady -Path $triggerFile `
                    -IntervalSeconds (ConvertTo-IntValue $Job.Trigger.readyCheckIntervalSeconds 2 1) `
                    -TimeoutSeconds (ConvertTo-IntValue $Job.Trigger.readyTimeoutSeconds 60 1) `
                    -ShouldAbort $ShouldAbort

                if (-not $readiness.Ready) {
                    $Job.Status = $(if ($null -ne $ShouldAbort -and (& $ShouldAbort)) { 'Cancelled' } else { 'Failed' })
                    $errorType = $(if ($Job.Status -eq 'Cancelled') { 'Cancelled' } else { 'FileReadyTimeout' })
                    Write-AppLog -Level $(if ($Job.Status -eq 'Cancelled') { 'WARN' } else { 'ERROR' }) `
                        -RuleName $Job.RuleName -Stage 'WaitingForFile' -ErrorType $errorType -Message $readiness.Reason
                    [void]$Job.Results.Add(@{
                        Workbook = ''; Success = $false; ErrorType = $errorType
                        Message = $readiness.Reason; ElapsedSeconds = $readiness.WaitedSeconds; Saved = $false
                    })
                    Set-RuleState -Shared $Shared -RuleId $Job.RuleId -Values @{
                        LastError = $readiness.Reason; LastResultText = ('{0}: source file not ready' -f $Job.Status)
                    } | Out-Null
                    if ($Job.Status -ne 'Cancelled' -and (ConvertTo-BoolValue $AppSettings.showErrorNotifications $true)) {
                        Request-Notification -Shared $Shared -Level 'Error' -Title 'Refresh failed' -Text ('{0}: {1}' -f $Job.RuleName, $readiness.Reason)
                    }
                    return $Job
                }

                Write-AppLog -Level 'INFO' -RuleName $Job.RuleName -Stage 'WaitingForFile' `
                    -Message ('Source file ready after {0} second(s): {1}' -f $readiness.WaitedSeconds, (Split-Path -Leaf $triggerFile))
            }
        }

        # ---- stage 2: the Excel actions, strictly one at a time -----------
        foreach ($action in @($Job.Actions)) {
            if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
                $Job.Status = 'Cancelled'
                Write-AppLog -Level 'WARN' -RuleName $Job.RuleName -Message 'Job cancelled before all workbooks were processed.'
                return $Job
            }

            if ([string]$action.type -ne 'ExcelRefresh') {
                Write-AppLog -Level 'WARN' -RuleName $Job.RuleName -Message ('Unsupported action type "{0}" was skipped.' -f $action.type)
                continue
            }

            Set-CurrentJobDisplay -Job $Job -Stage 'CheckingWorkbook' -Workbook $action.path

            $onStage = {
                param([string]$Stage)
                Set-CurrentJobDisplay -Job $Job -Stage $Stage -Workbook $action.path
            }.GetNewClosure()

            $actionResult = Invoke-ExcelRefresh -Action $action -Shared $Shared -RuleName $Job.RuleName `
                -OnStage $onStage -ShouldAbort $ShouldAbort `
                -CheckDataSources (ConvertTo-BoolValue $AppSettings.checkDataSources $true)

            [void]$Job.Results.Add($actionResult)

            if ($actionResult.Success) {
                $successCount++
                $queryLogText = $(if (@($actionResult.RefreshedQueries).Count -gt 0) { @($actionResult.RefreshedQueries) -join ', ' } else { 'none reported' })
                Write-AppLog -Level 'SUCCESS' -RuleName $Job.RuleName -Workbook $actionResult.Workbook `
                    -Message ('Refresh completed successfully. Queries={0} Elapsed={1} Saved={2}' -f $queryLogText, (Format-Elapsed ([TimeSpan]::FromSeconds($actionResult.ElapsedSeconds))), $(if ($actionResult.Saved) { 'Yes' } else { 'No' }))
            }
            else {
                $failureCount++
                $level = 'ERROR'
                if ($actionResult.ErrorType -eq 'Cancelled') { $level = 'WARN' }
                Write-AppLog -Level $level -RuleName $Job.RuleName -Workbook $actionResult.Workbook `
                    -Stage $Job.Stage -ErrorType $actionResult.ErrorType -Message $actionResult.Message

                if ($actionResult.ErrorType -eq 'Cancelled') {
                    $Job.Status = 'Cancelled'
                    return $Job
                }

                if (-not (ConvertTo-BoolValue $action.continueOnError $true)) {
                    Write-AppLog -Level 'WARN' -RuleName $Job.RuleName `
                        -Message 'Remaining workbooks in this job were skipped because "continue on error" is off.'
                    break
                }
            }
        }

        if ($failureCount -eq 0) { $Job.Status = 'Completed' }
        elseif ($successCount -gt 0) { $Job.Status = 'CompletedWithErrors' }
        else { $Job.Status = 'Failed' }
    }
    catch {
        $Job.Status = 'Failed'
        Write-AppLog -Level 'ERROR' -RuleName $Job.RuleName -ErrorType 'UnexpectedError' `
            -Message ('Job aborted: {0}' -f $_.Exception.Message)
        [void]$Job.Results.Add(@{
            Workbook = ''; Success = $false; ErrorType = 'UnexpectedError'
            Message = $_.Exception.Message; ElapsedSeconds = 0; Saved = $false
        })
    }
    finally {
        $Job.FinishedAt = Get-Date
        Complete-Job -Job $Job -Shared $Shared -AppSettings $AppSettings
        Set-CurrentJobDisplay -Job $null
    }

    return $Job
}

function Complete-Job {
    <#  Writes the history record, the rule state and the notification.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Job,
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$AppSettings
    )

    $elapsed = [TimeSpan]::Zero
    if ($null -ne $Job.StartedAt -and $null -ne $Job.FinishedAt) { $elapsed = $Job.FinishedAt - $Job.StartedAt }

    $workbookSummaries = New-Object System.Collections.ArrayList
    foreach ($actionResult in @($Job.Results)) {
        [void]$workbookSummaries.Add(@{
            workbook  = [string]$actionResult.Workbook
            success   = [bool]$actionResult.Success
            errorType = [string]$actionResult.ErrorType
            message   = [string]$actionResult.Message
            elapsed   = [int]$actionResult.ElapsedSeconds
            saved     = [bool]$actionResult.Saved
            refreshMethod = [string]$actionResult.RefreshMethod
            queries    = @($actionResult.RefreshedQueries)
            queryTimings = $(if ($actionResult.ContainsKey('QueryTimings')) { @($actionResult.QueryTimings) } else { @() })
            refreshSeconds = $(if ($actionResult.ContainsKey('RefreshSeconds')) { [int]$actionResult.RefreshSeconds } else { 0 })
            refreshCallSeconds = $(if ($actionResult.ContainsKey('RefreshCallSeconds')) { [int]$actionResult.RefreshCallSeconds } else { 0 })
            refreshConfirmationSeconds = $(if ($actionResult.ContainsKey('RefreshConfirmationSeconds')) { [int]$actionResult.RefreshConfirmationSeconds } else { 0 })
            saveMode = $(if ($actionResult.ContainsKey('SaveMode')) { [string]$actionResult.SaveMode } else { 'Normal' })
            saveSeconds = $(if ($actionResult.ContainsKey('SaveSeconds')) { [int]$actionResult.SaveSeconds } else { 0 })
            queryRefreshedAt = $(if ($null -ne $actionResult.QueryRefreshedAt) {
                ([DateTime]$actionResult.QueryRefreshedAt).ToString('yyyy-MM-dd HH:mm:ss')
            } else { '' })
        })

        # Successful automatic and manual runs both update the Dashboard's
        # Data updated field. The UI performs the file read in its background
        # worker, so this hand-off never holds up the engine or Excel cleanup.
        if ([bool]$actionResult.Success -and
            -not [string]::IsNullOrWhiteSpace([string]$actionResult.Workbook) -and
            $null -ne $Shared.WorkbookInfoRefreshRequests) {
            $Shared.WorkbookInfoRefreshRequests.Enqueue([string]$actionResult.Workbook)
        }
    }

    $record = @{
        jobId         = [string]$Job.Id
        ruleId        = [string]$Job.RuleId
        rule          = [string]$Job.RuleName
        triggerType   = [string]$Job.TriggerType
        triggerSource = [string]$Job.TriggerSource
        triggerFile   = [string]$Job.TriggerFile
        triggeredAt   = $Job.TriggeredAt.ToString('yyyy-MM-dd HH:mm:ss')
        startedAt     = $(if ($null -ne $Job.StartedAt) { $Job.StartedAt.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
        finishedAt    = $(if ($null -ne $Job.FinishedAt) { $Job.FinishedAt.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
        elapsed       = (Format-Elapsed $elapsed)
        status        = [string]$Job.Status
        stage         = [string]$Job.Stage
        results       = @($workbookSummaries.ToArray())
    }
    Write-JobHistory -Record $record

    $succeeded = @($Job.Results | Where-Object { $_.Success }).Count
    $total     = @($Job.Results).Count
    $summary   = '{0}: {1}/{2} workbook(s) refreshed in {3}' -f $Job.RuleName, $succeeded, $total, (Format-Elapsed $elapsed)
    if ($Job.Results.Count -eq 1 -and $Job.Results[0].Success -and @($Job.Results[0].RefreshedQueries).Count -gt 0) {
        $queryNames = @($Job.Results[0].RefreshedQueries)
        $querySummary = $(if ($queryNames.Count -le 3) { $queryNames -join ', ' } else { '{0}, {1}, {2} (+{3} more)' -f $queryNames[0],$queryNames[1],$queryNames[2],($queryNames.Count-3) })
        $summary += (' | Queries: {0}' -f $querySummary)
    }

    $stateValues = @{ LastResultText = ('{0} ({1})' -f $Job.Status, (Format-Elapsed $elapsed)); LastRun = $Job.FinishedAt; LastDuration = (Format-Elapsed $elapsed); LastRunStatus = [string]$Job.Status }
    if ($Job.Status -eq 'Completed') {
        $stateValues['LastSuccess'] = $Job.FinishedAt
        $stateValues['LastError']   = ''
        Write-AppLog -Level 'SUCCESS' -RuleName $Job.RuleName -Message ('{0} completed. {1}' -f $Job.Id, $summary)
        if (ConvertTo-BoolValue $AppSettings.showSuccessNotifications $true) {
            Request-Notification -Shared $Shared -Level 'Info' -Title 'Refresh completed' -Text $summary
        }
    }
    else {
        $firstError = @($Job.Results | Where-Object { -not $_.Success } | Select-Object -First 1)
        $errorText = 'Job did not complete'
        if ($firstError.Count -gt 0) { $errorText = '{0}: {1}' -f $firstError[0].ErrorType, $firstError[0].Message }
        $stateValues['LastError'] = $errorText
        Write-AppLog -Level 'ERROR' -RuleName $Job.RuleName -Message ('{0} finished with status {1}. {2}' -f $Job.Id, $Job.Status, $errorText)
        if (ConvertTo-BoolValue $AppSettings.showErrorNotifications $true) {
            $notificationTitle = 'Refresh failed'
            $notificationLevel = 'Error'
            $notificationText  = ('{0} - {1}' -f $Job.RuleName, $errorText)
            if ($firstError.Count -gt 0 -and [string]$firstError[0].ErrorType -eq 'WorkbookLocked') {
                $notificationTitle = 'Refresh not started'
                $notificationLevel = 'Warning'
                $notificationText  = ('{0}: {1}' -f $Job.RuleName, [string]$firstError[0].Message)
            }
            Request-Notification -Shared $Shared -Level $notificationLevel -Title $notificationTitle -Text $notificationText
        }
    }

    Set-RuleState -Shared $Shared -RuleId $Job.RuleId -Values $stateValues | Out-Null
}
