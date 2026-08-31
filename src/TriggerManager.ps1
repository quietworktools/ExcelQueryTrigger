# ==============================================================================
#  TriggerManager.ps1  (Engine runspace only)
#
#  Watchers are registered with Register-ObjectEvent WITHOUT -Action on purpose.
#  A -Action script block has to be executed inside a runspace; the engine
#  runspace is busy running the worker loop (and may be inside a 5 minute Excel
#  refresh), so an -Action handler would either block or be dropped. Without
#  -Action the events are pushed into the runspace event queue by the raising
#  thread and the worker loop drains them with Get-Event whenever it comes back
#  around. Nothing is lost while Excel is busy.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:Watchers      = @{}   # ruleId -> watcher entry
$script:TriggerState  = @{}   # ruleId -> debounce / cooldown bookkeeping
$script:TriggerShared = $null

function Initialize-TriggerManager {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:TriggerShared = $Shared
    $script:Watchers     = @{}
    $script:TriggerState = @{}
}

function Get-TriggerState {
    param([Parameter(Mandatory = $true)][string]$RuleId)
    if (-not $script:TriggerState.ContainsKey($RuleId)) {
        $script:TriggerState[$RuleId] = @{
            PendingFiles  = @{}
            PendingSince  = $null
            LastEventAt   = $null
            CooldownUntil = [DateTime]::MinValue
            EventCount    = 0
        }
    }
    return $script:TriggerState[$RuleId]
}

# ------------------------------------------------------------------------------
# Region: watcher lifecycle
# ------------------------------------------------------------------------------

function Get-WatchFolderForRule {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    $trigger = $Rule.trigger
    if (Test-TriggerUsesFolder $trigger.type) { return [string]$trigger.path }
    if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return '' }
    return (Split-Path -Parent ([string]$trigger.path))
}

function Get-WatchFilterForRule {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    $trigger = $Rule.trigger
    if (Test-TriggerUsesFolder $trigger.type) {
        if ([string]::IsNullOrWhiteSpace([string]$trigger.filter)) { return '*.*' }
        $configuredFilter = [string]$trigger.filter
        # FileSystemWatcher on Windows PowerShell 5.1 supports one Filter only.
        # For a preset such as Excel files we therefore watch all filenames and
        # apply the small set of wildcard patterns after Windows raises an event.
        # This remains event-driven; it does not enumerate or poll the folder.
        if ($configuredFilter.Contains(';')) { return '*.*' }
        return $configuredFilter
    }
    if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return '*.*' }
    return (Split-Path -Leaf ([string]$trigger.path))
}

function Start-RuleWatcher {
    <#  Creates and arms one FileSystemWatcher for a rule. Idempotent.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [switch]$QuietUnavailable
    )

    $ruleId = [string]$Rule.id
    Stop-RuleWatcher -RuleId $ruleId

    if (-not (Test-TriggerUsesWatcher $Rule.trigger.type)) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Manual' } | Out-Null
        return $true
    }

    $folder = Get-WatchFolderForRule -Rule $Rule
    if ([string]::IsNullOrWhiteSpace($folder)) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Misconfigured' } | Out-Null
        Write-AppLog -Level 'WARN' -RuleName $Rule.name -Message 'Watcher not started: no folder configured.'
        return $false
    }

    $reachable = $false
    try { $reachable = Test-Path -LiteralPath $folder } catch { $reachable = $false }
    if (-not $reachable) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'PathUnavailable' } | Out-Null
        if (-not $QuietUnavailable) {
            Write-AppLog -Level 'WARN' -RuleName $Rule.name -ErrorType 'PathUnavailable' `
                -Message ('Watch folder is not reachable, will retry on the next health check: {0}' -f $folder)
        }
        return $false
    }

    try {
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path                  = $folder
        $watcher.Filter                = (Get-WatchFilterForRule -Rule $Rule)
        $watcher.IncludeSubdirectories = $false
        # 64 KB buffer: network shares can burst several events for one copy.
        $watcher.InternalBufferSize    = 65536
        $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor
                                [System.IO.NotifyFilters]::LastWrite -bor
                                [System.IO.NotifyFilters]::Size

        $eventNames = @('Error')
        switch ($Rule.trigger.type) {
            'FileCreated'         { $eventNames += @('Created', 'Renamed') }
            'FileChangedAny'      { $eventNames += @('Changed') }
            'FileChangedSpecific' { $eventNames += @('Changed', 'Created', 'Renamed') }
        }

        $subscriptions = New-Object System.Collections.ArrayList
        foreach ($eventName in $eventNames) {
            $sourceIdentifier = 'EQT:{0}:{1}' -f $ruleId, $eventName
            $subscription = Register-ObjectEvent -InputObject $watcher -EventName $eventName `
                -SourceIdentifier $sourceIdentifier -MessageData $ruleId
            [void]$subscriptions.Add($sourceIdentifier)
            $subscription | Out-Null
        }

        $watcher.EnableRaisingEvents = $true

        $script:Watchers[$ruleId] = @{
            Watcher       = $watcher
            Subscriptions = @($subscriptions.ToArray())
            Folder        = $folder
            RuleName      = [string]$Rule.name
            LastError     = ''
            StartedAt     = (Get-Date)
        }

        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Active' } | Out-Null
        Write-AppLog -Level 'INFO' -RuleName $Rule.name `
            -Message ('Watcher started. Folder={0} Filter={1} Events={2}' -f $folder, $watcher.Filter, (($eventNames | Where-Object { $_ -ne 'Error' }) -join ','))
        return $true
    }
    catch {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Error' } | Out-Null
        Write-AppLog -Level 'ERROR' -RuleName $Rule.name -ErrorType 'WatcherError' `
            -Message ('Failed to start watcher: {0}' -f $_.Exception.Message)
        return $false
    }
}

function Stop-RuleWatcher {
    param([Parameter(Mandatory = $true)][string]$RuleId)

    if (-not $script:Watchers.ContainsKey($RuleId)) { return }
    $entry = $script:Watchers[$RuleId]

    try {
        if ($null -ne $entry.Watcher) {
            $entry.Watcher.EnableRaisingEvents = $false
        }
    }
    catch { }

    foreach ($sourceIdentifier in @($entry.Subscriptions)) {
        try { Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue } catch { }
        try {
            Get-Event -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue |
                ForEach-Object { Remove-Event -EventIdentifier $_.EventIdentifier -ErrorAction SilentlyContinue }
        }
        catch { }
    }

    try { if ($null -ne $entry.Watcher) { $entry.Watcher.Dispose() } } catch { }
    $script:Watchers.Remove($RuleId)
}

function Restart-RuleWatcher {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    Write-AppLog -Level 'INFO' -RuleName $Rule.name -Message 'Restarting watcher.'
    return (Start-RuleWatcher -Rule $Rule)
}

function Stop-AllRuleWatchers {
    foreach ($ruleId in @($script:Watchers.Keys)) { Stop-RuleWatcher -RuleId $ruleId }
}

function Sync-RuleWatchers {
    <#
        Brings the running watchers in line with the configuration. Called at
        startup, after every config change and after resume - so enabling or
        editing a rule never needs an application restart.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [bool]$Paused = $false,
        [switch]$QuietUnavailable
    )

    if ($null -eq $Rules) { $Rules = @() }
    $wanted = @{}
    foreach ($rule in $Rules) {
        if ((ConvertTo-BoolValue $rule.enabled $true) -and (Test-TriggerUsesWatcher $rule.trigger.type) -and -not $Paused) {
            $wanted[[string]$rule.id] = $rule
        }
    }

    foreach ($ruleId in @($script:Watchers.Keys)) {
        if (-not $wanted.ContainsKey($ruleId)) {
            Stop-RuleWatcher -RuleId $ruleId
        }
    }

    $watcherNumber = 0
    foreach ($ruleId in @($wanted.Keys)) {
        $rule = $wanted[$ruleId]
        $script:TriggerShared.StartupMessage = ('Activating monitor: {0}' -f [string]$rule.name)
        if ($script:Watchers.ContainsKey($ruleId)) {
            # Re-arm if the folder/filter changed, or if an object is present
            # but is no longer actually raising events. Sync is used by the
            # startup readiness gate, so it must validate as well as create.
            $entry = $script:Watchers[$ruleId]
            $watcherHealthy = $false
            try {
                $sameFolder = ($entry.Folder -eq (Get-WatchFolderForRule -Rule $rule))
                $sameFilter = ($entry.Watcher.Filter -eq (Get-WatchFilterForRule -Rule $rule))
                $watcherHealthy = $sameFolder -and $sameFilter -and
                    [string]::IsNullOrWhiteSpace([string]$entry.LastError) -and
                    [bool]$entry.Watcher.EnableRaisingEvents
            }
            catch { $watcherHealthy = $false }

            if ($watcherHealthy) {
                $watcherNumber++
                $script:TriggerShared.StartupCurrent = $watcherNumber
                continue
            }

            Stop-RuleWatcher -RuleId $ruleId
        }
        Start-RuleWatcher -Rule $rule -QuietUnavailable:$QuietUnavailable | Out-Null
        $watcherNumber++
        $script:TriggerShared.StartupCurrent = $watcherNumber
    }

    foreach ($rule in $Rules) {
        $ruleId = [string]$rule.id
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Disabled' } | Out-Null
        }
        elseif ($Paused -and (Test-TriggerUsesWatcher $rule.trigger.type)) {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Paused' } | Out-Null
        }
        elseif (-not (Test-TriggerUsesWatcher $rule.trigger.type)) {
            # Scheduled / logon / manual rules have no watcher, but the status
            # column should still say what the rule is waiting for.
            $status = switch ([string]$rule.trigger.type) {
                'Scheduled' { 'Waiting for time' }
                'Logon'     { 'Waiting for logon' }
                default     { 'Manual only' }
            }
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = $status } | Out-Null
        }
    }
}

# ------------------------------------------------------------------------------
# Region: event intake, filtering, debounce, cooldown
# ------------------------------------------------------------------------------

function Test-TriggerFileMatch {
    <#  Applies the conditions the FileSystemWatcher filter cannot express.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$FullPath
    )

    $trigger = $Rule.trigger
    $leaf = Split-Path -Leaf $FullPath

    if ($trigger.type -eq 'FileChangedSpecific') {
        if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return $false }
        return ([string]::Equals([System.IO.Path]::GetFullPath($FullPath), [System.IO.Path]::GetFullPath([string]$trigger.path), [StringComparison]::OrdinalIgnoreCase))
    }

    # When a rule contains several wildcard patterns (for example the Excel
    # preset), FileSystemWatcher itself is deliberately broad (*.*) and this
    # inexpensive filename test narrows the event back down here.
    if (Test-TriggerUsesFolder $trigger.type) {
        $configuredFilter = [string]$trigger.filter
        if (-not [string]::IsNullOrWhiteSpace($configuredFilter) -and $configuredFilter.Contains(';')) {
            $patternMatched = $false
            foreach ($pattern in $configuredFilter.Split(';')) {
                $trimmed = $pattern.Trim()
                if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
                if ($leaf -like $trimmed) { $patternMatched = $true; break }
            }
            if (-not $patternMatched) { return $false }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$trigger.contains)) {
        if ($leaf.IndexOf([string]$trigger.contains, [StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$trigger.exclude)) {
        if ($leaf.IndexOf([string]$trigger.exclude, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $false }
    }

    # Temporary artefacts of a copy in progress are never a trigger.
    if ($leaf.StartsWith('~$') -or $leaf.EndsWith('.tmp') -or $leaf.EndsWith('.crdownload') -or $leaf.EndsWith('.part')) {
        return $false
    }

    return $true
}

function Read-WatcherEvents {
    <#
        Drains the runspace event queue. Returns an array of
        @{ RuleId; FullPath; ChangeType; EventName } and records watcher errors.
    #>
    $results = New-Object System.Collections.ArrayList

    $queued = @()
    try { $queued = @(Get-Event -ErrorAction SilentlyContinue) } catch { $queued = @() }

    foreach ($queuedEvent in $queued) {
        $identifier = [string]$queuedEvent.SourceIdentifier
        if (-not $identifier.StartsWith('EQT:')) { continue }

        try {
            $parts     = $identifier.Split(':')
            $ruleId    = [string]$queuedEvent.MessageData
            if ([string]::IsNullOrWhiteSpace($ruleId) -and $parts.Count -ge 2) { $ruleId = $parts[1] }
            $eventName = $parts[$parts.Count - 1]

            if ($eventName -eq 'Error') {
                $message = 'Unknown watcher error'
                try { $message = $queuedEvent.SourceArgs[1].GetException().Message } catch { }
                if ($script:Watchers.ContainsKey($ruleId)) { $script:Watchers[$ruleId].LastError = $message }
                [void]$results.Add(@{ RuleId = $ruleId; EventName = 'Error'; Message = $message; FullPath = ''; ChangeType = 'Error' })
            }
            else {
                $eventArgs = $queuedEvent.SourceArgs[1]
                [void]$results.Add(@{
                    RuleId     = $ruleId
                    EventName  = $eventName
                    FullPath   = [string]$eventArgs.FullPath
                    ChangeType = [string]$eventArgs.ChangeType
                    Message    = ''
                })
            }
        }
        catch {
            Write-AppLog -Level 'DEBUG' -Message ('Malformed watcher event ignored: {0}' -f $_.Exception.Message)
        }
        finally {
            try { Remove-Event -EventIdentifier $queuedEvent.EventIdentifier -ErrorAction SilentlyContinue } catch { }
        }
    }

    $events = @($results.ToArray())
    return ,$events
}

function Register-TriggerEvent {
    <#
        Records one matching filesystem event against the rule's debounce
        window. Returns $true when the event was accepted, $false when it was
        dropped by a filter or by cooldown.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$FullPath,
        [string]$ChangeType = 'Changed'
    )

    $ruleId = [string]$Rule.id
    $state  = Get-TriggerState -RuleId $ruleId
    $now    = Get-Date

    if (-not (Test-TriggerFileMatch -Rule $Rule -FullPath $FullPath)) {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name -Message ('Event ignored by filename condition: {0}' -f $FullPath)
        return $false
    }

    if ($now -lt $state.CooldownUntil) {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name `
            -Message ('Event dropped by cooldown (until {0:HH:mm:ss}): {1}' -f $state.CooldownUntil, $FullPath)
        return $false
    }

    if ($state.PendingFiles.Count -eq 0) {
        $state.PendingSince = $now
        $state.EventCount   = 0
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ LastTrigger = $now } | Out-Null
        Write-AppLog -Level 'INFO' -RuleName $Rule.name -Stage 'TriggerDetected' `
            -Message ('Trigger detected ({0}): {1}' -f $ChangeType, (Split-Path -Leaf $FullPath))
    }
    else {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name `
            -Message ('Debounce window extended by {0} event: {1}' -f $ChangeType, (Split-Path -Leaf $FullPath))
    }

    $fileKey = $FullPath.ToLowerInvariant()
    if (-not $state.PendingFiles.ContainsKey($fileKey)) {
        $state.PendingFiles[$fileKey] = $FullPath
    }
    $state.LastEventAt = $now
    $state.EventCount  = [int]$state.EventCount + 1
    return $true
}

function Get-DueTriggers {
    <#
        Collapses everything accumulated inside the debounce window into a
        single fire per rule, and opens the cooldown window immediately so the
        tail of a noisy copy cannot queue a second job.
    #>
    param([AllowNull()][array]$Rules = @())

    $due = New-Object System.Collections.ArrayList
    $now = Get-Date

    foreach ($rule in $Rules) {
        $ruleId = [string]$rule.id
        if (-not $script:TriggerState.ContainsKey($ruleId)) { continue }

        $state = $script:TriggerState[$ruleId]
        if ($state.PendingFiles.Count -eq 0) { continue }

        $debounce = ConvertTo-IntValue $rule.trigger.debounceSeconds 5 0
        if (($now - $state.LastEventAt).TotalSeconds -lt $debounce) { continue }

        $filePaths = @($state.PendingFiles.Values)
        $filePath = $(if ($filePaths.Count -gt 0) { [string]$filePaths[0] } else { '' })
        $count    = $state.EventCount

        $state.PendingFiles  = @{}
        $state.PendingSince  = $null
        $state.EventCount    = 0
        $cooldown = ConvertTo-IntValue $rule.trigger.cooldownSeconds 30 0
        $state.CooldownUntil = $now.AddSeconds($cooldown)

        Write-AppLog -Level 'DEBUG' -RuleName $rule.name `
            -Message ('Debounce elapsed. {0} raw event(s) collapsed into one job. Cooldown until {1:HH:mm:ss}.' -f $count, $state.CooldownUntil)

        [void]$due.Add(@{ Rule = $rule; FilePath = $filePath; FilePaths = $filePaths; EventCount = $count })
    }

    $dueList = @($due.ToArray())
    return ,$dueList
}

function Clear-TriggerState {
    param([string]$RuleId)
    if ([string]::IsNullOrWhiteSpace($RuleId)) { $script:TriggerState = @{}; return }
    if ($script:TriggerState.ContainsKey($RuleId)) { $script:TriggerState.Remove($RuleId) }
}

# ------------------------------------------------------------------------------
# Region: source file readiness
# ------------------------------------------------------------------------------

function Test-SourceFileReady {
    <#
        A file that has just appeared on a share is usually still being written.
        Two independent signals must agree before we let Excel read it:
          1. the size has stopped changing between two probes
          2. the file can be opened with FileShare.Read (i.e. nobody holds a
             write handle on it any more)
        Returns @{ Ready = [bool]; Reason = [string]; WaitedSeconds = [int] }
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$IntervalSeconds = 2,
        [int]$TimeoutSeconds = 60,
        [scriptblock]$ShouldAbort
    )

    $started      = Get-Date
    $deadline     = $started.AddSeconds($TimeoutSeconds)
    $lastSize     = -1L
    $lastWrite    = [DateTime]::MinValue
    $intervalMs   = [Math]::Max(250, $IntervalSeconds * 1000)

    while ($true) {
        if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
            return @{ Ready = $false; Reason = 'Cancelled while waiting for the source file.'; WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
        }

        $info = $null
        try {
            if (Test-Path -LiteralPath $Path) { $info = Get-Item -LiteralPath $Path -ErrorAction Stop }
        }
        catch { $info = $null }

        if ($null -ne $info) {
            $stable = ($info.Length -eq $lastSize -and $info.LastWriteTimeUtc -eq $lastWrite -and $lastSize -ge 0)
            $lastSize  = $info.Length
            $lastWrite = $info.LastWriteTimeUtc

            if ($stable -and (Test-FileNotLocked -Path $Path)) {
                return @{ Ready = $true; Reason = 'Size stable and no writer holds the file.'; WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
            }
        }

        if ((Get-Date) -ge $deadline) {
            $reason = 'Source file did not become ready within {0} seconds' -f $TimeoutSeconds
            if ($null -eq $info) { $reason = 'Source file disappeared or is unreachable ({0})' -f $Path }
            return @{ Ready = $false; Reason = $reason; WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
        }

        Start-Sleep -Milliseconds $intervalMs
    }
}

function Test-FileNotLocked {
    param([Parameter(Mandatory = $true)][string]$Path)
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

# ------------------------------------------------------------------------------
# Region: watcher health
# ------------------------------------------------------------------------------

function Test-WatcherHealth {
    <#
        Cheap periodic check - no folder enumeration. Recreates watchers that
        died silently (the classic network share failure) and picks up shares
        that came back. Returns the number of watchers that were recreated.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [bool]$Paused = $false
    )

    if ($Paused) { return 0 }
    if ($null -eq $Rules) { $Rules = @() }
    $recreated = 0

    foreach ($rule in $Rules) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (-not (Test-TriggerUsesWatcher $rule.trigger.type)) { continue }

        $ruleId  = [string]$rule.id
        $folder  = Get-WatchFolderForRule -Rule $rule
        $healthy = $true
        $reason  = ''

        if (-not $script:Watchers.ContainsKey($ruleId)) {
            $healthy = $false
            $reason  = 'watcher missing'
        }
        else {
            $entry = $script:Watchers[$ruleId]
            if (-not [string]::IsNullOrWhiteSpace($entry.LastError)) {
                $healthy = $false
                $reason  = ('watcher reported an error: {0}' -f $entry.LastError)
            }
            else {
                try {
                    if (-not $entry.Watcher.EnableRaisingEvents) {
                        $healthy = $false
                        $reason  = 'EnableRaisingEvents is false'
                    }
                }
                catch {
                    $healthy = $false
                    $reason  = 'watcher object is no longer usable'
                }
            }
        }

        $reachable = $false
        try { $reachable = (-not [string]::IsNullOrWhiteSpace($folder)) -and (Test-Path -LiteralPath $folder) } catch { $reachable = $false }

        if (-not $reachable) {
            if ($script:Watchers.ContainsKey($ruleId)) {
                Write-AppLog -Level 'WARN' -RuleName $rule.name -ErrorType 'PathUnavailable' `
                    -Message ('Watch folder became unreachable, watcher disposed until it returns: {0}' -f $folder)
                Stop-RuleWatcher -RuleId $ruleId
            }
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'PathUnavailable' } | Out-Null
            continue
        }

        if (-not $healthy) {
            Write-AppLog -Level 'WARN' -RuleName $rule.name -ErrorType 'WatcherError' `
                -Message ('Watcher unhealthy ({0}). Recreating.' -f $reason)
            if (Restart-RuleWatcher -Rule $rule) { $recreated++ }
        }
        else {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Active' } | Out-Null
        }
    }

    return $recreated
}

function Get-WatcherSummary {
    <#  Used by the engine to decide between Running and Degraded.  #>
    param([AllowNull()][array]$Rules = @())

    $expected = 0
    $active   = 0
    $inactive = New-Object System.Collections.ArrayList
    foreach ($rule in $Rules) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (-not (Test-TriggerUsesWatcher $rule.trigger.type)) { continue }

        $expected++
        $ruleId = [string]$rule.id
        $isActive = $false
        $watcherStatus = 'not started'

        if ($script:Watchers.ContainsKey($ruleId)) {
            $entry = $script:Watchers[$ruleId]
            try {
                $isActive = [string]::IsNullOrWhiteSpace([string]$entry.LastError) -and
                    [bool]$entry.Watcher.EnableRaisingEvents
                if (-not $isActive) {
                    $watcherStatus = $(if (-not [string]::IsNullOrWhiteSpace([string]$entry.LastError)) {
                        'watcher error: {0}' -f [string]$entry.LastError
                    } else {
                        'watcher is not raising events'
                    })
                }
            }
            catch {
                $isActive = $false
                $watcherStatus = 'watcher object unavailable'
            }
        }
        elseif ($null -ne $script:TriggerShared -and
            $script:TriggerShared.RuleState.ContainsKey($ruleId)) {
            $watcherStatus = [string]$script:TriggerShared.RuleState[$ruleId].WatcherStatus
        }

        if ($isActive) {
            $active++
            continue
        }

        if ($watcherStatus -eq 'PathUnavailable' -and
            -not [string]::IsNullOrWhiteSpace([string]$rule.trigger.path)) {
            $watcherStatus = 'path unavailable: {0}' -f [string]$rule.trigger.path
        }
        [void]$inactive.Add(('{0} ({1})' -f [string]$rule.name, $watcherStatus))
    }

    return @{ Expected = $expected; Active = $active; Inactive = @($inactive.ToArray()) }
}
