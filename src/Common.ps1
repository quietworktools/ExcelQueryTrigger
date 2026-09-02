# ==============================================================================
#  Common.ps1
#  Shared primitives used by BOTH the UI runspace and the Engine runspace.
#  No WinForms and no Excel COM here - keep this file dependency free.
# ==============================================================================

Set-StrictMode -Version 1.0

# Single source of truth for release/config versions. Both runspaces load this file.
$script:AppName = 'Excel Query Trigger Manager'
$script:AppVersion = '1.0.3'
$script:ConfigSchemaVersion = 4

function Get-AppVersion { return [string]$script:AppVersion }

function Test-EngineStatusReady {
    <#
        The statuses in which the Dashboard may be shown and treated as
        operational. 'Waiting for network' is one of them: a laptop away from
        the office is a normal condition, not a startup failure, and the splash
        must not sit there forever waiting for a share that is not coming back
        until Monday.
    #>
    param([string]$Status)
    return (@('Running', 'Waiting for network') -contains [string]$Status)
}

function Get-SafeLeaf {
    <#
        Split-Path throws on an empty string, and several records carry an
        empty path on purpose (a job that failed before it reached a workbook
        has no workbook). Callers only ever want a display name, so give them
        one instead of an exception that reaches the UI as a red log line.
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    try { return (Split-Path -Leaf $Path) } catch { return $Path }
}
function Get-ConfigSchemaVersion { return [int]$script:ConfigSchemaVersion }

# ------------------------------------------------------------------------------
# Region: type / value helpers
# ------------------------------------------------------------------------------

function ConvertTo-HashtableDeep {
    <#
        Converts the PSCustomObject graph produced by ConvertFrom-Json into plain
        hashtables / arrays. Working with hashtables everywhere means a missing
        property is simply $null instead of a strict-mode explosion.
    #>
    param($InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in @($InputObject.Keys)) {
            $result[[string]$key] = ConvertTo-HashtableDeep $InputObject[$key]
        }
        return $result
    }

    if ($InputObject -is [string] -or $InputObject.GetType().IsValueType) {
        return $InputObject
    }

    if ($InputObject.GetType().Name -eq 'PSCustomObject') {
        $result = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $result[$property.Name] = ConvertTo-HashtableDeep $property.Value
        }
        return $result
    }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $InputObject) { [void]$list.Add((ConvertTo-HashtableDeep $item)) }
        return , ($list.ToArray())
    }

    return $InputObject
}

function Merge-DefaultValues {
    <#
        Returns a copy of $Default with every non-null key of $Value applied on
        top. Nested hashtables are merged recursively. This is what makes the
        JSON schema tolerant of missing / future properties.
    #>
    param(
        [hashtable]$Default,
        $Value
    )

    $result = @{}
    foreach ($key in $Default.Keys) { $result[$key] = $Default[$key] }

    if ($Value -is [hashtable]) {
        foreach ($key in @($Value.Keys)) {
            $incoming = $Value[$key]
            if ($null -eq $incoming) { continue }

            if ($result.ContainsKey($key) -and $result[$key] -is [hashtable] -and $incoming -is [hashtable]) {
                $result[$key] = Merge-DefaultValues -Default $result[$key] -Value $incoming
            }
            else {
                $result[$key] = $incoming
            }
        }
    }

    return $result
}


function ConvertTo-StringValue {
    param($Value, [string]$Default = '')
    if ($null -eq $Value) { return $Default }
    try { return [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return $Default }
}

function ConvertTo-BoolValue {
    param($Value, [bool]$Default = $false)
    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $Default }
    if ($text -match '^(?i:true|yes|1|on)$') { return $true }
    if ($text -match '^(?i:false|no|0|off)$') { return $false }
    return $Default
}

function ConvertTo-IntValue {
    param($Value, [int]$Default = 0, [int]$Minimum = [int]::MinValue)
    $parsed = 0
    if ($null -ne $Value -and [int]::TryParse(([string]$Value), [ref]$parsed)) {
        if ($parsed -lt $Minimum) { return $Minimum }
        return $parsed
    }
    return $Default
}

function Format-Elapsed {
    param([TimeSpan]$TimeSpan)
    if ($null -eq $TimeSpan) { return '00:00:00' }
    return ('{0:00}:{1:00}:{2:00}' -f [Math]::Floor($TimeSpan.TotalHours), $TimeSpan.Minutes, $TimeSpan.Seconds)
}

function New-ShortId {
    param([string]$Prefix = 'ID')
    return ('{0}-{1}' -f $Prefix, ([Guid]::NewGuid().ToString('N').Substring(0, 8).ToUpperInvariant()))
}

function Copy-SharedList {
    <#
        Snapshots a synchronized ArrayList under its SyncRoot so the UI thread
        can enumerate it while the engine thread is still appending to it.
    #>
    param($List)

    $locked = $false
    $copy = @()
    try {
        [System.Threading.Monitor]::Enter($List.SyncRoot, [ref]$locked)
        $copy = @($List.ToArray())
    }
    catch {
        $copy = @()
    }
    finally {
        if ($locked) { [System.Threading.Monitor]::Exit($List.SyncRoot) }
    }
    return $copy
}

function Get-StageDisplayText {
    param([string]$Stage)
    switch ($Stage) {
        'Pending'        { return 'Queued' }
        'WaitingForFile' { return 'Waiting for source file' }
        'CheckingWorkbook' { return 'Checking workbook availability' }
        'OpeningExcel'   { return 'Opening Excel' }
        'Refreshing'     { return 'Refreshing queries' }
        'RefreshingLong' { return 'Refreshing queries - taking longer than expected' }
        'Saving'         { return 'Saving workbook' }
        'Closing'        { return 'Closing workbook' }
        'Cleanup'        { return 'Releasing Excel' }
        'Completed'      { return 'Completed' }
        'Failed'         { return 'Failed' }
        'Idle'           { return 'Idle' }
        default          { return $Stage }
    }
}

function Test-DirectoryWritable {
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $Path ('.write-probe-{0}' -f ([Guid]::NewGuid().ToString('N')))
        [System.IO.File]::WriteAllText($probe, 'probe')
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

# ------------------------------------------------------------------------------
# Region: application paths
# ------------------------------------------------------------------------------

function Get-AppPaths {
    <#
        Resolves where config / logs / state live. Preference is the application
        folder itself so the tool stays portable; if that folder is read-only
        (for example when copied under Program Files) fall back to LOCALAPPDATA.
    #>
    param([Parameter(Mandatory = $true)][string]$AppRoot)

    $dataRoot = $AppRoot
    if (-not (Test-DirectoryWritable (Join-Path $AppRoot 'logs'))) {
        $dataRoot = Join-Path $env:LOCALAPPDATA 'ExcelQueryTrigger'
        New-Item -ItemType Directory -Path $dataRoot -Force -ErrorAction SilentlyContinue | Out-Null
    }

    $paths = @{
        AppRoot      = $AppRoot
        DataRoot     = $dataRoot
        ConfigDir    = (Join-Path $dataRoot 'config')
        LogDir       = (Join-Path $dataRoot 'logs')
        ConfigPath   = (Join-Path (Join-Path $dataRoot 'config') 'rules.json')
        StatePath    = (Join-Path (Join-Path $dataRoot 'config') 'state.json')
        HistoryPath  = (Join-Path (Join-Path $dataRoot 'logs') 'history.jsonl')
        LauncherPath = (Join-Path $AppRoot 'Start-Hidden.vbs')
        # Registered in HKCU\...\Run. It passes -StartedFromLogon, which is how
        # the application tells a logon start apart from a manual one.
        LogonLauncherPath = (Join-Path $AppRoot 'Start-AtLogon.vbs')
    }

    foreach ($dir in @($paths.ConfigDir, $paths.LogDir)) {
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    return $paths
}

# ------------------------------------------------------------------------------
# Region: configuration defaults (single source of truth)
# ------------------------------------------------------------------------------

function Get-DefaultAppSettings {
    return @{
        startWithWindows           = $false
        startMinimized             = $true
        minimizeToTray             = $true
        showSuccessNotifications   = $true
        showErrorNotifications     = $true
        showTriggerNotifications   = $false
        showTriggeredResultPopup    = $true  # detailed result dialog after automatic/scheduled/logon trigger jobs
        watcherHealthCheckSeconds  = 60
        defaultDebounceSeconds     = 5
        defaultCooldownSeconds     = 30
        defaultRefreshWarningSeconds = 300
        logRetentionDays           = 30
        logMaxTotalMegabytes       = 50
        debugLogging               = $false
        coalesceDuplicateWorkbooks = $true
        rulesPaneMode              = 'Compact' # Compact / ShowAll; toggled in the rule-group title
        checkForUpdatesAutomatically = $true # one quiet public GitHub release check per day
        lastUpdateCheckUtc         = ''   # quiet GitHub release check, persisted across launches
        checkDataSources           = $true   # resolve each connection's server before refreshing
        autoDismissExcelDialogs    = $true   # close dialogs a hidden Excel raises during a job
        excelDialogGraceSeconds    = 15      # how long such a dialog may stand before it is closed
        cancelEscalationSeconds    = 12      # how long Cancel Job waits before offering to end Excel
        enginePollMilliseconds     = 500
        uiRefreshMilliseconds      = 500
        scheduleCheckSeconds       = 30
        startupPromptDelaySeconds  = 30
        uiFontSize                 = 9
        uiFontName                 = ''    # blank = pick the best installed font
        allowWorkbookMacrosByDefault = $false # safe default; per-action opt-in is available
    }
}

function Get-DefaultTrigger {
    return @{
        type                      = 'FileCreated'
        path                      = ''
        filter                    = '*.csv'
        contains                  = ''
        exclude                   = ''
        debounceSeconds           = 5
        cooldownSeconds           = 30
        waitForReady              = $true
        readyCheckIntervalSeconds = 2
        readyTimeoutSeconds       = 60
        # 0 disables the guard. A positive value asks before an automatic run
        # when the workbook records a refresh/save inside this many minutes.
        recentRefreshPromptMinutes = 0
        # Scheduled triggers. Ignored by every other type, and written to
        # rules.json for all of them so the schema stays uniform.
        scheduleTime              = '07:30'
        scheduleDays              = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
        # Logon triggers: Ask (default) or Automatic.
        logonBehavior             = 'Ask'
    }
}

function Get-DefaultAction {
    return @{
        type                   = 'ExcelRefresh'
        path                   = ''
        refreshMethod          = 'RefreshAll'
        selectedQueries        = @()
        save                   = $true
        close                  = $true
        visible                = $false
        timeoutSeconds         = 300
        continueOnError        = $true
        disableBackgroundQuery = $true
        allowWorkbookMacros     = $false
    }
}

function Get-TriggerTypeList {
    # Order matters: it is the order shown in the rule editor combo box.
    # File types first, then the ones that do not watch anything.
    return @('FileCreated', 'FileChangedSpecific', 'FileChangedAny', 'Scheduled', 'Logon', 'Manual')
}

function Get-TriggerTypeLabel {
    <#
        Plain-language names for the rule editor. -Short is used in the
        dashboard column, where the space is narrow.
    #>
    param([string]$Type, [switch]$Short)
    if ($Short) {
        switch ($Type) {
            'FileCreated'         { return 'New file' }
            'FileChangedSpecific' { return 'File updated' }
            'FileChangedAny'      { return 'File updated' }
            'Scheduled'           { return 'Scheduled' }
            'Logon'               { return 'At logon' }
            'Manual'              { return 'Manual' }
            default               { return $Type }
        }
    }
    switch ($Type) {
        'FileCreated'         { return 'New file added to folder' }
        'FileChangedSpecific' { return 'Specific file updated' }
        'FileChangedAny'      { return 'Matching file updated' }
        'Scheduled'           { return 'Scheduled time' }
        'Logon'               { return 'At Windows logon' }
        'Manual'              { return 'Manual only' }
        default               { return $Type }
    }
}

function Test-TriggerUsesFolder {
    param([string]$Type)
    return ($Type -eq 'FileCreated' -or $Type -eq 'FileChangedAny')
}

function Test-TriggerUsesWatcher {
    <#
        The single switch that decides whether a rule needs a
        FileSystemWatcher. Scheduled, Logon and Manual rules never create one,
        so adding them cost the monitoring layer nothing.
    #>
    param([string]$Type)
    return ($Type -eq 'FileCreated' -or $Type -eq 'FileChangedSpecific' -or $Type -eq 'FileChangedAny')
}

function Test-TriggerIsScheduled {
    param([string]$Type)
    return ($Type -eq 'Scheduled')
}

function Test-TriggerIsLogon {
    param([string]$Type)
    return ($Type -eq 'Logon')
}

function Get-WeekDayNameList {
    return @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
}

function ConvertTo-ScheduleTime {
    <#
        Parses "HH:mm" (24-hour) into @{ Hour; Minute } or $null. Deliberately
        strict and culture independent so a rule means the same thing on a
        machine with different regional settings.
    #>
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $match = [regex]::Match($Text.Trim(), '^([0-1]?[0-9]|2[0-3]):([0-5][0-9])$')
    if (-not $match.Success) { return $null }
    return @{ Hour = [int]$match.Groups[1].Value; Minute = [int]$match.Groups[2].Value }
}

function Format-ScheduleDays {
    <#  "Every day" / "Mon-Fri" / "Mon, Wed, Fri" for the dashboard column.  #>
    param([string[]]$Days)

    $selected = @($Days | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($selected.Count -eq 0)  { return 'no days selected' }
    if ($selected.Count -eq 7)  { return 'every day' }

    $weekdays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
    $isWeekdays = ($selected.Count -eq 5)
    if ($isWeekdays) {
        foreach ($day in $weekdays) { if ($selected -notcontains $day) { $isWeekdays = $false; break } }
    }
    if ($isWeekdays) { return 'Mon-Fri' }

    $short = @()
    foreach ($day in (Get-WeekDayNameList)) {
        if ($selected -contains $day) { $short += $day.Substring(0, 3) }
    }
    return ($short -join ', ')
}

function Get-TriggerTargetDescription {
    param([hashtable]$Rule)
    $trigger = $Rule.trigger
    if ($null -eq $trigger) { return '--' }
    if ($trigger.type -eq 'Manual') { return '--' }
    if (Test-TriggerIsScheduled $trigger.type) {
        return ('{0}  {1}' -f [string]$trigger.scheduleTime, (Format-ScheduleDays @($trigger.scheduleDays)))
    }
    if (Test-TriggerIsLogon $trigger.type) {
        if ([string]$trigger.logonBehavior -eq 'Automatic') { return 'refresh automatically' }
        return 'ask before refreshing'
    }
    if (Test-TriggerUsesFolder $trigger.type) {
        $text = $trigger.filter
        if ([string]::IsNullOrWhiteSpace($text)) { $text = '*.*' }
        if (-not [string]::IsNullOrWhiteSpace($trigger.contains)) { $text += (' [{0}]' -f $trigger.contains) }
        return $text
    }
    if ([string]::IsNullOrWhiteSpace($trigger.path)) { return '--' }
    return (Split-Path -Leaf $trigger.path)
}

# ------------------------------------------------------------------------------
# Region: shared state (the only object both runspaces touch)
# ------------------------------------------------------------------------------

function New-SharedState {
    <#
        Thread-safe hand-off between the UI runspace and the Engine runspace.
        Rule: the engine WRITES status / activity / history, the UI READS them.
        The UI WRITES commands, the engine READS them. Nothing else is shared.
    #>
    $shared = [hashtable]::Synchronized(@{})

    $shared.Status          = 'Starting'   # Starting | Running | Waiting for network | Paused | Degraded | Error | Stopping | Stopped
    $shared.StatusDetail    = ''
    # Written by the engine's network monitor, read by the UI.
    # Online | NetworkOffline | CheckingPaths | WaitingForPaths
    $shared.NetworkState       = 'Online'
    $shared.NetworkStateDetail = ''
    $shared.EngineAlive     = $false
    $shared.ShouldExit      = $false
    $shared.Paused          = $false
    $shared.ReloadRequested = $false
    $shared.ConfigVersion   = 0
    $shared.CancelCurrentJob = $false
    $shared.OwnedExcelPid   = 0
    $shared.OwnedExcelStartedAtUtc = $null
    $shared.OwnedExcelWindowHandle = 0
    # Cancellation is cooperative, so it cannot reach Excel while Excel is stuck
    # inside a COM call. These two let the UI escalate when that happens.
    $shared.CancelRequestedAt       = $null   # UI  -> when Cancel Job was confirmed
    $shared.ForcedExcelTermination  = $false  # UI  -> the dedicated Excel was ended
    $shared.FatalError      = ''
    $shared.LogSeq          = 0
    $shared.DebugLogging    = $false
    # Startup progress is written by the engine and read by the splash screen.
    # It is deliberately descriptive rather than pretending every step takes the
    # same amount of time.
    $shared.StartupMessage          = 'Starting the monitoring engine...'
    $shared.StartupCurrent          = 0
    $shared.StartupTotal            = 0
    # Startup uses two explicit gates. The engine first proves that every
    # enabled FileSystemWatcher is armed, then the UI finishes the initial
    # workbook-information scan. Only after both are true does StartupReady
    # become true and Status change to Running.
    $shared.StartupMonitoringReady  = $false
    $shared.StartupUiReady          = $false
    $shared.StartupReady            = $false

    $shared.CurrentJob = [hashtable]::Synchronized(@{
        Active       = $false
        # Whether the Excel instance is on screen. A hidden instance must never
        # be left sitting on a dialog, because nobody can answer it.
        ExcelVisible = $false
        # Cancellation can be requested only at cooperative checkpoints. Forced
        # process termination is a narrower permission: it is never safe while
        # Excel may be writing or closing the workbook.
        CanCancel         = $false
        CanForceTerminate = $false
        # Wider than CanForceTerminate on purpose: during a direct save Excel
        # must never be killed, but a modal dialog raised there still has to be
        # detected and closed, or the job waits forever for an answer.
        CanInspectDialogs = $false
        JobId     = ''
        RuleId    = ''
        RuleName  = ''
        Workbook  = ''
        Stage     = 'Idle'
        StartedAt = $null
        Source    = ''
        # Query-level progress is observational. Excel can evaluate Power Query
        # dependencies in parallel and does not expose every internal query, so
        # these fields report only the connections Excel makes observable.
        QueryTotal          = 0
        QueryCompleted      = 0
        QueryPosition       = 0
        QueryName           = ''
        QueryElapsedSeconds = 0
        QueryProgressDetail = ''
        SaveStartedAt       = $null
        SaveMode            = ''
        SaveProgressDetail  = ''
    })
    $shared.JobSequence = 0

    $shared.Commands      = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    $shared.Notifications = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    # Automatic rules may optionally require a user decision before they enter
    # the refresh queue. The engine posts lightweight requests here; the UI owns
    # the dialog and replies through the normal command queue.
    $shared.RefreshApprovalRequests = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    $shared.PendingApprovalRuleIds  = [hashtable]::Synchronized(@{})
    # A completed job posts its workbook paths here. The Dashboard drains the
    # queue and immediately re-reads saved/refresh metadata instead of waiting
    # for the normal periodic workbook-information scan.
    $shared.WorkbookInfoRefreshRequests = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    $shared.Activity      = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    $shared.PendingJobs   = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    $shared.History       = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    $shared.RuleState     = [hashtable]::Synchronized(@{})

    return $shared
}

# ==============================================================================
#  The pending job queue
#
#  One synchronized list, written from both runspaces. The engine appends jobs
#  raised by triggers and takes the head when it is free; the dashboard appends
#  manual runs and reorders or removes whatever is still waiting. That is what
#  makes a Run Now pressed during a refresh appear straight away: the engine is
#  blocked inside Excel and will not read its command queue until the current
#  job ends, but the list itself is always writable.
#
#  Every operation that reads and then writes takes the list's own lock, so a
#  reorder can never race the engine taking the head. Moves and removals are
#  addressed by job id rather than by position for the same reason.
# ==============================================================================

function New-PendingJobId {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        $Shared.JobSequence = [int]$Shared.JobSequence + 1
        return ('JOB-{0:0000}' -f [int]$Shared.JobSequence)
    }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function New-PendingJob {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [string]$TriggerSource = 'File',
        [string]$TriggerFile = '',
        [string[]]$TriggerFiles = @()
    )

    $sourceFiles = New-Object System.Collections.ArrayList
    foreach ($sourceFile in @($TriggerFiles) + @($TriggerFile)) {
        $value = [string]$sourceFile
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if (@($sourceFiles) -notcontains $value) { [void]$sourceFiles.Add($value) }
    }
    $primaryTriggerFile = $(if ($sourceFiles.Count -gt 0) { [string]$sourceFiles[0] } else { '' })

    return @{
        Id            = (New-PendingJobId -Shared $Shared)
        RuleId        = [string]$Rule.id
        RuleName      = [string]$Rule.name
        TriggerType   = [string]$Rule.trigger.type
        TriggerSource = $TriggerSource
        TriggerFile   = $primaryTriggerFile
        TriggerFiles  = @($sourceFiles.ToArray())
        TriggeredAt   = (Get-Date)
        Trigger       = $Rule.trigger
        Actions       = @($Rule.actions)
        Status        = 'Pending'
        Stage         = 'Pending'
        StartedAt     = $null
        FinishedAt    = $null
        Results       = (New-Object System.Collections.ArrayList)
    }
}

function Get-PendingJobWorkbookPaths {
    param([Parameter(Mandatory = $true)][hashtable]$Job)
    $paths = New-Object System.Collections.ArrayList
    foreach ($action in @($Job.Actions)) { [void]$paths.Add([string]$action.path) }
    $result = @($paths.ToArray())
    return ,$result
}

function Copy-FileForReading {
    <#
        Creates a read-only snapshot for workbook metadata/diagnostic readers.
        Kept in Common because both readers need the same file-system Adapter;
        workbook metadata must not depend on the diagnostics UI load order.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('eqt-read-{0}.zip' -f ([guid]::NewGuid().ToString('N')))
    $source = $null
    $target = $null
    try {
        # FileShare::Delete matters as much as ReadWrite here. Excel saves an
        # existing workbook by writing a temporary file and then replacing the
        # original, which needs delete/rename rights. A reader that omits Delete
        # makes that final step fail with "Someone else is working in ...".
        $source = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $target = [System.IO.File]::Create($temp)
        $source.CopyTo($target)
        return $temp
    }
    catch { return '' }
    finally {
        if ($null -ne $target) { try { $target.Dispose() } catch { } }
        if ($null -ne $source) { try { $source.Dispose() } catch { } }
    }
}

function ConvertTo-JobComparisonPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $value = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if ([System.IO.Path]::DirectorySeparatorChar -eq [char]'\') { $value = $value.Replace('/', '\') }
    else { $value = $value.Replace('\', '/') }
    try { $value = [System.IO.Path]::GetFullPath($value) } catch { }
    $value = $value.ToLowerInvariant()
    if ($value.Length -gt 3) { $value = $value.TrimEnd([char[]]@([char]'\', [char]'/')) }
    return $value
}

function Get-RefreshActionIdentity {
    <#
        A workbook path alone does not identify an operation. Two requests for
        different query sets or macro policies must both remain in the queue.
    #>
    param([Parameter(Mandatory = $true)]$Action)

    $queries = @(@($Action.selectedQueries) | ForEach-Object {
        ([string]$_).Trim().ToLowerInvariant()
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    return (@(
        (ConvertTo-JobComparisonPath ([string]$Action.path)),
        ([string]$Action.type).ToLowerInvariant(),
        ([string]$Action.refreshMethod).ToLowerInvariant(),
        ($queries -join ','),
        [string](ConvertTo-BoolValue $Action.allowWorkbookMacros $false),
        [string](ConvertTo-BoolValue $Action.disableBackgroundQuery $true),
        [string](ConvertTo-BoolValue $Action.continueOnError $true),
        [string](ConvertTo-IntValue $Action.timeoutSeconds 300 5)
    ) -join '|')
}

function Get-RefreshJobActionIdentity {
    param([Parameter(Mandatory = $true)][hashtable]$Job)
    return (@(@($Job.Actions) | ForEach-Object { Get-RefreshActionIdentity $_ }) -join "`n")
}

function Add-PendingJob {
    <#
        Applies the de-duplication rules before queueing:
          * the same rule already waiting  -> drop the new job entirely
          * the same workbook already waiting in another pending job ->
            drop just that action (optional, controlled by appSettings)
        Returns $true when something was queued.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$Job,
        [bool]$CoalesceDuplicateWorkbooks = $true
    )

    $queue = $Shared.PendingJobs
    $queuedCount = 0
    $dropMessages = New-Object System.Collections.ArrayList

    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        foreach ($pending in @($queue.ToArray())) {
            if ([string]$pending.RuleId -eq [string]$Job.RuleId -and
                (Get-RefreshJobActionIdentity -Job $pending) -eq (Get-RefreshJobActionIdentity -Job $Job)) {
                $mergedFiles = New-Object System.Collections.ArrayList
                foreach ($file in @($pending.TriggerFiles) + @($pending.TriggerFile) + @($Job.TriggerFiles) + @($Job.TriggerFile)) {
                    $value = [string]$file
                    if (-not [string]::IsNullOrWhiteSpace($value) -and @($mergedFiles) -notcontains $value) {
                        [void]$mergedFiles.Add($value)
                    }
                }
                $pending.TriggerFiles = @($mergedFiles.ToArray())
                if ($pending.TriggerFiles.Count -gt 0) { $pending.TriggerFile = [string]$pending.TriggerFiles[0] }
                [void]$dropMessages.Add(@{ Workbook = ''; Text = ('An identical job for this rule is already queued ({0}); its source-file set was updated.' -f $pending.Id) })
                $Job.Actions = @()
                break
            }
        }

        if (@($Job.Actions).Count -gt 0 -and $CoalesceDuplicateWorkbooks) {
            $queuedActions = @{}
            foreach ($pending in @($queue.ToArray())) {
                foreach ($pendingAction in @($pending.Actions)) {
                    $queuedActions[(Get-RefreshActionIdentity $pendingAction)] = $pending.Id
                }
            }
            if ($queuedActions.Count -gt 0) {
                $kept = New-Object System.Collections.ArrayList
                foreach ($action in @($Job.Actions)) {
                    $key = Get-RefreshActionIdentity $action
                    if ($queuedActions.ContainsKey($key)) {
                        [void]$dropMessages.Add(@{ Workbook = [string]$action.path; Text = ('An identical refresh operation is already queued by {0}; this duplicate was skipped.' -f $queuedActions[$key]) })
                    }
                    else { [void]$kept.Add($action) }
                }
                $Job.Actions = @($kept.ToArray())
            }
        }

        if (@($Job.Actions).Count -gt 0) {
            [void]$queue.Add($Job)
            $queuedCount = $queue.Count
        }
    }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }

    # Logging happens outside the lock: it writes to disk and must never be the
    # thing that keeps the engine waiting to take the next job.
    foreach ($drop in @($dropMessages)) {
        Write-AppLog -Level 'INFO' -RuleName $Job.RuleName -Workbook ([string]$drop.Workbook) -Message ([string]$drop.Text)
    }
    if ($queuedCount -eq 0) {
        if (@($dropMessages).Count -eq 0) {
            Write-AppLog -Level 'INFO' -RuleName $Job.RuleName -Message 'Nothing left to do after de-duplication; no job was queued.'
        }
        return $false
    }

    Write-AppLog -Level 'INFO' -RuleName $Job.RuleName `
        -Message ('Queued {0} with {1} workbook action(s). Queue length: {2}.' -f $Job.Id, @($Job.Actions).Count, $queuedCount)
    return $true
}

function Get-NextPendingJob {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        if ($queue.Count -eq 0) { return $null }
        $job = $queue[0]
        $queue.RemoveAt(0)
        return $job
    }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function Get-PendingJobCount {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    return $Shared.PendingJobs.Count
}

function Clear-PendingJobs {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try { $queue.Clear() }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function Get-PendingJobSnapshot {
    <#  A stable copy for the dashboard to draw from.  #>
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        # No , prefix here: every caller wraps this in @(), which would turn a
        # deliberately unrolled array back into a single nested element.
        $copy = @($queue.ToArray())
        return $copy
    }
    catch { return @() }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function Move-PendingJob {
    <#
        Moves one queued job by Delta positions. Addressed by id, so the move
        still lands on the right job if the engine took the head in between.
        Returns $true when the queue changed.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][string]$JobId,
        [Parameter(Mandatory = $true)][int]$Delta
    )
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        $index = -1
        for ($i = 0; $i -lt $queue.Count; $i++) {
            if ([string]$queue[$i].Id -eq $JobId) { $index = $i; break }
        }
        if ($index -lt 0) { return $false }
        $target = $index + $Delta
        if ($target -lt 0 -or $target -ge $queue.Count) { return $false }
        $job = $queue[$index]
        $queue.RemoveAt($index)
        $queue.Insert($target, $job)
        return $true
    }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function Remove-PendingJob {
    <#  Returns the removed job, or $null when it had already started.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][string]$JobId
    )
    $queue = $Shared.PendingJobs
    [System.Threading.Monitor]::Enter($queue.SyncRoot)
    try {
        for ($i = 0; $i -lt $queue.Count; $i++) {
            if ([string]$queue[$i].Id -eq $JobId) {
                $job = $queue[$i]
                $queue.RemoveAt($i)
                return $job
            }
        }
        return $null
    }
    finally { [System.Threading.Monitor]::Exit($queue.SyncRoot) }
}

function New-ManualWorkbookRule {
    <#
        Wraps a one-off workbook refresh in an ad-hoc rule, so that queueing,
        de-duplication, cancellation, logging and history all behave exactly as
        they do for a real rule. Shared so the engine and the dashboard build
        the identical thing.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        $Action = $null,
        [int]$DefaultTimeoutSeconds = 300,
        [bool]$AllowWorkbookMacros = $false
    )

    if ($null -ne $Action) {
        $manualAction = Merge-DefaultValues -Default (Get-DefaultAction) -Value (ConvertTo-HashtableDeep $Action)
        $manualAction['path'] = $Path
    }
    else {
        $manualAction = Get-DefaultAction
        $manualAction['path'] = $Path
        $manualAction['timeoutSeconds'] = $DefaultTimeoutSeconds
        $manualAction['allowWorkbookMacros'] = $AllowWorkbookMacros
    }

    return (ConvertTo-NormalizedRule @{
        id      = 'MANUAL-REFRESH'
        name    = ('Manual refresh: {0}' -f (Split-Path -Leaf $Path))
        enabled = $true
        trigger = @{ type = 'Manual'; waitForReady = $false }
        actions = @($manualAction)
    })
}

function Get-WorkbookConnectionSummary {
    <#
        Every workbook connection, with what can be learned from it without
        touching anything: its name, its connection string, and the Power Query
        it belongs to if it is one.

        $Release is the COM release helper to use, because the engine and the
        dashboard each have their own.
    #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [scriptblock]$Release = $null
    )

    $found = New-Object System.Collections.ArrayList
    $connections = $null
    try { $connections = $Workbook.Connections } catch { $connections = $null }
    if ($null -eq $connections) { return @() }

    try {
        for ($i = 1; $i -le $connections.Count; $i++) {
            $connection = $null
            $inner = $null
            try {
                $connection = $connections.Item($i)
                $name = ''
                try { $name = [string]$connection.Name } catch { $name = '' }
                if ([string]::IsNullOrWhiteSpace($name)) { continue }

                $type = 0
                try { $type = [int]$connection.Type } catch { $type = 0 }
                if ($type -eq 1) { try { $inner = $connection.OLEDBConnection } catch { $inner = $null } }
                elseif ($type -eq 2) { try { $inner = $connection.ODBCConnection } catch { $inner = $null } }

                $connectionString = ''
                if ($null -ne $inner) { try { $connectionString = [string]$inner.Connection } catch { $connectionString = '' } }

                [void]$found.Add(@{
                    Name             = $name
                    ConnectionString = $connectionString
                    QueryLocation    = (Get-MashupQueryLocation -ConnectionString $connectionString)
                    SheetTable       = (Get-WorksheetConnectionTarget -Name $name)
                })
            }
            catch { }
            finally { if ($null -ne $Release) { & $Release $inner $connection } }
        }
    }
    catch { }
    finally { if ($null -ne $Release) { & $Release $connections } }

    return @($found.ToArray())
}

function Get-MashupQueryLocation {
    <#
        The query a Power Query workbook connection belongs to.

        A connection string of the form
            Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Orderdata
        names its query in Location=. That is the only locale-independent link
        between Workbook.Queries and Workbook.Connections: the connection's own
        name is localised - "Query - Orderdata" in English, "クエリ - Orderdata"
        in Japanese - so matching on an English prefix silently fails on every
        other language build of Excel.

        Returns '' when the connection is not a Power Query one.
    #>
    param([string]$ConnectionString)

    if ([string]::IsNullOrWhiteSpace($ConnectionString)) { return '' }
    if ($ConnectionString -notmatch '(?i)Microsoft\.Mashup') { return '' }
    $match = [regex]::Match($ConnectionString, '(?i)(?:^|;)\s*Location\s*=\s*([^;]+)')
    if (-not $match.Success) { return '' }
    return $match.Groups[1].Value.Trim().Trim('"').Trim()
}

function Get-WorksheetConnectionTarget {
    <#
        Excel names the connection it creates when a query is loaded onto a
        sheet "WorksheetConnection_<file>!<table>". The part after the "!" is
        the loaded table, which is the proof that the query landed on a sheet.
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    if ($Name -notmatch '^(?i)WorksheetConnection_') { return '' }
    $bang = $Name.LastIndexOf('!')
    if ($bang -lt 0 -or $bang -ge ($Name.Length - 1)) { return '' }
    return $Name.Substring($bang + 1).Trim()
}

function Get-ConnectionServerTarget {
    <#
        Pulls a real network server out of an OLEDB/ODBC connection string.

        Returns $null whenever the string does not name one, and that is the
        common case: a Power Query connection reads
        "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=..."
        because the actual source lives inside the M formula, not here. Only
        values that positively look like a host name or an IP address are
        returned - excluding the shapes we happen to think of is not safe
        enough, because anything missed becomes a workbook that will not
        refresh.
    #>
    param([string]$ConnectionString)

    if ([string]::IsNullOrWhiteSpace($ConnectionString)) { return $null }

    # The Mashup provider never names the real server, so there is nothing here
    # that could be checked.
    if ($ConnectionString -match '(?i)Microsoft\.Mashup') { return $null }

    $value = $null
    foreach ($key in @('Data Source', 'Server', 'Address', 'Addr', 'Network Address')) {
        $pattern = '(?i)(?:^|;)\s*' + [regex]::Escape($key) + '\s*=\s*([^;]+)'
        $match = [regex]::Match($ConnectionString, $pattern)
        if ($match.Success) { $value = $match.Groups[1].Value.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($value)) { break }
    }
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }

    # Placeholders such as $Workbook$ and $EmbeddedMashup(..)$ are not servers.
    if ($value.Contains('$') -or $value.Contains('%')) { return $null }

    $serverHost = $value

    # Optional protocol prefix used by SQL Server: tcp:, np:, lpc:, admin:
    if ($serverHost -match '^(?i)(tcp|np|lpc|admin):(?<rest>.+)$') { $serverHost = $Matches['rest'].Trim() }

    # host,port
    $port = 0
    if ($serverHost -match '^(?<h>[^,]+),\s*(?<p>\d+)\s*$') {
        $serverHost = $Matches['h'].Trim()
        $port = [int]$Matches['p']
    }

    # host\instance
    if ($serverHost.Contains('\')) { $serverHost = $serverHost.Split('\')[0].Trim() }
    if ([string]::IsNullOrWhiteSpace($serverHost)) { return $null }

    # Bracketed IPv6 is a valid target; anything else with a colon is not a
    # plain host name and is left alone.
    if ($serverHost.StartsWith('[') -and $serverHost.EndsWith(']')) {
        $serverHost = $serverHost.Substring(1, $serverHost.Length - 2)
    }

    $parsedAddress = $null
    $isIpAddress = [System.Net.IPAddress]::TryParse($serverHost, [ref]$parsedAddress)
    if (-not $isIpAddress) {
        # A DNS label: letters, digits, dots and hyphens, starting and ending
        # with a letter or digit. Everything else - paths, URLs, quoted values,
        # anything containing a space, an equals sign or a brace - is not
        # something this check is able to judge.
        if ($serverHost.Length -gt 253) { return $null }
        if ($serverHost -notmatch '^[A-Za-z0-9]([A-Za-z0-9._-]*[A-Za-z0-9])?$') { return $null }
        # A bare word ending in a known file extension is a file, not a host.
        if ($serverHost -match '(?i)\.(xlsx?|xlsb|xlsm|csv|txt|accdb|mdb|cub|odc|dbf|xml|json)$') { return $null }
    }

    if ($serverHost -match '^(?i)(localhost|localhost\.localdomain)$') { return $null }
    if ($isIpAddress -and ([System.Net.IPAddress]::IsLoopback($parsedAddress))) { return $null }

    # Only used when the connection string did not name a port, and only ever to
    # raise a warning - a named SSAS or SQL instance does not listen here.
    $defaultPort = 0
    if     ($ConnectionString -match '(?i)MSOLAP')                          { $defaultPort = 2383 }
    elseif ($ConnectionString -match '(?i)(SQLOLEDB|MSOLEDBSQL|SQL Server)'){ $defaultPort = 1433 }

    return @{ Host = $serverHost; Port = $port; DefaultPort = $defaultPort; Raw = $value }
}

function Test-ServerReachable {
    <#
        Resolves the name, then optionally probes a port.

        Only "no such host", seen twice, is treated as conclusive. A DNS server
        that is down, a VPN that has not come up yet or any other transient
        resolver failure must never make every workbook look broken, so those
        return Conclusive = $false and the caller lets the refresh proceed.
    #>
    param([string]$ServerHost, [int]$Port = 0, [int]$DefaultPort = 0, [int]$TimeoutMs = 2000)

    $outcome = @{ Resolved = $false; Conclusive = $false; Probed = $false; PortOpen = $false }

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            $addresses = @([System.Net.Dns]::GetHostAddresses($ServerHost))
            if ($addresses.Count -gt 0) { $outcome.Resolved = $true; $outcome.Conclusive = $true; break }
        }
        catch {
            $inner = $_.Exception
            while ($null -ne $inner -and -not ($inner -is [System.Net.Sockets.SocketException])) { $inner = $inner.InnerException }
            if ($null -ne $inner -and $inner.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound) {
                $outcome.Conclusive = $true
            }
            else {
                # Not a "this name does not exist" answer - stop and say nothing.
                $outcome.Conclusive = $false
                return $outcome
            }
        }
        if ($attempt -eq 1) { Start-Sleep -Milliseconds 300 }
    }

    if (-not $outcome.Resolved) { return $outcome }

    $probePort = $(if ($Port -gt 0) { $Port } else { $DefaultPort })
    if ($probePort -le 0) { return $outcome }
    $outcome.Probed = $true

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ServerHost, $probePort, $null, $null)
        if ($async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            try { $client.EndConnect($async); $outcome.PortOpen = [bool]$client.Connected } catch { $outcome.PortOpen = $false }
        }
    }
    catch { $outcome.PortOpen = $false }
    finally { if ($null -ne $client) { try { $client.Close() } catch { } } }

    return $outcome
}

function Send-EngineCommand {
    <#  Called from the UI thread only. Never blocks.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][string]$Type,
        [hashtable]$Payload
    )
    $command = @{ Type = $Type; SentAt = (Get-Date) }
    if ($Payload) {
        foreach ($key in $Payload.Keys) { $command[$key] = $Payload[$key] }
    }
    $Shared.Commands.Enqueue($command)
}

function Request-Notification {
    <#  Engine -> UI balloon request. The UI decides whether to show it.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [string]$Title,
        [string]$Text,
        [ValidateSet('Info', 'Warning', 'Error')][string]$Level = 'Info'
    )
    $Shared.Notifications.Enqueue(@{ Title = $Title; Text = $Text; Level = $Level })
}

function Set-RuleState {
    <#
        Per-rule runtime facts shown in the dashboard (last trigger, last result,
        watcher health). Kept out of rules.json because it is state, not config.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][string]$RuleId,
        [hashtable]$Values
    )

    if (-not $Shared.RuleState.ContainsKey($RuleId)) {
        $Shared.RuleState[$RuleId] = [hashtable]::Synchronized(@{
            WatcherStatus  = 'Unknown'
            LastTrigger    = $null
            LastSuccess    = $null
            LastError      = $null
            LastResultText = ''
            LastRun        = $null
            LastDuration   = ''
            LastRunStatus  = ''
            LastScheduledRun = ''   # "yyyy-MM-dd HH:mm" of the last schedule honoured
        })
    }

    if ($Values) {
        $entry = $Shared.RuleState[$RuleId]
        foreach ($key in $Values.Keys) { $entry[$key] = $Values[$key] }
    }

    return $Shared.RuleState[$RuleId]
}
