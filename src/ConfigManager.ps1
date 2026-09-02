# ==============================================================================
#  ConfigManager.ps1
#  rules.json  -> normalised hashtable graph (defaults applied, extras kept)
#  state.json  -> last known status per rule, restored on the dashboard
#  Files are written as BOM-less UTF-8 and read with File::ReadAllText so that
#  Japanese filename conditions survive a round trip on Windows PowerShell 5.1.
# ==============================================================================

Set-StrictMode -Version 1.0

function ConvertTo-NormalizedRule {
    param($Rule)

    $source = $Rule
    if ($source -isnot [hashtable]) { $source = ConvertTo-HashtableDeep $source }
    if ($null -eq $source) { $source = @{} }

    $normalized = @{}
    foreach ($key in $source.Keys) { $normalized[$key] = $source[$key] }

    if ([string]::IsNullOrWhiteSpace([string]$normalized['id'])) { $normalized['id'] = New-ShortId -Prefix 'RULE' }
    # The name is deliberately NOT defaulted here: the rule editor validates it,
    # and silently renaming an empty name would hide the mistake from the user.
    $normalized['name'] = ConvertTo-StringValue $normalized['name']
    $normalized['enabled'] = ConvertTo-BoolValue $normalized['enabled'] $true
    # Per-rule confirmation before automatic file/schedule refreshes. Missing in
    # older configs => false, preserving the historical unattended behaviour.
    $normalized['askBeforeRefresh'] = ConvertTo-BoolValue $normalized['askBeforeRefresh'] $false

    $normalized['trigger'] = Merge-DefaultValues -Default (Get-DefaultTrigger) -Value $normalized['trigger']
    $trigger = $normalized['trigger']
    if ((Get-TriggerTypeList) -notcontains $trigger['type']) { $trigger['type'] = 'FileCreated' }
    $trigger['debounceSeconds']           = ConvertTo-IntValue $trigger['debounceSeconds'] 5 0
    $trigger['cooldownSeconds']           = ConvertTo-IntValue $trigger['cooldownSeconds'] 30 0
    $trigger['readyCheckIntervalSeconds'] = ConvertTo-IntValue $trigger['readyCheckIntervalSeconds'] 2 1
    $trigger['readyTimeoutSeconds']       = ConvertTo-IntValue $trigger['readyTimeoutSeconds'] 60 1
    $trigger['recentRefreshPromptMinutes'] = ConvertTo-IntValue $trigger['recentRefreshPromptMinutes'] 0 0
    $trigger['waitForReady']              = ConvertTo-BoolValue $trigger['waitForReady'] $true
    $trigger['path']                      = ConvertTo-StringValue $trigger['path']
    $trigger['filter']                    = ConvertTo-StringValue $trigger['filter']
    $trigger['contains']                  = ConvertTo-StringValue $trigger['contains']
    $trigger['exclude']                   = ConvertTo-StringValue $trigger['exclude']

    # Scheduled triggers. An unparsable time is kept as typed so the editor can
    # show the user what is wrong; validation - not normalisation - rejects it.
    $trigger['scheduleTime'] = ConvertTo-StringValue $trigger['scheduleTime']
    if ([string]::IsNullOrWhiteSpace($trigger['scheduleTime'])) { $trigger['scheduleTime'] = '07:30' }

    $validDays = Get-WeekDayNameList
    $days = New-Object System.Collections.ArrayList
    foreach ($day in @($trigger['scheduleDays'])) {
        if ($null -eq $day) { continue }
        $name = ConvertTo-StringValue $day
        foreach ($valid in $validDays) {
            if ($name -eq $valid -or $name -eq $valid.Substring(0, 3)) {
                if ($days -notcontains $valid) { [void]$days.Add($valid) }
                break
            }
        }
    }
    # Keep them in week order regardless of how they were written.
    $ordered = New-Object System.Collections.ArrayList
    foreach ($valid in $validDays) { if ($days -contains $valid) { [void]$ordered.Add($valid) } }
    $trigger['scheduleDays'] = @($ordered.ToArray())

    # Logon triggers. Anything unrecognised becomes Ask, never Automatic - an
    # unattended refresh must not be something a typo can switch on.
    if ((ConvertTo-StringValue $trigger['logonBehavior']) -eq 'Automatic') { $trigger['logonBehavior'] = 'Automatic' }
    else { $trigger['logonBehavior'] = 'Ask' }

    $actions = New-Object System.Collections.ArrayList
    if ($null -ne $normalized['actions']) {
        foreach ($rawAction in @($normalized['actions'])) {
            $action = Merge-DefaultValues -Default (Get-DefaultAction) -Value $rawAction
            $action['type']                   = ConvertTo-StringValue $action['type']
            if ([string]::IsNullOrWhiteSpace($action['type'])) { $action['type'] = 'ExcelRefresh' }
            $action['path']                   = ConvertTo-StringValue $action['path']
            $refreshMethod = ConvertTo-StringValue $action['refreshMethod']
            if ($refreshMethod -eq 'SelectedQueries') { $action['refreshMethod'] = 'SelectedQueries' }
            else { $action['refreshMethod'] = 'RefreshAll' }
            $selectedQueries = New-Object System.Collections.ArrayList
            foreach ($queryName in @($action['selectedQueries'])) {
                $cleanQueryName = ConvertTo-StringValue $queryName
                if (-not [string]::IsNullOrWhiteSpace($cleanQueryName) -and -not (@($selectedQueries) -contains $cleanQueryName)) { [void]$selectedQueries.Add($cleanQueryName) }
            }
            $action['selectedQueries']        = @($selectedQueries.ToArray())
            # These are product safety invariants, not user-configurable options.
            # Enforce them for legacy and hand-edited JSON as well as the UI.
            $action['save']                   = $true
            $action['close']                  = $true
            $action['visible']                = $false
            $action['continueOnError']        = ConvertTo-BoolValue $action['continueOnError'] $true
            $action['disableBackgroundQuery'] = ConvertTo-BoolValue $action['disableBackgroundQuery'] $true
            $action['allowWorkbookMacros']     = ConvertTo-BoolValue $action['allowWorkbookMacros'] $false
            $action['timeoutSeconds']         = ConvertTo-IntValue $action['timeoutSeconds'] 300 5
            [void]$actions.Add($action)
        }
    }
    $normalized['actions'] = @($actions.ToArray())

    return $normalized
}

function ConvertTo-RuleComparisonPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $value = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if ([System.IO.Path]::DirectorySeparatorChar -eq [char]'\') { $value = $value.Replace('/', '\') }
    else { $value = $value.Replace('\', '/') }
    # Treat relative paths, redundant . segments, slash direction, case and a
    # trailing separator as the same workbook. GetFullPath is deliberately
    # best-effort because an unreachable network path must still be comparable.
    try { $value = [System.IO.Path]::GetFullPath($value) } catch { }
    $value = $value.ToLowerInvariant()
    if ($value.Length -gt 3) { $value = $value.TrimEnd([char[]]@([char]'\', [char]'/')) }
    return $value
}

function Find-DuplicateWorkbookAction {
    <# Returns the zero-based index of the same workbook, or -1. #>
    param(
        [object[]]$Actions = @(),
        [Parameter(Mandatory = $true)][string]$CandidatePath,
        [int]$ExcludeIndex = -1
    )

    $candidate = ConvertTo-RuleComparisonPath $CandidatePath
    if ([string]::IsNullOrWhiteSpace($candidate)) { return -1 }
    for ($i = 0; $i -lt @($Actions).Count; $i++) {
        if ($i -eq $ExcludeIndex) { continue }
        $existing = ConvertTo-RuleComparisonPath ([string]$Actions[$i].path)
        if ($existing -eq $candidate) { return $i }
    }
    return -1
}

function Get-RuleDefinitionSignature {
    <#
        Canonical rule definition used only for duplicate detection. Identity,
        display name, enabled state and runtime history are deliberately absent.
        Trigger fields that have no effect for the selected type are also
        absent, so two rules that behave identically compare identically.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    $normalized = ConvertTo-NormalizedRule $Rule
    $sourceTrigger = $normalized['trigger']
    $triggerType = [string]$sourceTrigger['type']
    $trigger = [ordered]@{ type = $triggerType }

    if (Test-TriggerUsesFolder $triggerType) {
        $trigger['path'] = ConvertTo-RuleComparisonPath ([string]$sourceTrigger['path'])
        $trigger['filter'] = ([string]$sourceTrigger['filter']).Trim().ToLowerInvariant()
        $trigger['contains'] = ([string]$sourceTrigger['contains']).Trim().ToLowerInvariant()
        $trigger['exclude'] = ([string]$sourceTrigger['exclude']).Trim().ToLowerInvariant()
    }
    elseif ($triggerType -eq 'FileChangedSpecific') {
        $trigger['path'] = ConvertTo-RuleComparisonPath ([string]$sourceTrigger['path'])
    }

    if (Test-TriggerUsesWatcher $triggerType) {
        $trigger['debounceSeconds'] = [int]$sourceTrigger['debounceSeconds']
        $trigger['cooldownSeconds'] = [int]$sourceTrigger['cooldownSeconds']
        $trigger['waitForReady'] = [bool]$sourceTrigger['waitForReady']
        $trigger['readyCheckIntervalSeconds'] = [int]$sourceTrigger['readyCheckIntervalSeconds']
        $trigger['readyTimeoutSeconds'] = [int]$sourceTrigger['readyTimeoutSeconds']
        $trigger['recentRefreshPromptMinutes'] = [int]$sourceTrigger['recentRefreshPromptMinutes']
    }
    elseif (Test-TriggerIsScheduled $triggerType) {
        $trigger['scheduleTime'] = [string]$sourceTrigger['scheduleTime']
        $trigger['scheduleDays'] = @($sourceTrigger['scheduleDays'])
        $trigger['recentRefreshPromptMinutes'] = [int]$sourceTrigger['recentRefreshPromptMinutes']
    }
    elseif (Test-TriggerIsLogon $triggerType) {
        $trigger['logonBehavior'] = [string]$sourceTrigger['logonBehavior']
        if ([string]$sourceTrigger['logonBehavior'] -eq 'Automatic') {
            $trigger['recentRefreshPromptMinutes'] = [int]$sourceTrigger['recentRefreshPromptMinutes']
        }
    }

    $actions = New-Object System.Collections.ArrayList
    foreach ($sourceAction in @($normalized['actions'])) {
        $queries = @()
        if ([string]$sourceAction['refreshMethod'] -eq 'SelectedQueries') {
            $queries = @(@($sourceAction['selectedQueries']) | ForEach-Object {
                ([string]$_).Trim().ToLowerInvariant()
            } | Sort-Object)
        }
        $action = [ordered]@{
            type                   = [string]$sourceAction['type']
            path                   = ConvertTo-RuleComparisonPath ([string]$sourceAction['path'])
            refreshMethod          = [string]$sourceAction['refreshMethod']
            selectedQueries        = $queries
            save                   = [bool]$sourceAction['save']
            close                  = [bool]$sourceAction['close']
            visible                = [bool]$sourceAction['visible']
            timeoutSeconds         = [int]$sourceAction['timeoutSeconds']
            continueOnError        = [bool]$sourceAction['continueOnError']
            disableBackgroundQuery = [bool]$sourceAction['disableBackgroundQuery']
            allowWorkbookMacros    = [bool]$sourceAction['allowWorkbookMacros']
        }
        [void]$actions.Add($action)
    }

    $definition = [ordered]@{
        askBeforeRefresh = [bool]$normalized['askBeforeRefresh']
        trigger          = $trigger
        actions          = @($actions.ToArray())
    }
    return ($definition | ConvertTo-Json -Depth 8 -Compress)
}

function Find-DuplicateRuleDefinition {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Candidate,
        [object[]]$ExistingRules = @(),
        [string]$ExcludeRuleId = ''
    )

    $candidateSignature = Get-RuleDefinitionSignature -Rule $Candidate
    foreach ($rawExisting in @($ExistingRules)) {
        $existing = ConvertTo-NormalizedRule $rawExisting
        if ((-not [string]::IsNullOrWhiteSpace($ExcludeRuleId)) -and [string]::Equals([string]$existing['id'], $ExcludeRuleId, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        if ((Get-RuleDefinitionSignature -Rule $existing) -eq $candidateSignature) { return $existing }
    }
    return $null
}

function Import-AppConfiguration {
    <#
        Never throws for a missing / corrupt file: a monitoring app that refuses
        to start because of one bad character is worse than one that starts with
        defaults and says so loudly in the log.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $config = @{
        schemaVersion = (Get-ConfigSchemaVersion)
        appSettings   = (Get-DefaultAppSettings)
        rules         = @()
        loadError     = ''
        migrationRequired = $false
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return $config
    }

    try {
        $raw = [System.IO.File]::ReadAllText($Path)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $config }

        $parsed = ConvertTo-HashtableDeep ($raw | ConvertFrom-Json)
        $parsedSchema = ConvertTo-IntValue $parsed['schemaVersion'] 1 1
        $config.schemaVersion = $parsedSchema
        if ($parsedSchema -lt (Get-ConfigSchemaVersion)) { $config.migrationRequired = $true }
        if ($parsedSchema -gt (Get-ConfigSchemaVersion)) {
            $config.loadError = ('This configuration uses schema version {0}, which is newer than this application supports ({1}). Unknown settings will be preserved where possible.' -f $parsedSchema, (Get-ConfigSchemaVersion))
        }
        $config.appSettings = Merge-DefaultValues -Default (Get-DefaultAppSettings) -Value $parsed['appSettings']

        $settings = $config.appSettings
        # The feedback/reporting feature was removed. Strip the
        # old destination and credential fields during import so the next save
        # cannot carry stale GitHub tokens or report settings forward.
        foreach ($retired in @('githubRepo', 'githubTokenProtected', 'feedbackFormUrl',
                'feedbackEmail', 'errorReportMode', 'errorReportFolder', 'rulesPaneHeight',
                'startupWorkbookScanTimeoutSeconds')) {
            if ($settings.ContainsKey($retired)) {
                [void]$settings.Remove($retired)
                $config.migrationRequired = $true
            }
        }
        $settings['watcherHealthCheckSeconds'] = ConvertTo-IntValue $settings['watcherHealthCheckSeconds'] 60 10
        $settings['defaultDebounceSeconds']    = ConvertTo-IntValue $settings['defaultDebounceSeconds'] 5 0
        $settings['defaultCooldownSeconds']    = ConvertTo-IntValue $settings['defaultCooldownSeconds'] 30 0
        $settings['defaultRefreshWarningSeconds'] = ConvertTo-IntValue $settings['defaultRefreshWarningSeconds'] 300 5
        $settings['logRetentionDays']          = ConvertTo-IntValue $settings['logRetentionDays'] 30 1
        $settings['logMaxTotalMegabytes']      = ConvertTo-IntValue $settings['logMaxTotalMegabytes'] 50 1
        $settings['enginePollMilliseconds']    = ConvertTo-IntValue $settings['enginePollMilliseconds'] 250 100
        if ([string]$settings['rulesPaneMode'] -ne 'ShowAll') { $settings['rulesPaneMode'] = 'Compact' }
        foreach ($flag in @('startWithWindows', 'startMinimized', 'minimizeToTray', 'showSuccessNotifications',
                'showErrorNotifications', 'showTriggerNotifications', 'debugLogging', 'coalesceDuplicateWorkbooks',
                'checkDataSources', 'autoDismissExcelDialogs', 'allowWorkbookMacrosByDefault')) {
            $settings[$flag] = ConvertTo-BoolValue $settings[$flag] ((Get-DefaultAppSettings)[$flag])
        }

        $rules = New-Object System.Collections.ArrayList
        if ($null -ne $parsed['rules']) {
            foreach ($rawRule in @($parsed['rules'])) {
                $normalizedRule = ConvertTo-NormalizedRule $rawRule
                # Only a rule that is already on disk gets a placeholder name,
                # so a hand-edited rules.json can never produce a blank row.
                if ([string]::IsNullOrWhiteSpace($normalizedRule['name'])) { $normalizedRule['name'] = 'Untitled rule' }
                [void]$rules.Add($normalizedRule)
            }
        }
        $config.rules = @($rules.ToArray())
    }
    catch {
        $config.loadError = $_.Exception.Message
    }

    return $config
}

function Save-AppConfiguration {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Config
    )

    $payload = @{
        schemaVersion = (Get-ConfigSchemaVersion)
        appSettings   = $Config.appSettings
        rules         = @($Config.rules)
    }

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop | Out-Null
    }

    $json = $payload | ConvertTo-Json -Depth 8
    $encoding = New-Object System.Text.UTF8Encoding($false)

    # Write to a temp file first so a crash mid-write cannot destroy the config.
    $temp = $Path + '.tmp'
    [System.IO.File]::WriteAllText($temp, $json, $encoding)
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination ($Path + '.bak') -Force -ErrorAction SilentlyContinue
    }
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function New-RuleTemplate {
    param([hashtable]$AppSettings)

    $trigger = Get-DefaultTrigger
    if ($null -ne $AppSettings) {
        $trigger['debounceSeconds'] = ConvertTo-IntValue $AppSettings['defaultDebounceSeconds'] 5 0
        $trigger['cooldownSeconds'] = ConvertTo-IntValue $AppSettings['defaultCooldownSeconds'] 30 0
    }

    return @{
        id      = (New-ShortId -Prefix 'RULE')
        name    = ''
        enabled = $true
        askBeforeRefresh = $false
        trigger = $trigger
        actions = @()
    }
}

function Test-RuleConfiguration {
    <#
        Returns @{ Errors = @(); Warnings = @() }.
        Errors block saving. Warnings do not - an unreachable network share is
        expected on a laptop and must not prevent the rule from being stored

    #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    $errors   = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList

    if ([string]::IsNullOrWhiteSpace([string]$Rule['name'])) {
        [void]$errors.Add('Rule name is required.')
    }

    $trigger = $Rule['trigger']
    if ($null -eq $trigger) {
        [void]$errors.Add('Trigger configuration is missing.')
    }
    else {
        if (Test-TriggerUsesWatcher $trigger['type']) {
            if ([string]::IsNullOrWhiteSpace([string]$trigger['path'])) {
                if (Test-TriggerUsesFolder $trigger['type']) { [void]$errors.Add('A watch folder is required for this trigger type.') }
                else { [void]$errors.Add('A target file is required for this trigger type.') }
            }
            else {
                $probe = $trigger['path']
                if (-not (Test-TriggerUsesFolder $trigger['type'])) { $probe = Split-Path -Parent $trigger['path'] }

                $reachable = $false
                try { $reachable = Test-Path -LiteralPath $probe } catch { $reachable = $false }
                if (-not $reachable) {
                    [void]$warnings.Add(('Path currently unavailable: {0}' -f $probe))
                }
            }

            if (Test-TriggerUsesFolder $trigger['type']) {
                $filter = [string]$trigger['filter']
                if ([string]::IsNullOrWhiteSpace($filter)) {
                    [void]$warnings.Add('No file filter set - every file in the folder will be considered.')
                }
                else {
                    # NOT Path.GetInvalidFileNameChars(): on Windows that list
                    # contains * and ?, which are exactly what a filter is made
                    # of. Only characters that cannot appear in a path at all
                    # are rejected here.
                    $illegal = @('"', '<', '>', '|', ':', '/', '\\')
                    $badChar = $false
                    foreach ($character in $filter.ToCharArray()) {
                        if (($illegal -contains [string]$character) -or ([int]$character -lt 32)) {
                            $badChar = $true
                            break
                        }
                    }
                    if ($badChar) {
                        [void]$errors.Add('The file type can contain letters, digits, * and ?, for example *.csv')
                    }
                }
            }
        }

        if (Test-TriggerIsScheduled $trigger['type']) {
            if ($null -eq (ConvertTo-ScheduleTime ([string]$trigger['scheduleTime']))) {
                [void]$errors.Add('Enter the time as HH:mm on a 24-hour clock, for example 07:30.')
            }
            if (@($trigger['scheduleDays']).Count -eq 0) {
                [void]$errors.Add('Select at least one day to run on.')
            }
        }

        if ((ConvertTo-IntValue $trigger['debounceSeconds'] -1) -lt 0) { [void]$errors.Add('Combine window must be 0 or greater.') }
        if ((ConvertTo-IntValue $trigger['cooldownSeconds'] -1) -lt 0) { [void]$errors.Add('Minimum interval between runs must be 0 or greater.') }
        if ((ConvertTo-IntValue $trigger['readyTimeoutSeconds'] 0) -le 0) { [void]$errors.Add('Ready check timeout must be greater than 0.') }
        if ((ConvertTo-IntValue $trigger['readyCheckIntervalSeconds'] 0) -le 0) { [void]$errors.Add('Ready check interval must be greater than 0.') }
    }

    $actions = @($Rule['actions'])
    if ($actions.Count -eq 0) {
        [void]$errors.Add('At least one Excel action is required.')
    }
    else {
        $seenWorkbookPaths = @{}
        for ($actionIndex = 0; $actionIndex -lt $actions.Count; $actionIndex++) {
            $action = $actions[$actionIndex]
            if ([string]::IsNullOrWhiteSpace([string]$action['path'])) {
                [void]$errors.Add('An Excel action has no workbook path.')
                continue
            }
            $comparisonPath = ConvertTo-RuleComparisonPath ([string]$action['path'])
            if ($seenWorkbookPaths.ContainsKey($comparisonPath)) {
                $duplicateMessage = 'The same workbook cannot be added twice to one rule: {0}' -f [string]$action['path']
                [void]$errors.Add($duplicateMessage)
            }
            else { $seenWorkbookPaths[$comparisonPath] = $true }
            if ((ConvertTo-IntValue $action['timeoutSeconds'] 0) -le 0) {
                [void]$errors.Add(('Long refresh warning threshold must be greater than 0 ({0}).' -f (Split-Path -Leaf $action['path'])))
            }
            if ([string]$action['refreshMethod'] -eq 'SelectedQueries' -and @($action['selectedQueries']).Count -eq 0) {
                [void]$errors.Add(('Selected queries is enabled but no query is selected ({0}).' -f (Split-Path -Leaf $action['path'])))
            }
            $exists = $false
            try { $exists = Test-Path -LiteralPath $action['path'] } catch { $exists = $false }
            if (-not $exists) {
                [void]$warnings.Add(('Workbook currently unavailable: {0}' -f $action['path']))
            }
        }
    }

    return @{ Errors = @($errors.ToArray()); Warnings = @($warnings.ToArray()) }
}

# ------------------------------------------------------------------------------
# Region: persisted state
# ------------------------------------------------------------------------------

function Import-AppState {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @{ rules = @{} } }
    try {
        $raw = [System.IO.File]::ReadAllText($Path)
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{ rules = @{} } }
        $state = ConvertTo-HashtableDeep ($raw | ConvertFrom-Json)
        if ($null -eq $state['rules']) { $state['rules'] = @{} }
        return $state
    }
    catch {
        return @{ rules = @{} }
    }
}

function Save-AppState {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Shared
    )

    try {
        $ruleStates = @{}
        foreach ($ruleId in @($Shared.RuleState.Keys)) {
            # One-off "Refresh a Workbook" jobs use a synthetic rule id; there is
            # nothing worth remembering about them between sessions.
            if ($ruleId -eq 'MANUAL-REFRESH') { continue }
            $entry = $Shared.RuleState[$ruleId]
            $ruleStates[$ruleId] = @{
                lastTrigger    = (Format-StateTime $entry['LastTrigger'])
                lastSuccess    = (Format-StateTime $entry['LastSuccess'])
                lastError      = [string]$entry['LastError']
                lastResultText = [string]$entry['LastResultText']
                lastRun        = (Format-StateTime $entry['LastRun'])
                lastDuration   = [string]$entry['LastDuration']
                lastRunStatus  = [string]$entry['LastRunStatus']
                lastScheduledRun = [string]$entry['LastScheduledRun']
            }
        }

        $payload = @{
            savedAt   = (Get-Date).ToString('o')
            lastStatus = [string]$Shared.Status
            rules     = $ruleStates
        }

        $encoding = New-Object System.Text.UTF8Encoding($false)
        $directory = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop | Out-Null
        }
        $temp = $Path + '.tmp'
        [System.IO.File]::WriteAllText($temp, ($payload | ConvertTo-Json -Depth 6), $encoding)
        Move-Item -LiteralPath $temp -Destination $Path -Force
    }
    catch {
        try { if (Test-Path -LiteralPath ($Path + '.tmp')) { Remove-Item -LiteralPath ($Path + '.tmp') -Force -ErrorAction SilentlyContinue } } catch { }
    }
}

function Format-StateTime {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [DateTime]) { return $Value.ToString('o') }
    return [string]$Value
}

function Restore-RuleStateFromDisk {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$State
    )

    if ($null -eq $State['rules'] -or $State['rules'] -isnot [System.Collections.IDictionary]) { return }
    $skipped = 0
    foreach ($rawRuleId in @($State['rules'].Keys)) {
        $ruleId = [string]$rawRuleId
        # Older or interrupted builds could leave an empty key in state.json.
        # A mandatory string parameter rejects that key and used to stop the
        # whole engine even when the current configuration contained no rules.
        if ([string]::IsNullOrWhiteSpace($ruleId)) { $skipped++; continue }

        try {
            $entry = $State['rules'][$rawRuleId]
            if ($entry -isnot [System.Collections.IDictionary]) { $skipped++; continue }
            $values = @{
                LastResultText   = [string]$entry['lastResultText']
                LastError        = [string]$entry['lastError']
                LastDuration     = [string]$entry['lastDuration']
                LastRunStatus    = [string]$entry['lastRunStatus']
                LastScheduledRun = [string]$entry['lastScheduledRun']
            }
            foreach ($pair in @(@('lastTrigger', 'LastTrigger'), @('lastSuccess', 'LastSuccess'), @('lastRun', 'LastRun'))) {
                $text = [string]$entry[$pair[0]]
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $parsed = [DateTime]::MinValue
                    if ([DateTime]::TryParse($text, [ref]$parsed)) { $values[$pair[1]] = $parsed }
                }
            }
            Set-RuleState -Shared $Shared -RuleId $ruleId -Values $values | Out-Null
        }
        catch { $skipped++ }
    }
    if ($skipped -gt 0) {
        try { Write-AppLog -Level 'WARN' -Message ('Ignored {0} invalid saved rule-state entr{1}.' -f $skipped, $(if ($skipped -eq 1) { 'y' } else { 'ies' })) } catch { }
    }
}
