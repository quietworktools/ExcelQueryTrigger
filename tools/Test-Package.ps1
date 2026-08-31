# Run on Windows PowerShell 5.1 before publishing a release package.
[CmdletBinding()]
param(
    [string]$Root = (Split-Path -Parent (Split-Path -Parent $PSCommandPath))
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
$failures = New-Object System.Collections.ArrayList

function Add-Failure([string]$Message) {
    [void]$failures.Add($Message)
    Write-Host ('FAIL  ' + $Message) -ForegroundColor Red
}

Write-Host ('Checking package: {0}' -f $Root)

# Parse every PowerShell file with the same parser used by Windows PowerShell.
foreach ($file in @(Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.ps1' -File)) {
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName, [ref]$tokens, [ref]$parseErrors)
    foreach ($parseError in @($parseErrors)) {
        Add-Failure ('{0}:{1} {2}' -f $file.FullName, $parseError.Extent.StartLineNumber, $parseError.Message)
    }
}

$required = @('Setup.cmd', 'ExcelQueryTrigger.ps1', 'src\Common.ps1',
    'src\UIUpdate.ps1', 'src\UISplash.ps1', 'tools\Setup.ps1')
foreach ($relative in $required) {
    if (-not (Test-Path -LiteralPath (Join-Path $Root $relative) -PathType Leaf)) {
        Add-Failure ('Required file is missing: {0}' -f $relative)
    }
}

$commonText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Common.ps1'))
# The version is read rather than compared with a literal. Hard-coding it here
# means every release needs two edits, and forgetting the second one fails the
# build for a reason that has nothing to do with the package.
$versionMatch = [regex]::Match($commonText, "AppVersion\s*=\s*'([^']+)'")
if (-not $versionMatch.Success) {
    Add-Failure 'src\Common.ps1 does not declare AppVersion.'
}
elseif ($versionMatch.Groups[1].Value -notmatch '^\d+\.\d+(?:\.\d+)?$') {
    Add-Failure ('AppVersion must look like 0.9 or 0.9.0, but it is {0}.' -f $versionMatch.Groups[1].Value)
}
else {
    $packageVersion = $versionMatch.Groups[1].Value
    Write-Host ('Package version: {0}' -f $packageVersion)
    # The publish workflow cuts its release notes out of RELEASE-NOTES.md by
    # version number, so a missing entry silently produces an empty release.
    $notesText = [System.IO.File]::ReadAllText((Join-Path $Root 'RELEASE-NOTES.md'))
    $releaseHeadingPattern = '(?m)^##\s+v' + [regex]::Escape($packageVersion) + '(?:\s|$)'
    if ($notesText -notmatch $releaseHeadingPattern) {
        Add-Failure ('RELEASE-NOTES.md has no entry for v{0}.' -f $packageVersion)
    }
}

$entryText = [System.IO.File]::ReadAllText((Join-Path $Root 'ExcelQueryTrigger.ps1'))
foreach ($requiredText in @('ExcelQueryTriggerActivateExisting', 'ExcelQueryTriggerActivateAcknowledged', 'AbandonedMutexException')) {
    if ($entryText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Single-instance recovery is missing required behavior: {0}' -f $requiredText)
    }
}
$exceptionModePosition = $entryText.IndexOf('SetUnhandledExceptionMode')
$splashPosition = $entryText.IndexOf('Show-StartupSplash')
if ($exceptionModePosition -lt 0 -or $splashPosition -lt 0 -or
    $exceptionModePosition -gt $splashPosition) {
    Add-Failure 'WinForms exception mode must be set before the splash creates controls.'
}

$configPath = Join-Path $Root 'config\rules.json'
try {
    $config = [System.IO.File]::ReadAllText($configPath) | ConvertFrom-Json
    if ([int]$config.schemaVersion -ne 4) { Add-Failure 'config\rules.json must use schema version 4.' }
    if ($config.appSettings.PSObject.Properties.Name -notcontains 'checkForUpdatesAutomatically' -or
        -not [bool]$config.appSettings.checkForUpdatesAutomatically) {
        Add-Failure 'Automatic update checking must exist and default to on in config\rules.json.'
    }
    if (@($config.rules).Count -ne 0) {
        Add-Failure 'config\rules.json must be a clean template with no saved rules.'
    }
    $retiredKeys = @('githubRepo', 'githubTokenProtected', 'feedbackFormUrl',
        'feedbackEmail', 'errorReportMode', 'errorReportFolder', 'rulesPaneHeight',
        'startupWorkbookScanTimeoutSeconds')
    foreach ($key in $retiredKeys) {
        if ($config.appSettings.PSObject.Properties.Name -contains $key) {
            Add-Failure ('Retired setting remains in config\rules.json: {0}' -f $key)
        }
    }
}
catch { Add-Failure ('config\rules.json is invalid: {0}' -f $_.Exception.Message) }

$runtimeGeneratedPaths = @(
    'logs',
    'config\rules.json.bak',
    'config\state.json',
    'config\welcome-seen.txt',
    'config\history.jsonl'
)
foreach ($relative in $runtimeGeneratedPaths) {
    if (Test-Path -LiteralPath (Join-Path $Root $relative)) {
        Add-Failure ('Runtime-generated path must not be committed or packaged: {0}' -f $relative)
    }
}

$requiredIgnoreEntries = @(
    'logs/',
    'config/rules.json.bak',
    'config/state.json',
    'config/welcome-seen.txt',
    'config/history.jsonl',
    'config/*.tmp'
)
$gitIgnorePath = Join-Path $Root '.gitignore'
if (-not (Test-Path -LiteralPath $gitIgnorePath -PathType Leaf)) {
    Add-Failure '.gitignore is missing.'
}
else {
    $gitIgnoreEntries = @(
        Get-Content -LiteralPath $gitIgnorePath |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
    )
    foreach ($entry in $requiredIgnoreEntries) {
        if ($gitIgnoreEntries -notcontains $entry) {
            Add-Failure ('.gitignore is missing required entry: {0}' -f $entry)
        }
    }
}

$settingsText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UISettings.ps1'))
if ($settingsText -match '(?i)feedback|errorReport|githubRepo|githubToken|tokenProtected') {
    Add-Failure 'src\UISettings.ps1 still contains a retired reporting, repository or token setting.'
}

$engineText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Engine.ps1'))
if ($engineText -match [regex]::Escape('@(Get-EngineRules)')) {
    Add-Failure 'Engine startup must not wrap Get-EngineRules in @(...); that turns the rule array into one nested item and makes watcher discovery report 0 monitors.'
}

$managerText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIManager.ps1'))
if ($managerText -match '(?i)RulesSplitter|UiRulesDragging|Update-WorkbookInfoSlice') {
    Add-Failure 'The retired draggable pane splitter or synchronous workbook scan remains.'
}
foreach ($requiredText in @('RulesViewToggle', 'Show all', 'Compact',
        'Start-WorkbookInfoBackgroundScan', 'Reading workbook information',
        'Get-StartupMonitoringLines', 'Update-CompletedJobWorkbookInfo',
        '$script:UiControls.Tray.Visible = $true',
        'Start-RefreshApprovalBackgroundCheck', 'Stop-RefreshApprovalBackgroundCheck',
        'RecentRefreshPromptMinutes', 'UIDiagnostics.ps1',
        'ExcelQueryTriggerActivateExisting', 'ActivateAckEvent')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Dashboard is missing required behavior: {0}' -f $requiredText)
    }
}

$splashText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UISplash.ps1'))
foreach ($requiredText in @(
        'Overall startup progress',
        'ActivityText',
        'The Dashboard opens only after monitoring is ready'
    )) {
    if ($splashText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The startup dialog is missing required behavior: {0}' -f $requiredText)
    }
}

$aboutText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIUpdate.ps1'))
foreach ($requiredText in @('Version and updates', 'Installed version', 'Release notes',
        'function ConvertTo-ReleaseNotesDisplayText', '$bulletPrefix', '-split "`r?`n"',
        'ConvertTo-ReleaseNotesDisplayText -Markdown')) {
    if ($aboutText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The About dialog is missing required layout: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('StartupMonitoringReady', 'StartupUiReady', 'StartupReady')) {
    if ($managerText -notmatch [regex]::Escape($requiredText) -and
        $commonText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Startup readiness is missing required behavior: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @(
        '[string]$Shared.Status -ne ''Running''',
        '$Shared.StartupUiReady = $true',
        'Final check - preparing the Dashboard'
    )) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Dashboard final startup gate is missing: {0}' -f $requiredText)
    }
}
if ($managerText -match [regex]::Escape('Startup workbook information scan exceeded its foreground time budget')) {
    Add-Failure 'The Dashboard can still open before the startup workbook-information scan finishes.'
}
$closeSplashPosition = $managerText.IndexOf('Close-StartupSplash -Splash $Splash')
$showDashboardPosition = $managerText.IndexOf('$form.Show()', $closeSplashPosition)
if ($closeSplashPosition -lt 0 -or $showDashboardPosition -lt $closeSplashPosition) {
    Add-Failure 'The Dashboard can become visible before the startup splash closes.'
}

$commonPath = Join-Path $Root 'src\Common.ps1'
$configManagerPath = Join-Path $Root 'src\ConfigManager.ps1'
. $commonPath
. $configManagerPath
. (Join-Path $Root 'src\ExcelManager.ps1')
. (Join-Path $Root 'src\JobManager.ps1')
. (Join-Path $Root 'src\UIDiagnostics.ps1')
$defaultAppSettingsForTest = Get-DefaultAppSettings
if (-not (ConvertTo-BoolValue $defaultAppSettingsForTest['checkForUpdatesAutomatically'] $false)) {
    Add-Failure 'Automatic update checking does not default to on.'
}
if ($defaultAppSettingsForTest.ContainsKey('startupWorkbookScanTimeoutSeconds')) {
    Add-Failure 'The retired startup workbook scan timeout setting is still present.'
}

$unsafeSavedActionRule = ConvertTo-NormalizedRule @{
    id = 'FIXED-ACTION'; name = 'Hand edited'; enabled = $true
    trigger = @{ type = 'Manual' }
    actions = @(@{ path = 'C:\Reports\Safety.xlsx'; save = $false; close = $false; visible = $true })
}
$fixedAction = @($unsafeSavedActionRule.actions)[0]
if (-not [bool]$fixedAction.save -or -not [bool]$fixedAction.close -or [bool]$fixedAction.visible) {
    Add-Failure 'Legacy or hand-edited action values can bypass hidden/save/close safety invariants.'
}

$queryAIdentity = Get-RefreshActionIdentity @{
    type = 'ExcelRefresh'; path = 'C:\Reports\.\Sales.xlsx'; refreshMethod = 'SelectedQueries'
    selectedQueries = @('A'); allowWorkbookMacros = $false; disableBackgroundQuery = $true
    continueOnError = $true; timeoutSeconds = 300
}
$queryBIdentity = Get-RefreshActionIdentity @{
    type = 'ExcelRefresh'; path = 'c:/reports/Sales.xlsx'; refreshMethod = 'SelectedQueries'
    selectedQueries = @('B'); allowWorkbookMacros = $false; disableBackgroundQuery = $true
    continueOnError = $true; timeoutSeconds = 300
}
if ($queryAIdentity -eq $queryBIdentity) {
    Add-Failure 'Queue identity still conflates different selected-query operations on one workbook.'
}
$macroIdentityOff = Get-RefreshActionIdentity @{
    type='ExcelRefresh'; path='C:\Reports\Large.xlsx'; refreshMethod='RefreshAll'; selectedQueries=@()
    allowWorkbookMacros=$false; disableBackgroundQuery=$true; continueOnError=$true; timeoutSeconds=300
}
$macroIdentityOn = Get-RefreshActionIdentity @{
    type='ExcelRefresh'; path='C:\Reports\Large.xlsx'; refreshMethod='RefreshAll'; selectedQueries=@()
    allowWorkbookMacros=$true; disableBackgroundQuery=$true; continueOnError=$true; timeoutSeconds=300
}
if ($macroIdentityOff -eq $macroIdentityOn) {
    Add-Failure 'Queue identity conflates different workbook macro policies.'
}

$quietWorkbook = [pscustomobject]@{ Connections = @(); Worksheets = @() }
$quietExcel = [pscustomobject]@{ CalculationState = 0 }
if ((Get-WorkbookRefreshState -Excel $quietExcel -Workbook $quietWorkbook) -ne 'Quiet') {
    Add-Failure 'A fully observable idle workbook did not report Quiet.'
}
$throwingWorkbookType = @'
using System;
public sealed class EqtThrowingWorkbookState
{
    public object Connections { get { throw new InvalidOperationException("Simulated RPC rejection"); } }
    public object[] Worksheets { get { return new object[0]; } }
}
'@
if (-not ('EqtThrowingWorkbookState' -as [type])) {
    Add-Type -TypeDefinition $throwingWorkbookType -Language CSharp
}
$unknownWorkbook = New-Object EqtThrowingWorkbookState
if ((Get-WorkbookRefreshState -Excel $quietExcel -Workbook $unknownWorkbook) -ne 'Unknown') {
    Add-Failure 'A top-level COM-state read failure was not reported as Unknown.'
}
$unconfirmedExcel = [pscustomobject]@{ CalculationState = 0 }
$unconfirmedExcel | Add-Member -MemberType ScriptMethod -Name CalculateUntilAsyncQueriesDone -Value { throw 'Simulated confirmation failure' }
$confirmationFailedSafely = $false
try {
    [void](Wait-ExcelRefreshCompletion -Excel $unconfirmedExcel -Workbook $quietWorkbook `
        -PollMilliseconds 1 -RequiredQuietSamples 1)
}
catch {
    $confirmationFailedSafely = ($_.Exception.Message -like '*will not be saved*')
}
if (-not $confirmationFailedSafely) {
    Add-Failure 'Failed asynchronous completion confirmation did not stop the save path.'
}

$stageShared = New-SharedState
Initialize-JobManager -Shared $stageShared
$stageJob = New-PendingJob -Shared $stageShared -Rule $unsafeSavedActionRule
Set-CurrentJobDisplay -Job $stageJob -Stage 'Refreshing'
if (-not [bool]$stageShared.CurrentJob.CanCancel -or -not [bool]$stageShared.CurrentJob.CanForceTerminate) {
    Add-Failure 'Refreshing stage does not expose the expected safe cancellation permissions.'
}
Set-CurrentJobDisplay -Job $stageJob -Stage 'Saving'
if ([bool]$stageShared.CurrentJob.CanCancel -or [bool]$stageShared.CurrentJob.CanForceTerminate) {
    Add-Failure 'Saving stage still permits cancellation or forced process termination.'
}
# A direct save may not be interrupted, but a modal dialog raised during it
# must still be visible to the watchdog, or the job waits for ever.
if (-not [bool]$stageShared.CurrentJob.CanInspectDialogs) {
    Add-Failure 'Saving stage hides Excel dialogs from the watchdog.'
}

# Exercise queue semantics without touching Excel or disk logs.
function Write-AppLog {
    param($Level, $RuleName, $Workbook, $Message, $Stage, $ErrorType, [switch]$NoActivity)
}
$queueShared = New-SharedState
$queueRuleA = ConvertTo-NormalizedRule @{
    id = 'QUEUE-A'; name = 'Query A'; enabled = $true; trigger = @{ type = 'Manual' }
    actions = @(@{ path = 'C:\Reports\Sales.xlsx'; refreshMethod = 'SelectedQueries'; selectedQueries = @('A') })
}
$queueRuleB = ConvertTo-NormalizedRule @{
    id = 'QUEUE-B'; name = 'Query B'; enabled = $true; trigger = @{ type = 'Manual' }
    actions = @(@{ path = 'c:/reports/Sales.xlsx'; refreshMethod = 'SelectedQueries'; selectedQueries = @('B') })
}
$queueJobA = New-PendingJob -Shared $queueShared -Rule $queueRuleA
$queueJobB = New-PendingJob -Shared $queueShared -Rule $queueRuleB
[void](Add-PendingJob -Shared $queueShared -Job $queueJobA -CoalesceDuplicateWorkbooks $true)
[void](Add-PendingJob -Shared $queueShared -Job $queueJobB -CoalesceDuplicateWorkbooks $true)
if ((Get-PendingJobCount -Shared $queueShared) -ne 2) {
    Add-Failure 'Queue coalescing dropped a different selected-query operation on the same workbook.'
}

$sameRuleJob1 = New-PendingJob -Shared $queueShared -Rule $queueRuleA -TriggerFile 'C:\Drop\A.csv'
$sameRuleJob2 = New-PendingJob -Shared $queueShared -Rule $queueRuleA -TriggerFile 'C:\Drop\B.csv'
$sameRuleShared = New-SharedState
[void](Add-PendingJob -Shared $sameRuleShared -Job $sameRuleJob1 -CoalesceDuplicateWorkbooks $true)
[void](Add-PendingJob -Shared $sameRuleShared -Job $sameRuleJob2 -CoalesceDuplicateWorkbooks $true)
$mergedJob = @((Get-PendingJobSnapshot -Shared $sameRuleShared))[0]
if ($null -eq $mergedJob -or @($mergedJob.TriggerFiles).Count -ne 2) {
    Add-Failure 'Identical queued rule runs did not retain both distinct source files.'
}

. (Join-Path $Root 'src\TriggerManager.ps1')
$triggerShared = New-SharedState
Initialize-TriggerManager -Shared $triggerShared
$multiFileRule = ConvertTo-NormalizedRule @{
    id = 'TRIGGER-MULTI'; name = 'Multi-file debounce'; enabled = $true
    trigger = @{ type = 'FileChangedAny'; path = 'C:\Drop'; filter = '*.csv'; debounceSeconds = 0; cooldownSeconds = 0 }
    actions = @(@{ path = 'C:\Reports\Sales.xlsx' })
}
[void](Register-TriggerEvent -Rule $multiFileRule -FullPath 'C:\Drop\A.csv')
[void](Register-TriggerEvent -Rule $multiFileRule -FullPath 'C:\Drop\B.csv')
$dueMultiFile = @(Get-DueTriggers -Rules @($multiFileRule))
if ($dueMultiFile.Count -ne 1 -or @($dueMultiFile[0].FilePaths).Count -ne 2) {
    Add-Failure 'Debounce did not retain both distinct source files in one rule window.'
}

# Exercise the Add-time workbook inspection against minimal Open XML packages.
# This does not need Excel and catches regressions where the UI still calls a
# checker but the checker no longer recognises connections.xml.
$queryWorkbook = ''
$emptyWorkbook = ''
try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $queryWorkbook = Join-Path ([System.IO.Path]::GetTempPath()) ('eqt-query-test-{0}.xlsx' -f ([guid]::NewGuid().ToString('N')))
    $emptyWorkbook = Join-Path ([System.IO.Path]::GetTempPath()) ('eqt-empty-test-{0}.xlsx' -f ([guid]::NewGuid().ToString('N')))
    foreach ($packageSpec in @(
            @{ Path = $queryWorkbook; WithConnection = $true },
            @{ Path = $emptyWorkbook; WithConnection = $false })) {
        $archive = [System.IO.Compression.ZipFile]::Open(
            [string]$packageSpec.Path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            if ([bool]$packageSpec.WithConnection) {
                $entry = $archive.CreateEntry('xl/connections.xml')
                $writer = New-Object System.IO.StreamWriter($entry.Open())
                try {
                    $writer.Write('<?xml version="1.0" encoding="UTF-8"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="Power Query - Sales"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1" command="Sales" /></connection></connections>')
                }
                finally { $writer.Dispose() }
            }
        }
        finally { $archive.Dispose() }
    }

    $queryCheck = Test-WorkbookHasQueries -Path $queryWorkbook
    if (-not $queryCheck.Decided -or -not $queryCheck.HasQueries -or [int]$queryCheck.Count -ne 1) {
        Add-Failure 'Add-time workbook inspection did not recognise a refreshable connection.'
    }
    $emptyCheck = Test-WorkbookHasQueries -Path $emptyWorkbook
    if (-not $emptyCheck.Decided -or $emptyCheck.HasQueries -or [int]$emptyCheck.Count -ne 0) {
        Add-Failure 'Add-time workbook inspection did not reject an empty workbook.'
    }
}
catch { Add-Failure ('Add-time workbook inspection test failed: {0}' -f $_.Exception.Message) }
finally {
    foreach ($temporaryWorkbook in @($queryWorkbook, $emptyWorkbook)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$temporaryWorkbook)) {
            Remove-Item -LiteralPath $temporaryWorkbook -Force -ErrorAction SilentlyContinue
        }
    }
}

$legacyRule = ConvertTo-NormalizedRule @{
    id = 'TEST'; name = 'Legacy rule'; enabled = $true
    trigger = @{ type = 'Manual' }; actions = @()
}
if ([int]$legacyRule.trigger.recentRefreshPromptMinutes -ne 0) {
    Add-Failure 'A legacy rule must default recentRefreshPromptMinutes to 0.'
}
$guardedRule = ConvertTo-NormalizedRule @{
    id = 'TEST'; name = 'Guarded rule'; enabled = $true
    trigger = @{ type = 'Scheduled'; recentRefreshPromptMinutes = '90' }; actions = @()
}
if ([int]$guardedRule.trigger.recentRefreshPromptMinutes -ne 90) {
    Add-Failure 'recentRefreshPromptMinutes was not normalized correctly.'
}

$duplicateBase = ConvertTo-NormalizedRule @{
    id = 'RULE-A'; name = 'Morning refresh'; enabled = $true; askBeforeRefresh = $false
    trigger = @{ type = 'Scheduled'; scheduleTime = '07:30'; scheduleDays = @('Monday', 'Tuesday') }
    actions = @(@{ path = 'C:\Reports\Sales.xlsx'; refreshMethod = 'SelectedQueries'; selectedQueries = @('Sales', 'Stock') })
}
$duplicateRenamed = ConvertTo-NormalizedRule @{
    id = 'RULE-B'; name = 'Different display name'; enabled = $false; askBeforeRefresh = $false
    trigger = @{ type = 'Scheduled'; scheduleTime = '07:30'; scheduleDays = @('Monday', 'Tuesday') }
    actions = @(@{ path = 'c:\reports\sales.xlsx'; refreshMethod = 'SelectedQueries'; selectedQueries = @('Stock', 'Sales') })
}
$differentWorkbook = ConvertTo-NormalizedRule @{
    id = 'RULE-C'; name = 'Different workbook'; enabled = $true; askBeforeRefresh = $false
    trigger = @{ type = 'Scheduled'; scheduleTime = '07:30'; scheduleDays = @('Monday', 'Tuesday') }
    actions = @(@{ path = 'C:\Reports\Inventory.xlsx'; refreshMethod = 'SelectedQueries'; selectedQueries = @('Sales', 'Stock') })
}
if ($null -eq (Find-DuplicateRuleDefinition -Candidate $duplicateRenamed -ExistingRules @($duplicateBase))) {
    Add-Failure 'An identical rule with another name was not detected.'
}
if ($null -ne (Find-DuplicateRuleDefinition -Candidate $differentWorkbook -ExistingRules @($duplicateBase))) {
    Add-Failure 'A rule with a different workbook was incorrectly treated as a duplicate.'
}
if ($null -ne (Find-DuplicateRuleDefinition -Candidate $duplicateBase -ExistingRules @($duplicateBase) -ExcludeRuleId 'RULE-A')) {
    Add-Failure 'Editing a rule without changing it was treated as a duplicate of itself.'
}

$comparisonRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'eqt-path-comparison'
$sameWorkbookActions = @(
    @{ path = [System.IO.Path]::Combine($comparisonRoot, '.', 'Sales.xlsx') },
    @{ path = [System.IO.Path]::Combine($comparisonRoot, 'Sales.xlsx') }
)
if ((Find-DuplicateWorkbookAction -Actions $sameWorkbookActions -CandidatePath ([System.IO.Path]::Combine($comparisonRoot, 'Sales.xlsx')) -ExcludeIndex 0) -ne 1) {
    Add-Failure 'Equivalent workbook paths were not detected as the same action.'
}
$duplicateWorkbookRule = ConvertTo-NormalizedRule @{
    id = 'RULE-D'; name = 'Duplicate workbook'; enabled = $true
    trigger = @{ type = 'Manual' }
    actions = $sameWorkbookActions
}
$duplicateWorkbookValidation = Test-RuleConfiguration -Rule $duplicateWorkbookRule
if (@($duplicateWorkbookValidation.Errors | Where-Object { $_ -like 'The same workbook cannot be added twice*' }).Count -ne 1) {
    Add-Failure 'Rule validation did not reject a duplicate workbook action exactly once.'
}

$engineText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Engine.ps1'))
foreach ($requiredText in @('CheckRecentRefresh', 'AutomaticApproval',
        'WorkbookPaths', 'recentRefreshPromptMinutes')) {
    if ($engineText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The engine is missing recent-refresh behavior: {0}' -f $requiredText)
    }
}

# Public GitHub-facing text is English. Japanese remains in the dedicated
# Japanese instructions/manuals, not in the repository landing page or the
# release body generated from RELEASE-NOTES.md.
foreach ($relative in @('README.md', 'RELEASE-NOTES.md')) {
    $publicText = [System.IO.File]::ReadAllText((Join-Path $Root $relative))
    if ($publicText -match '[ぁ-んァ-ヶ一-龠]') {
        Add-Failure ('Public GitHub text contains Japanese: {0}' -f $relative)
    }
}

$ruleEditorText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIRuleEditor.ps1'))
foreach ($requiredText in @('Recent refresh protection',
        'Ask before refreshing a workbook whose queries were updated recently',
        '$gbRecentRefresh.Controls.Add($chkRecentRefresh)', 'minutes', 'hours',
        'recentRefreshPromptMinutes', 'function New-FormWrappedLabel',
        '$label.AutoSize  = $true', '$label.MinimumSize', '$label.MaximumSize',
        '$label.AutoEllipsis = $false')) {
    if ($ruleEditorText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The rule editor is missing the recent-refresh setting: {0}' -f $requiredText)
    }
}
if ($ruleEditorText -match [regex]::Escape('$gbAdvanced.Controls.Add($chkRecentRefresh)')) {
    Add-Failure 'Recent refresh protection is still inside Advanced settings.'
}

$updateText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIUpdate.ps1'))
if ($updateText -notmatch [regex]::Escape("([0-9a-f]{{64}})")) {
    Add-Failure 'Update checksum regex does not escape the SHA-256 quantifier for the PowerShell format operator.'
}
try {
    $escapedAssetNameForTest = [regex]::Escape('ExcelQueryTrigger-v0.9.6.zip')
    $checksumPatternForTest = ('(?im)^([0-9a-f]{{64}})\s+\*?{0}\s*$' -f $escapedAssetNameForTest)
    if (-not ('a' * 64 + '  ExcelQueryTrigger-v0.9.6.zip' -match $checksumPatternForTest)) {
        Add-Failure 'Update checksum regex regression test did not match a valid checksum line.'
    }
}
catch { Add-Failure ('Update checksum regex formatting failed: {0}' -f $_.Exception.Message) }

$wizardText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIAddFileWizard.ps1'))
foreach ($requiredText in @('Recent refresh protection',
        'Ask if queries were refreshed within',
        'If any workbook above was refreshed recently, ask before refreshing it again to prevent unnecessary refreshes.',
        '$gbRecentRefresh.Location = New-Object System.Drawing.Point(28, 376)',
        '$gbRecentRefresh.Size = New-Object System.Drawing.Size(524, 98)',
        '$recentRefreshHint = New-FormWrappedLabel',
        '$lvWorkbooks.Size          = New-Object System.Drawing.Size(524, 190)',
        '$step1.Controls.Add($gbRecentRefresh)',
        '$rule.trigger.recentRefreshPromptMinutes = $recentMinutes')) {
    if ($wizardText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Add wizard is missing recent-refresh protection: {0}' -f $requiredText)
    }
}
if ($wizardText -match [regex]::Escape('$step2.Controls.Add($gbRecentRefresh)')) {
    Add-Failure 'Recent refresh protection is still on the trigger page instead of workbook selection.'
}

foreach ($requiredText in @('Workbooks to refresh  (processed from top to bottom)',
        "Columns.Add('Queries'", "Columns.Add('Warn after'", "Columns.Add('If it fails'")) {
    if ($ruleEditorText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The detailed workbook list is missing its simplified field: {0}' -f $requiredText)
    }
}
foreach ($retiredColumn in @("Columns.Add('Save'", "Columns.Add('Close'", "Columns.Add('Visible'")) {
    if ($ruleEditorText -match [regex]::Escape($retiredColumn)) {
        Add-Failure ('The detailed workbook list still exposes a fixed implementation column: {0}' -f $retiredColumn)
    }
}

$commonText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Common.ps1'))
$jobManagerText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\JobManager.ps1'))
foreach ($requiredText in @('WorkbookInfoRefreshRequests', 'Update-CompletedJobWorkbookInfo')) {
    $found = ($commonText -match [regex]::Escape($requiredText)) -or
             ($jobManagerText -match [regex]::Escape($requiredText)) -or
             ($managerText -match [regex]::Escape($requiredText))
    if (-not $found) {
        Add-Failure ('Immediate Data updated refresh hand-off is missing: {0}' -f $requiredText)
    }
}

# Reading a workbook for its metadata must never stop Excel from saving it.
# Excel replaces the original file on save, so the reader has to allow deletion,
# and it must not read a workbook a running job has open at all.
if ($commonText -notmatch [regex]::Escape('[System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete')) {
    Add-Failure 'The workbook metadata reader can still block Excel from replacing a saved workbook.'
}
foreach ($requiredText in @(
        'if ([bool]$script:UiShared.CurrentJob.Active) { return }',
        'if ($scan.Running) { Stop-WorkbookInfoBackgroundScan }',
        '[void]$ps.AddArgument($knownStamps)')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Workbook metadata scanning is not held back during a refresh job: {0}' -f $requiredText)
    }
}

$excelManagerText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\ExcelManager.ps1'))
$workbookInfoText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIWorkbookInfo.ps1'))
foreach ($requiredText in @('ExcelQueryTriggerLastQueryRefreshUtc',
        'Set-WorkbookQueryRefreshStamp', 'QueryRefreshedAt')) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Excel refresh completion metadata is missing: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @(
        'local job history will retain the query refresh time',
        'Remove-ComReference $workbook $workbooks',
        'Test-WorkbookLocked -Path $Action.path')) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Workbook lock-release behavior is missing: {0}' -f $requiredText)
    }
}
# Safe temporary-file saving was removed. Only a direct Excel.Save remains.
foreach ($retiredText in @('SaveCopyAs', 'SavingSafe', 'safeSave', 'Commit-SafeSavedCopy')) {
    if ($excelManagerText -match [regex]::Escape($retiredText)) {
        Add-Failure ('Retired safe-save code is still present in ExcelManager: {0}' -f $retiredText)
    }
}
foreach ($requiredText in @(
        'Get-SuccessfulQueryRefreshHistoryMap',
        'Excel Query Trigger Manager (local history)',
        '$localRefreshHistory = Get-SuccessfulQueryRefreshHistoryMap')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Local query-refresh history fallback is missing: {0}' -f $requiredText)
    }
}
if ($excelManagerText -match [regex]::Escape("-Level 'WARN' -RuleName $RuleName -Workbook $Action.path -Stage $saveStage") -and
    $excelManagerText -match [regex]::Escape('timestamp could not be written to the workbook metadata')) {
    Add-Failure 'A failed optional workbook refresh stamp still produces a warning on successful refreshes.'
}
foreach ($requiredText in @('Get-WorkbookRefreshState', "return 'Unknown'",
        'The workbook will not be saved', 'GetWindowThreadProcessId',
        '$saveStarted', 'Test-ExcelProcessIdentity')) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Excel safety behavior is missing: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('$script:UiShared.CancelCurrentJob = $true',
        'CanForceTerminate', 'OwnedExcelStartedAtUtc', 'Cannot cancel safely')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Cancellation safety behavior is missing: {0}' -f $requiredText)
    }
}
$triggerManagerText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\TriggerManager.ps1'))
foreach ($requiredText in @('PendingFiles', 'FilePaths = $filePaths')) {
    if ($triggerManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Multi-file debounce behavior is missing: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('ExcelQueryTriggerLastQueryRefreshUtc',
        "return 'Not recorded'", 'not used for Data updated')) {
    if ($workbookInfoText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Query-only Data updated behavior is missing: {0}' -f $requiredText)
    }
}
$ageFunction = [regex]::Match($workbookInfoText,
    '(?s)function Get-WorkbookDataAgeText\s*\{.*?\n\}').Value
if ($ageFunction -match '\.SavedAt') {
    Add-Failure 'Data updated still falls back to the workbook save time.'
}
$recentCheck = [regex]::Match($managerText,
    '(?s)function Start-RefreshApprovalBackgroundCheck\s*\{.*?\n\}').Value
if ($recentCheck -match '\.SavedAt') {
    Add-Failure 'Recent-refresh protection still treats the workbook save time as a query refresh.'
}

foreach ($requiredText in @('New-FormWrappedLabel -Text $Explanation',
        '-Parent $tabAdv -Y 370', '-Parent $tabAdv -Y 448',
        '-Parent $tabAdv -Y 526', '-Parent $tabAdv -Y 616',
        '-X 18 -Y 752 -Width 466')) {
    if ($settingsText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Settings is missing adaptive wrapped-text layout: {0}' -f $requiredText)
    }
}
if ($settingsText -match '\$note\.Size\s*=') {
    Add-Failure 'A Settings explanation still uses a fixed-height note label.'
}
foreach ($requiredText in @('System.Collections.Stack', '$control.AutoEllipsis = $true')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The global text-overflow guard is missing: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('$aboutState = @{', 'Update = $null', 'Install-Update -Update $aboutState.Update')) {
    if ($aboutText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The About dialog is missing required update state: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('Setup finishes automatically',
        'Ready. Closing this version and installing automatically...',
        '$lblStatus.Multiline = $true', '$lblStatus.WordWrap = $true',
        '$lblStatus.ScrollBars = ''Vertical''')) {
    if ($aboutText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The About dialog is missing the automatic update hand-off: {0}' -f $requiredText)
    }
}
if ($aboutText -match '\$script:AboutUpdate') {
    Add-Failure 'The About dialog still uses closure-unsafe script-scoped update state.'
}

# Prove the reference-sharing behavior used by the two independent WinForms
# event closures. This is the exact boundary that caused Install it to no-op.
$closureState = @{ Update = $null }
$closureWriter = { $closureState.Update = @{ Version = 'test' } }.GetNewClosure()
$closureReader = { return [string]$closureState.Update.Version }.GetNewClosure()
& $closureWriter
if ((& $closureReader) -ne 'test') {
    Add-Failure 'Reference state is not shared between independent event closures.'
}

$setupText = [System.IO.File]::ReadAllText((Join-Path $Root 'tools\Setup.ps1'))
foreach ($requiredText in @('backupNotice', 'trayNotice', '$isFirstInstall',
        'elseif ($UpdateMode)', '$btnPrimary.PerformClick()',
        '$isFirstInstall -or $UpdateMode', "`$launchArguments += ' -ShowWindow'")) {
    if ($setupText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Setup is missing required first-install behavior: {0}' -f $requiredText)
    }
}
$hiddenLauncherText = [System.IO.File]::ReadAllText((Join-Path $Root 'Start-Hidden.vbs'))
if ($hiddenLauncherText -notmatch [regex]::Escape('command = command & " -ShowWindow"')) {
    Add-Failure 'Start-Hidden.vbs does not forward the first-install ShowWindow override.'
}

$wizardText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIAddFileWizard.ps1'))
foreach ($requiredText in @('$wizardState = @{', 'ShowStep',
        'Config   = $script:UiConfig', '$wizardState.ShowStep = $showStep', '& $wizardState.ShowStep',
        'What should trigger the update?', 'Get-TriggerTypeList',
        "@('Folder', 'Specific', 'Schedule', 'Logon', 'Manual')",
        'FileChangedSpecific', 'Test-TriggerUsesFolder', 'Test-TriggerIsScheduled',
        'Test-TriggerIsLogon', 'File to watch:', 'Custom pattern',
        'Name contains:', 'Ignore if name contains:', '$clientHeight = 570',
        '$rule.trigger.type = $selectedType',
        'Which workbooks should refresh?', 'Add workbook...', 'Edit...', 'Remove', 'Move up', 'Move down',
        '$workbookState = @{ Actions = @() }', 'Show-ActionEditor',
        '$step1.Visible = ($currentStep -eq 2)', '$step2.Visible = ($currentStep -eq 1)',
        'Name and review this rule', 'Rule name', 'Workbooks to refresh (in order)', 'Trigger settings',
        'System.Windows.Forms.TableLayoutPanel', '$populateReview = {',
        '$rule.name = $cleanRuleName', '$rule.actions = @($workbookState.Actions)', '$clientHeight = 640',
        'Scheduled refreshes run only while this PC is turned on', '$tableHeight = [Math]::Min',
        'Find-DuplicateRuleDefinition -Candidate $rule',
        '$config = $wizardState.Config', '$config[''rules''] = @($rules.ToArray())',
        'They will refresh in the order shown.')) {
    if ($wizardText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Add wizard is missing required behavior: {0}' -f $requiredText)
    }
}
if ($wizardText -match '\$script:Wizard(Step|FileOk)') {
    Add-Failure 'The Add file wizard still uses closure-unsafe script-scoped state.'
}
if ($wizardText -match '\$lblSummary|\$describe\s*=') {
    Add-Failure 'The Add wizard review has fallen back to an unstructured summary paragraph.'
}
if ($wizardText -match '\$cboQueryMethod|\$queryState|\$txtPath') {
    Add-Failure 'The Add wizard still keeps a rule-level query selector or single-workbook path state.'
}

foreach ($requiredText in @('$jobActionY = 132', 'About & updates',
        'Check for updates...', 'Show-AboutDialog -Owner $form')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Dashboard/help UI is missing required behavior: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('$lblShow.AutoSize = $true',
        '$lblForRule.AutoSize = $true', '$lblContaining.AutoSize = $true',
        'New-Object System.Drawing.Point(62, 18)',
        'New-Object System.Drawing.Point(662, 18)')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Recent Activity filters are missing font-safe layout: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @("return ('{0} workbooks' -f `$names.Count)",
        '$workbookSubItem.Font = Get-UiFont 9 ''Bold''',
        '$script:UiControls.RulesHoverToolTip', '.Rules.Add_MouseMove({',
        '$columnIndex', 'FieldTooltips', 'Get-RuleQueryFieldTooltip',
        'LocationTooltip', 'DataTooltip', '$workbookTooltipLines')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Trigger Rules field-specific hover behavior is missing: {0}' -f $requiredText)
    }
}
if ($managerText -match [regex]::Escape('$item.Tag[''WorkbookTooltip'']') -or
    $managerText -match [regex]::Escape('[string]$hit.Item.ToolTipText')) {
    Add-Failure 'Trigger Rules still falls back to the retired whole-row tooltip.'
}
foreach ($requiredText in @('Warn if still running after:',
        'It does not stop the refresh; the application keeps waiting.',
        'Excel stays hidden during refresh.',
        'After success, the workbook is saved and closed.',
        'Wait for supported connections to finish before saving',
        'Allow workbook macros (advanced)',
        "`$result['save']                   = `$true",
        "`$result['close']                  = `$true",
        "`$result['visible']                = `$false")) {
    if ($ruleEditorText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The Excel action dialog is missing required behavior: {0}' -f $requiredText)
    }
}
foreach ($requiredText in @('Automatically check for updates',
        "`$working['checkForUpdatesAutomatically'] `$true",
        "`$result['lastUpdateCheckUtc'] = ''")) {
    if ($settingsText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Settings is missing automatic update-check control: {0}' -f $requiredText)
    }
}
foreach ($removedChoice in @('Save after refresh', 'Close after refresh',
        'Show the Excel window while refreshing', 'Run supported queries synchronously')) {
    if ($ruleEditorText -match [regex]::Escape($removedChoice)) {
        Add-Failure ('The Excel action dialog still exposes a retired choice: {0}' -f $removedChoice)
    }
}
foreach ($requiredText in @('$form.ClientSize      = New-Object System.Drawing.Size(900, 700)',
        '$tabUpdates.AutoScroll = $true', '$tabRefresh.AutoScroll = $true',
        '$tabMonitor.AutoScroll = $true', '$tabTerms.AutoScroll = $true',
        '$tabTrouble.AutoScroll = $true', 'Updating the application',
        '1.  Check for updates', '2.  Choose Install it', '3.  Let Setup finish',
        'Your rules, history, logs, and settings are retained during an update.')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Information & Help is missing required adaptive layout: {0}' -f $requiredText)
    }
}
if ($managerText -match [regex]::Escape("`$updateHelp.Text = 'Use Check for updates / install")) {
    Add-Failure 'Information & Help has regressed to the old run-on update paragraph.'
}
foreach ($labelName in @('updateExplanation', 'updateStep1', 'updateStep2', 'updateStep3',
        'updateKeepsData', 'lblNormal', 'lblLong', 'lblSync', 'lblMon')) {
    if ($managerText -notmatch ('\$' + $labelName + '\s*=\s*New-FormWrappedLabel')) {
        Add-Failure ('Information & Help paragraph is not adaptively wrapped: {0}' -f $labelName)
    }
}
foreach ($requiredText in @('Find-DuplicateRuleDefinition -Candidate $rule',
        'Find-DuplicateRuleDefinition -Candidate $edited', "'Duplicate rule'")) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('The full rule editor save path is missing duplicate protection: {0}' -f $requiredText)
    }
}

$configManagerText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\ConfigManager.ps1'))
foreach ($requiredText in @('[string]::IsNullOrWhiteSpace($ruleId)',
        'Ignored {0} invalid saved rule-state', 'function Get-RuleDefinitionSignature',
        'function Find-DuplicateRuleDefinition', 'askBeforeRefresh', 'selectedQueries')) {
    if ($configManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Saved-state recovery is missing required behavior: {0}' -f $requiredText)
    }
}

$engineText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Engine.ps1'))
if ($engineText -notmatch [regex]::Escape('during {1}: {2}')) {
    Add-Failure 'Engine startup errors do not identify their startup phase.'
}

$updateText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIUpdate.ps1'))
if ($updateText -match '(?i)Invoke-RestMethod[^\r\n]*-Method\s+(Post|Patch|Put|Delete)') {
    Add-Failure 'src\UIUpdate.ps1 contains a write request.'
}
if ($updateText -notmatch "UpdateRepository\s*=\s*'quietworktools/ExcelQueryTrigger'") {
    Add-Failure 'The fixed public update repository is missing.'
}
foreach ($requiredText in @('checkForUpdatesAutomatically',
        'Stop-BackgroundUpdateCheck', 'if (-not $automatic)')) {
    if ($updateText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Background update checking does not honor its Settings switch: {0}' -f $requiredText)
    }
}
if ($wizardText -match [regex]::Escape('$script:UiConfig.rules')) {
    Add-Failure 'The Add file click closure still accesses script-scoped UiConfig directly.'
}

foreach ($requiredText in @('Test-WorkbookHasQueries -Path $Path',
        'Nothing to refresh', 'queries/connections will refresh ({0} found)',
        'Find-DuplicateWorkbookAction -Actions @($actions.ToArray())',
        'Find-DuplicateWorkbookAction -Actions @($workbookState.Actions)')) {
    $combinedUiText = $ruleEditorText + [Environment]::NewLine + $wizardText
    if ($combinedUiText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Workbook Add validation is missing required behavior: {0}' -f $requiredText)
    }
}

$diagnosticsText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIDiagnostics.ps1'))
foreach ($requiredText in @('$form.Font            = Get-UiFont',
        '$form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::None',
        'Set-FormWithinWorkingArea -Form $form')) {
    if ($diagnosticsText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Test data sources is missing the common UI font/layout: {0}' -f $requiredText)
    }
}
if ($diagnosticsText -match [regex]::Escape('$script:UiFonts.Normal')) {
    Add-Failure 'Test data sources still references the nonexistent UiFonts.Normal key.'
}

# Reproduce the Add button's closure boundary without opening WinForms. The
# live configuration must be mutated through the captured reference, not a
# script-scoped variable owned by GetNewClosure's dynamic module.
$wizardReferenceState = @{ Config = @{ rules = @() } }
$wizardWriter = {
    $liveConfig = $wizardReferenceState.Config
    $liveConfig['rules'] = @('added')
}.GetNewClosure()
& $wizardWriter
if (@($wizardReferenceState.Config['rules']).Count -ne 1 -or
    [string]$wizardReferenceState.Config['rules'][0] -ne 'added') {
    Add-Failure 'The Add wizard configuration reference is not shared across its click closure.'
}
foreach ($requiredText in @('Invoke-RestMethod @proxyArgs', 'Invoke-WebRequest @proxyArgs',
        'Invoke-RestMethod @RequestProxyArgs', '.AddArgument($proxyArgs)')) {
    if ($updateText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('An update network path is missing proxy support: {0}' -f $requiredText)
    }
}
if ($updateText -match [regex]::Escape('@(Get-UpdateProxyArgs)')) {
    Add-Failure 'Update proxy arguments use an array expression instead of PowerShell splatting.'
}

foreach ($requiredText in @('Complete-StartupWatcherInitialization',
        'Waiting for file and folder monitors', 'StartupMonitoringReady = $true',
        'StartupUiReady', 'StartupReady = $true',
        'Sync-RuleWatchers -Rules $rules -Paused $false -QuietUnavailable')) {
    $engineTextForStartup = [System.IO.File]::ReadAllText((Join-Path $Root 'src\\Engine.ps1'))
    if ($engineTextForStartup -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Startup watcher readiness gate is missing: {0}' -f $requiredText)
    }
}

foreach ($requiredText in @('New-QueryProgressTracker', 'Update-QueryProgressTracker',
        'Observed query timings', 'Starting RefreshAll',
        'QueryProgressDetail', 'RefreshConfirmationSeconds',
        'Saving workbook directly')) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText) -and
        $managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Query progress reporting is missing: {0}' -f $requiredText)
    }
}

# Startup must not intentionally fall through to a Degraded Dashboard after
# an arbitrary retry budget. Monitoring readiness is an actual condition.
$engineTextForStartup = [System.IO.File]::ReadAllText((Join-Path $Root 'src\Engine.ps1'))
if ($engineTextForStartup -match 'MaxAttempts\s*=\s*30' -or
    $engineTextForStartup -match 'Monitoring initialized with warnings') {
    Add-Failure 'Startup still contains the retired bounded watcher fallback.'
}
if ($engineTextForStartup -notmatch [regex]::Escape('[string]$shared.Status -eq ''Running''')) {
    Add-Failure 'Startup does not gate readiness on the same Running status used by the Dashboard.'
}

# Query timing must stay a flat collection. A nested Object[] caused TND Tracking
# to fail when the log formatter cast `$_.seconds` to Int32 for several queries.
$timingShared = New-SharedState
$timingTracker = @{
    StartedAt = (Get-Date).AddSeconds(-5)
    Items = @(
        @{ Name='A'; ConnectionName='A'; State='Done'; StartedAt=(Get-Date).AddSeconds(-5); FinishedAt=(Get-Date); DurationSeconds=5; ObservedActive=$true },
        @{ Name='B'; ConnectionName='B'; State='Done'; StartedAt=(Get-Date).AddSeconds(-3); FinishedAt=(Get-Date); DurationSeconds=3; ObservedActive=$true }
    )
}
$timingResult = @(Complete-QueryProgressTracker -Tracker $timingTracker -Shared $timingShared)
if ($timingResult.Count -ne 2 -or
    $timingResult[0] -is [System.Array] -or
    $timingResult[1] -is [System.Array] -or
    [int]$timingResult[0].seconds -ne 5 -or
    [int]$timingResult[1].seconds -ne 3) {
    Add-Failure 'Query timing results are nested or cannot represent multiple scalar timings.'
}

# Safe temporary-file saving is gone: the default action must not carry the
# retired flag, and the rule editor must not offer it any more.
$safeDefault = Get-DefaultAction
if ($safeDefault.ContainsKey('safeSave')) {
    Add-Failure 'The retired safeSave action flag is still produced by Get-DefaultAction.'
}
$ruleEditorText = [System.IO.File]::ReadAllText((Join-Path $Root 'src\UIRuleEditor.ps1'))
if ($ruleEditorText -match [regex]::Escape('safeSave')) {
    Add-Failure 'The retired safe-save option is still present in the Excel Action dialog.'
}

# Every save mode must refuse to start Excel when the target workbook is
# already owned elsewhere. This preflight happens before any query work.
foreach ($requiredText in @(
        'CheckingWorkbook',
        'Refresh was not started because the workbook is already in use',
        'Workbook opened for writing. Excel will keep it open through refresh, save and close.',
        'No queries were refreshed',
        "$notificationTitle = 'Refresh not started'")) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText) -and
        $jobManagerText -notmatch [regex]::Escape($requiredText) -and
        $commonText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Workbook preflight protection is missing: {0}' -f $requiredText)
    }
}

# Excel's writable open owns the workbook for the whole refresh. A second
# size/timestamp guard cannot identify the writer and used to reject legitimate
# refreshes when Excel or the storage layer changed the file during RefreshAll.
foreach ($retiredText in @('Get-WorkbookFileSnapshot', 'WorkbookChangedExternally', 'checkExternalChange')) {
    if ($excelManagerText -match [regex]::Escape($retiredText)) {
        Add-Failure ('The retired file-metadata save guard is still present: {0}' -f $retiredText)
    }
}

# Trigger Rules can be reordered directly in the dashboard. The drag data is a
# stable rule id, and the resulting array is persisted through the normal
# configuration save/reload path.
foreach ($requiredText in @(
        '$lvRules.AllowDrop     = $true',
        'Add_ItemDrag',
        'Add_DragOver',
        'Add_DragDrop',
        'InsertionMark.AppearsAfterItem',
        'function Move-TriggerRule',
        '$finalRules = @($reordered.ToArray())',
        '$script:UiConfig.rules = $finalRules',
        'if (Save-UiConfiguration)')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Trigger Rules drag reorder support is missing: {0}' -f $requiredText)
    }
}

# Every Trigger Rules heading sorts the displayed rows. Text fields sort by
# normalized text while date/time and duration fields retain numeric keys.
foreach ($requiredText in @(
        'Add_ColumnClick',
        'function Update-RuleColumnHeaders',
        '$script:UiRuleSortDirection',
        "' ▼'",
        "' ▲'",
        'SortKeys = $sortKeys',
        'Sort-Object -Property $primarySort, $stableSort',
        'Get-RuleNextRunTime',
        'DataSortTicks',
        '[double]$lastSeconds')) {
    if ($managerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('Trigger Rules column sorting is missing: {0}' -f $requiredText)
    }
}

# Exercise the authoritative Excel owner-file signal without requiring Excel.
$lockTestDir = Join-Path ([System.IO.Path]::GetTempPath()) ('eqt-lock-test-{0}' -f [Guid]::NewGuid().ToString('N'))
try {
    [void](New-Item -ItemType Directory -Path $lockTestDir -Force)
    $lockWorkbook = Join-Path $lockTestDir 'Shared.xlsx'
    [System.IO.File]::WriteAllBytes($lockWorkbook, [byte[]](1,2,3))
    $lockOwner = Join-Path $lockTestDir ('~$' + 'Shared.xlsx')
    [System.IO.File]::WriteAllBytes($lockOwner, [byte[]](1,0,0,0))
    $heldLock = $null
    try {
        $heldLock = [System.IO.File]::Open($lockWorkbook, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $lockResult = Test-WorkbookLocked -Path $lockWorkbook -Attempts 1 -DelayMilliseconds 0
        if (-not [bool]$lockResult.Locked -or [string]$lockResult.Reason -ne 'ExcelOwnerFile') {
            Add-Failure 'An Excel owner file plus an active file handle did not block workbook refresh preflight.'
        }
    }
    finally {
        if ($null -ne $heldLock) { $heldLock.Dispose() }
    }
    Remove-Item -LiteralPath $lockOwner -Force -ErrorAction SilentlyContinue
    [System.IO.File]::WriteAllBytes($lockOwner, [byte[]](1,0,0,0))
    $staleOwnerResult = Test-WorkbookLocked -Path $lockWorkbook -Attempts 1 -DelayMilliseconds 0
    if ([bool]$staleOwnerResult.Locked -or [string]$staleOwnerResult.Reason -ne 'StaleExcelOwnerFile') {
        Add-Failure 'A stale Excel owner file incorrectly blocked an otherwise available workbook.'
    }
}
finally {
    Remove-Item -LiteralPath $lockTestDir -Recurse -Force -ErrorAction SilentlyContinue
}

# BackgroundQuery restore safety: only settings actually changed are recorded,
# and restore is resilient to connection collection reordering/transient COM.
foreach ($requiredText in @('Name = $connectionName', 'if (-not $original) { continue }',
        '$connections.Item($name)', 'MaxAttempts = 3', 'FailedNames')) {
    if ($excelManagerText -notmatch [regex]::Escape($requiredText)) {
        Add-Failure ('BackgroundQuery restoration hardening is missing: {0}' -f $requiredText)
    }
}

if ($failures.Count -gt 0) {
    Write-Host ('{0} check(s) failed.' -f $failures.Count) -ForegroundColor Red
    exit 1
}

Write-Host 'PASS  PowerShell syntax, package layout, schema, retired settings and read-only update checks.' -ForegroundColor Green
