# ==============================================================================
#  LogManager.ps1
#  Two sinks, one API:
#    1. rolling text file   logs\app-yyyyMMdd.log      (persistent)
#    2. shared activity list                            (GUI "Recent Activity")
#  Plus a JSON-lines job history used by the log detail window.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:LogDirectory   = $null
$script:HistoryPath    = $null
$script:LogShared      = $null
$script:LogRetention   = 30
$script:LogMaxBytes    = 50MB
$script:LogDebug       = $false
$script:LogSourceTag   = 'app'
$script:ActivityMax    = 400
$script:LastRetentionRun = [DateTime]::MinValue

function Initialize-LogManager {
    param(
        [Parameter(Mandatory = $true)][string]$LogDirectory,
        [string]$HistoryPath,
        [hashtable]$Shared,
        [int]$RetentionDays = 30,
        [int]$MaxTotalMegabytes = 50,
        [bool]$DebugLogging = $false,
        [string]$SourceTag = 'app'
    )

    $script:LogDirectory = $LogDirectory
    $script:HistoryPath  = $HistoryPath
    $script:LogShared    = $Shared
    $script:LogRetention = $RetentionDays
    $script:LogMaxBytes  = [int64]$MaxTotalMegabytes * 1MB
    $script:LogDebug     = $DebugLogging
    $script:LogSourceTag = $SourceTag

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Invoke-LogRetention
}

function Set-LogDebugMode {
    param([bool]$Enabled)
    $script:LogDebug = $Enabled
    if ($null -ne $script:LogShared) { $script:LogShared.DebugLogging = $Enabled }
}

function Get-CurrentLogFilePath {
    if ([string]::IsNullOrWhiteSpace($script:LogDirectory)) { return $null }
    return (Join-Path $script:LogDirectory ('app-{0}.log' -f (Get-Date -Format 'yyyyMMdd')))
}

function Write-AppLog {
    <#
        The single logging entry point. Safe to call from any thread; file
        writes are retried because the UI runspace and the engine runspace can
        touch the same file within the same millisecond.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO',
        [string]$RuleName,
        [string]$Workbook,
        [string]$Stage,
        [string]$ErrorType,
        [switch]$NoActivity
    )

    if ($Level -eq 'DEBUG' -and -not $script:LogDebug) { return }

    $timestamp = Get-Date
    $parts = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($RuleName))  { [void]$parts.Add(('Rule={0}' -f $RuleName)) }
    if (-not [string]::IsNullOrWhiteSpace($Workbook))  { [void]$parts.Add(('Workbook={0}' -f (Split-Path -Leaf $Workbook))) }
    if (-not [string]::IsNullOrWhiteSpace($Stage))     { [void]$parts.Add(('Stage={0}' -f $Stage)) }
    if (-not [string]::IsNullOrWhiteSpace($ErrorType)) { [void]$parts.Add(('ErrorType={0}' -f $ErrorType)) }
    [void]$parts.Add($Message)

    $line = '{0} [{1}] {2}' -f $timestamp.ToString('yyyy-MM-dd HH:mm:ss'), $Level, ($parts -join ' ')

    # ---- sink 1: text file -----------------------------------------------
    $path = Get-CurrentLogFilePath
    if ($null -ne $path) {
        for ($attempt = 1; $attempt -le 4; $attempt++) {
            try {
                [System.IO.File]::AppendAllText($path, ($line + [Environment]::NewLine), [System.Text.Encoding]::UTF8)
                break
            }
            catch {
                Start-Sleep -Milliseconds (25 * $attempt)
            }
        }
    }

    # ---- sink 2: GUI activity feed ---------------------------------------
    if (-not $NoActivity -and $null -ne $script:LogShared) {
        $activityLocked = $false
        try {
            # Both runspaces log, so the sequence number and the append have to
            # happen together or the dashboard could skip an entry.
            [System.Threading.Monitor]::Enter($script:LogShared.Activity.SyncRoot, [ref]$activityLocked)
            $script:LogShared.LogSeq = [int]$script:LogShared.LogSeq + 1
            $entry = @{
                Seq       = $script:LogShared.LogSeq
                Time      = $timestamp
                Level     = $Level
                Message   = $Message
                RuleName  = $RuleName
                Workbook  = $Workbook
                Stage     = $Stage
                ErrorType = $ErrorType
                Line      = $line
            }
            [void]$script:LogShared.Activity.Add($entry)
            while ($script:LogShared.Activity.Count -gt $script:ActivityMax) {
                $script:LogShared.Activity.RemoveAt(0)
            }
        }
        catch {
            # Never let logging break the caller.
        }
        finally {
            if ($activityLocked) { [System.Threading.Monitor]::Exit($script:LogShared.Activity.SyncRoot) }
        }
    }

    # Retention is cheap to check and must not depend on the app being restarted.
    if (((Get-Date) - $script:LastRetentionRun).TotalHours -ge 6) { Invoke-LogRetention }
}

function Write-JobHistory {
    <#  Appends one job record as JSON-lines and keeps a copy in memory.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Record
    )

    if ($null -ne $script:LogShared) {
        try {
            [void]$script:LogShared.History.Add($Record)
            while ($script:LogShared.History.Count -gt 200) { $script:LogShared.History.RemoveAt(0) }
        }
        catch { }
    }

    if ([string]::IsNullOrWhiteSpace($script:HistoryPath)) { return }
    try {
        $json = ($Record | ConvertTo-Json -Depth 6 -Compress)
        [System.IO.File]::AppendAllText($script:HistoryPath, ($json + [Environment]::NewLine), [System.Text.Encoding]::UTF8)
    }
    catch { }
}

function Invoke-LogRetention {
    <#
        Two independent guards: age based, then size based.
        Size trimming deletes oldest files first until under the cap.
    #>
    $script:LastRetentionRun = Get-Date
    if ([string]::IsNullOrWhiteSpace($script:LogDirectory)) { return }
    if (-not (Test-Path -LiteralPath $script:LogDirectory)) { return }

    try {
        $cutoff = (Get-Date).AddDays(-1 * $script:LogRetention)
        $files = @(Get-ChildItem -LiteralPath $script:LogDirectory -Filter 'app-*.log' -File -ErrorAction SilentlyContinue)

        foreach ($file in $files) {
            if ($file.LastWriteTime -lt $cutoff) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            }
        }

        $remaining = @(Get-ChildItem -LiteralPath $script:LogDirectory -Filter 'app-*.log' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime)
        $total = 0
        foreach ($file in $remaining) { $total += $file.Length }

        $index = 0
        while ($total -gt $script:LogMaxBytes -and $index -lt ($remaining.Count - 1)) {
            $total -= $remaining[$index].Length
            Remove-Item -LiteralPath $remaining[$index].FullName -Force -ErrorAction SilentlyContinue
            $index++
        }

        # The history file is trimmed by truncating to the newest 5000 lines.
        if (-not [string]::IsNullOrWhiteSpace($script:HistoryPath) -and (Test-Path -LiteralPath $script:HistoryPath)) {
            $historyFile = Get-Item -LiteralPath $script:HistoryPath -ErrorAction SilentlyContinue
            if ($null -ne $historyFile -and $historyFile.Length -gt 8MB) {
                $lines = @(Get-Content -LiteralPath $script:HistoryPath -Tail 5000 -ErrorAction SilentlyContinue)
                Set-Content -LiteralPath $script:HistoryPath -Value $lines -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
    }
    catch { }
}
