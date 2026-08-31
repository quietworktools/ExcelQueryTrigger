# ==============================================================================
#  ExcelManager.ps1  (Engine runspace only - the runspace is STA for COM)
#
#  The whole point of this file is that "RefreshAll + Start-Sleep 30 + Save" is
#  not acceptable. The refresh lifecycle here is:
#     create instance -> open -> force foreground queries -> RefreshAll
#     -> poll every connection / query table until quiet -> confirm with
#     CalculateUntilAsyncQueriesDone -> save -> close -> release -> verify the
#     process we started is really gone.
# ==============================================================================

Set-StrictMode -Version 1.0

function Remove-ComReference {
    <#  Releases an RCW without throwing. Always call, even on the error path. #>
    param([Parameter(ValueFromRemainingArguments = $true)]$ComObjects)
    foreach ($comObject in $ComObjects) {
        if ($null -eq $comObject) { continue }
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
        catch { }
    }
}

function Clear-ComMemory {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Test-WorkbookLocked {
    <#
        Excel usually advertises an open workbook through its owner file
        "~$<name>". Some network/provider combinations expose only a sharing
        violation, so check both signals. A transient SMB/antivirus handle must
        not cancel a refresh, therefore sharing violations are confirmed across
        a few short retries before we report the workbook as in use.

        The owner file is not trusted by itself because Office can leave a stale
        "~$" file after a crash. When the owner file exists but an exclusive
        read/write probe succeeds, the workbook is treated as available.

        This check is deliberately performed BEFORE Excel is started. If another
        user already has the workbook open, no query refresh is attempted.

        Returns @{ Locked; Owner; Reason; OwnerFile }
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$Attempts = 3,
        [int]$DelayMilliseconds = 300
    )

    $Attempts = [Math]::Max(1, $Attempts)
    $result = @{ Locked = $false; Owner = ''; Reason = ''; OwnerFile = '' }
    $directory = Split-Path -Parent $Path
    $leaf      = Split-Path -Leaf $Path
    $ownerFile = Join-Path $directory ('~$' + $leaf)
    $result.OwnerFile = $ownerFile

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $ownerPresent = $false
        $ownerCandidate = ''
        try {
            $ownerPresent = Test-Path -LiteralPath $ownerFile
            if ($ownerPresent) {
                try {
                    # Owner-file encodings vary between Office generations and
                    # storage providers. Parse only when the common UTF-16 form
                    # is recognizable; an empty Owner still leaves the sharing
                    # probe as the authoritative lock check.
                    $bytes = [System.IO.File]::ReadAllBytes($ownerFile)
                    if ($bytes.Length -gt 2) {
                        $length = [int]$bytes[0]
                        if ($length -gt 0 -and ($length * 2 + 2) -le $bytes.Length) {
                            $ownerCandidate = [System.Text.Encoding]::Unicode.GetString($bytes, 2, ($length - 1) * 2).Trim([char]0)
                        }
                    }
                }
                catch { }
            }
        }
        catch { $ownerPresent = $false }

        $stream = $null
        $sharingViolation = $false
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            $sharingViolation = $true
        }
        catch {
            # Permission/path problems are handled by Excel/open-path errors;
            # they are not automatically classified as "another user".
            $sharingViolation = $false
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
        }

        if (-not $sharingViolation) {
            if ($ownerPresent) { $result.Reason = 'StaleExcelOwnerFile' }
            return $result
        }

        if ($ownerPresent) {
            $result.Locked = $true
            $result.Owner  = $ownerCandidate
            $result.Reason = 'ExcelOwnerFile'
            return $result
        }

        if ($attempt -lt $Attempts) {
            Start-Sleep -Milliseconds ([Math]::Max(0, $DelayMilliseconds))
            continue
        }

        $result.Locked = $true
        $result.Reason = 'SharingViolation'
        return $result
    }

    return $result
}

function Set-WorkbookConnectionsForeground {
    <#
        Temporarily sets BackgroundQuery=False for compatible OLEDB/ODBC
        connections. Only a setting that was actually changed is recorded for
        restoration; writing the original value back to an untouched connection
        creates avoidable COM failures on some Data Model/cube connections.
    #>
    param([Parameter(Mandatory = $true)]$Workbook)

    $states = New-Object System.Collections.ArrayList
    $changed = 0
    $connections = $null
    try { $connections = $Workbook.Connections } catch { return @{ ChangedCount = 0; States = @() } }
    if ($null -eq $connections) { return @{ ChangedCount = 0; States = @() } }

    try {
        for ($index = 1; $index -le $connections.Count; $index++) {
            $connection = $null
            $inner = $null
            try {
                $connection = $connections.Item($index)
                $type = 0
                try { $type = [int]$connection.Type } catch { $type = 0 }
                if ($type -eq 1) { $inner = $connection.OLEDBConnection }
                elseif ($type -eq 2) { $inner = $connection.ODBCConnection }
                if ($null -eq $inner) { continue }

                $original = [bool]$inner.BackgroundQuery
                if (-not $original) { continue }

                # Add the restore state only after Excel confirms the temporary
                # change. A failed assignment means there is nothing to restore.
                $inner.BackgroundQuery = $false
                $connectionName = ''
                try { $connectionName = [string]$connection.Name } catch { $connectionName = '' }
                [void]$states.Add(@{
                    Name = $connectionName
                    Index = $index
                    Type = $type
                    BackgroundQuery = $true
                })
                $changed++
            }
            catch { }
            finally { Remove-ComReference $inner $connection }
        }
    }
    catch { }
    finally { Remove-ComReference $connections }

    return @{ ChangedCount = $changed; States = @($states.ToArray()) }
}

function Restore-WorkbookConnectionBackgroundQuery {
    <#
        Restores values captured by Set-WorkbookConnectionsForeground.
        Connection name is preferred over collection index because RefreshAll
        can rebuild/reorder connection objects. Transient COM rejection just
        after refresh is retried a few times. A real restore failure remains a
        hard save barrier so a temporary tool setting is never persisted.
    #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        $States,
        [int]$MaxAttempts = 3,
        [int]$RetryDelayMilliseconds = 300
    )

    $restored = 0
    $failed = 0
    $failedNames = New-Object System.Collections.ArrayList
    if ($null -eq $States) { return @{ Restored = 0; Failed = 0; FailedNames = @() } }

    foreach ($state in @($States)) {
        $success = $false
        $displayName = [string]$state['Name']
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = 'connection #{0}' -f (ConvertTo-IntValue $state['Index'] 0 1)
        }

        for ($attempt = 1; $attempt -le [Math]::Max(1, $MaxAttempts); $attempt++) {
            $connections = $null
            $connection = $null
            $inner = $null
            try {
                $connections = $Workbook.Connections
                if ($null -eq $connections) { throw 'Workbook.Connections is unavailable.' }

                $name = [string]$state['Name']
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    try { $connection = $connections.Item($name) } catch { $connection = $null }
                }
                if ($null -eq $connection) {
                    $index = ConvertTo-IntValue $state['Index'] 0 1
                    if ($index -le 0 -or $index -gt $connections.Count) { throw 'Saved connection is no longer present.' }
                    $connection = $connections.Item($index)
                }

                $type = ConvertTo-IntValue $state['Type'] 0 0
                if ($type -eq 1) { $inner = $connection.OLEDBConnection }
                elseif ($type -eq 2) { $inner = $connection.ODBCConnection }
                if ($null -eq $inner) { throw 'Connection no longer exposes BackgroundQuery.' }

                $expected = [bool]$state['BackgroundQuery']
                $current = [bool]$inner.BackgroundQuery
                if ($current -ne $expected) {
                    $inner.BackgroundQuery = $expected
                    $current = [bool]$inner.BackgroundQuery
                }
                if ($current -ne $expected) { throw 'BackgroundQuery did not retain the restored value.' }

                $success = $true
                break
            }
            catch { }
            finally { Remove-ComReference $inner $connection $connections }

            if (-not $success -and $attempt -lt [Math]::Max(1, $MaxAttempts)) {
                Start-Sleep -Milliseconds ([Math]::Max(50, $RetryDelayMilliseconds))
            }
        }

        if ($success) { $restored++ }
        else {
            $failed++
            [void]$failedNames.Add($displayName)
        }
    }

    return @{ Restored = $restored; Failed = $failed; FailedNames = @($failedNames.ToArray()) }
}

function Test-WorkbookConnectionTargets {
    <#
        Checks every external connection before the refresh starts. A data source
        whose server no longer resolves is the usual reason the provider raises
        its own credential wizard, and that dialog blocks Excel inside a COM call
        where cancellation cannot reach it. Failing here instead gives a clear
        message and leaves nothing running.

        Returns @{ Blocked = @(...); Warnings = @(...); Checked = n }
    #>
    param([Parameter(Mandatory = $true)]$Workbook)

    $blocked  = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    $skipped  = New-Object System.Collections.ArrayList
    $checked  = 0
    $seen     = @{}

    $connections = $null
    try { $connections = $Workbook.Connections } catch { return @{ Blocked = @(); Warnings = @(); Skipped = @(); Checked = 0 } }
    if ($null -eq $connections) { return @{ Blocked = @(); Warnings = @(); Skipped = @(); Checked = 0 } }

    try {
        for ($index = 1; $index -le $connections.Count; $index++) {
            $connection = $null
            $inner      = $null
            try {
                $connection = $connections.Item($index)
                $name = ''
                try { $name = [string]$connection.Name } catch { $name = '' }
                if ([string]::IsNullOrWhiteSpace($name)) { $name = ('connection {0}' -f $index) }

                $type = 0
                try { $type = [int]$connection.Type } catch { $type = 0 }
                if ($type -eq 1) { $inner = $connection.OLEDBConnection }
                elseif ($type -eq 2) { $inner = $connection.ODBCConnection }
                if ($null -eq $inner) { continue }

                $connectionString = ''
                try { $connectionString = [string]$inner.Connection } catch { $connectionString = '' }
                $target = Get-ConnectionServerTarget -ConnectionString $connectionString
                if ($null -eq $target) {
                    [void]$skipped.Add($name)
                    continue
                }

                $key = ('{0}|{1}' -f ([string]$target.Host).ToLowerInvariant(), $target.Port)
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true
                $checked++

                $probe = Test-ServerReachable -ServerHost $target.Host -Port $target.Port -DefaultPort $target.DefaultPort
                if (-not $probe.Resolved -and $probe.Conclusive) {
                    [void]$blocked.Add(('"{0}" uses the server {1}, which does not exist on the network.' -f $name, $target.Host))
                }
                elseif (-not $probe.Resolved) {
                    # Name resolution itself failed rather than answering "no such
                    # host". Say so and let Excel try: a resolver problem here must
                    # not stop a workbook that would refresh perfectly well.
                    [void]$warnings.Add(('"{0}": could not look up {1} right now. The refresh will still be attempted.' -f $name, $target.Host))
                }
                elseif ($target.Port -gt 0 -and $probe.Probed -and -not $probe.PortOpen) {
                    [void]$blocked.Add(('"{0}" uses {1} on port {2}, which is not accepting connections.' -f $name, $target.Host, $target.Port))
                }
                elseif ($probe.Probed -and -not $probe.PortOpen) {
                    [void]$warnings.Add(('"{0}": {1} was found but did not answer on the usual port. The refresh will still be attempted.' -f $name, $target.Host))
                }
            }
            catch { }
            finally { Remove-ComReference $inner $connection }
        }
    }
    catch { }
    finally { Remove-ComReference $connections }

    return @{ Blocked = @($blocked.ToArray()); Warnings = @($warnings.ToArray()); Skipped = @($skipped.ToArray()); Checked = $checked }
}

function Get-WorkbookRefreshState {
    <#
        Returns Busy, Quiet or Unknown. A top-level COM failure is never treated
        as completion: saving after an RPC rejection is less safe than failing
        the job and leaving the on-disk workbook unchanged.
    #>
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Workbook
    )

    # --- workbook connections ---------------------------------------------
    $stateUnknown = $false
    $connections = $null
    try {
        $connections = $Workbook.Connections
        if ($null -eq $connections) { $stateUnknown = $true }
    }
    catch { $connections = $null; $stateUnknown = $true }
    if ($null -ne $connections) {
        try {
            for ($index = 1; $index -le $connections.Count; $index++) {
                $connection = $null
                $inner      = $null
                try {
                    $connection = $connections.Item($index)
                    $type = 0
                    try { $type = [int]$connection.Type } catch { $type = 0 }
                    if ($type -eq 1) { $inner = $connection.OLEDBConnection }
                    elseif ($type -eq 2) { $inner = $connection.ODBCConnection }
                    if ($null -ne $inner -and $inner.Refreshing) { return 'Busy' }
                }
                catch { }
                finally { Remove-ComReference $inner $connection }
            }
        }
        catch { $stateUnknown = $true }
        finally { Remove-ComReference $connections }
    }

    # --- query tables and list objects on every sheet ----------------------
    $sheets = $null
    try {
        $sheets = $Workbook.Worksheets
        if ($null -eq $sheets) { $stateUnknown = $true }
    }
    catch { $sheets = $null; $stateUnknown = $true }
    if ($null -ne $sheets) {
        try {
            for ($sheetIndex = 1; $sheetIndex -le $sheets.Count; $sheetIndex++) {
                $sheet = $null
                try {
                    $sheet = $sheets.Item($sheetIndex)

                    $queryTables = $null
                    try { $queryTables = $sheet.QueryTables } catch { $queryTables = $null }
                    if ($null -ne $queryTables) {
                        try {
                            for ($q = 1; $q -le $queryTables.Count; $q++) {
                                $queryTable = $null
                                try {
                                    $queryTable = $queryTables.Item($q)
                                    if ($queryTable.Refreshing) { return 'Busy' }
                                }
                                catch { }
                                finally { Remove-ComReference $queryTable }
                            }
                        }
                        catch { }
                        finally { Remove-ComReference $queryTables }
                    }

                    $listObjects = $null
                    try { $listObjects = $sheet.ListObjects } catch { $listObjects = $null }
                    if ($null -ne $listObjects) {
                        try {
                            for ($l = 1; $l -le $listObjects.Count; $l++) {
                                $listObject = $null
                                $listQuery  = $null
                                try {
                                    $listObject = $listObjects.Item($l)
                                    $listQuery  = $listObject.QueryTable
                                    if ($null -ne $listQuery -and $listQuery.Refreshing) { return 'Busy' }
                                }
                                catch { }
                                finally { Remove-ComReference $listQuery $listObject }
                            }
                        }
                        catch { }
                        finally { Remove-ComReference $listObjects }
                    }
                }
                catch { }
                finally { Remove-ComReference $sheet }
            }
        }
        catch { $stateUnknown = $true }
        finally { Remove-ComReference $sheets }
    }

    # --- calculation engine ------------------------------------------------
    try {
        # 0 = xlDone, 1 = xlCalculating, 2 = xlPending
        if ([int]$Excel.CalculationState -ne 0) { return 'Busy' }
    }
    catch { $stateUnknown = $true }

    if ($stateUnknown) { return 'Unknown' }
    return 'Quiet'
}

function Test-WorkbookRefreshBusy {
    <# Compatibility wrapper for older diagnostic callers. Unknown is busy. #>
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Workbook
    )
    return ((Get-WorkbookRefreshState -Excel $Excel -Workbook $Workbook) -ne 'Quiet')
}

function Wait-ExcelRefreshCompletion {
    <#
        Polls at a human pace (no busy loop), requires several consecutive quiet
        samples, and only then blocks once on CalculateUntilAsyncQueriesDone as
        a final confirmation. The configured threshold is a warning only; cancellation is still honored.
    #>
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Workbook,
        [int]$TimeoutSeconds = 300,
        [int]$PollMilliseconds = 750,
        [int]$RequiredQuietSamples = 3,
        [scriptblock]$ShouldAbort,
        [scriptblock]$OnPoll,
        [scriptblock]$OnWarning
    )

    $started        = Get-Date
    $warningAt      = $started.AddSeconds($TimeoutSeconds)
    $warningRaised  = $false
    $quietSamples   = 0
    $asyncConfirmed = $false
    $unknownSamples = 0

    while ($true) {
        if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
            throw (New-Object System.OperationCanceledException('Refresh cancelled by the application.'))
        }

        if (-not $warningRaised -and (Get-Date) -gt $warningAt) {
            $warningRaised = $true
            if ($null -ne $OnWarning) {
                try { & $OnWarning $TimeoutSeconds | Out-Null } catch { }
            }
        }

        $refreshState = Get-WorkbookRefreshState -Excel $Excel -Workbook $Workbook

        if ($null -ne $OnPoll) {
            try { & $OnPoll $refreshState ([int]((Get-Date) - $started).TotalSeconds) | Out-Null } catch { }
        }

        if ($refreshState -eq 'Busy') {
            $quietSamples = 0
            $unknownSamples = 0
        }
        elseif ($refreshState -eq 'Unknown') {
            $quietSamples = 0
            $unknownSamples++
            if ($unknownSamples -ge 8) {
                throw ([System.InvalidOperationException]::new('Excel refresh state could not be verified after repeated COM errors. The workbook will not be saved.'))
            }
        }
        else {
            $unknownSamples = 0
            if (-not $asyncConfirmed) {
                # Blocking call: returns when Excel says every async query is done.
                try { $Excel.CalculateUntilAsyncQueriesDone() }
                catch {
                    throw ([System.InvalidOperationException]::new(
                        ('Excel could not confirm that asynchronous queries finished. The workbook will not be saved: {0}' -f $_.Exception.Message), $_.Exception))
                }
                $asyncConfirmed = $true
                continue
            }
            $quietSamples++
            if ($quietSamples -ge $RequiredQuietSamples) { break }
        }

        Start-Sleep -Milliseconds $PollMilliseconds
    }

    return [int]((Get-Date) - $started).TotalSeconds
}

function Get-WorkbookPowerQueryInfo {
    <#
        Everything in this workbook that can be refreshed on its own.

        A Power Query query is paired with its workbook connection through the
        Location= value in that connection's own string, not by guessing at the
        connection's name. Excel localises the name ("Query - Orderdata",
        "クエリ - Orderdata"), so name matching works on an English build and
        silently fails everywhere else - which is how a query could end up
        listed twice, once under its own name and once under its connection's.

        Whatever is left over - a cube, ODBC, OLEDB or legacy MS Query
        connection - is listed under its own name, because those are refreshable
        on their own but never appear in Workbook.Queries.
    #>
    param([Parameter(Mandatory = $true)]$Workbook)

    $connections = @(Get-WorkbookConnectionSummary -Workbook $Workbook -Release { param($a, $b) Remove-ComReference $a $b })

    # query name (lowercased) -> connection name
    $connectionForQuery = @{}
    foreach ($connection in $connections) {
        $location = [string]$connection.QueryLocation
        if (-not [string]::IsNullOrWhiteSpace($location)) {
            $connectionForQuery[$location.ToLowerInvariant()] = [string]$connection.Name
        }
    }

    $items = New-Object System.Collections.ArrayList
    $claimed = @{}

    $queries = $null
    try { $queries = $Workbook.Queries } catch { $queries = $null }
    if ($null -ne $queries) {
        try {
            for ($i = 1; $i -le $queries.Count; $i++) {
                $query = $null
                try {
                    $query = $queries.Item($i)
                    $queryName = [string]$query.Name
                    if ([string]::IsNullOrWhiteSpace($queryName)) { continue }

                    $key = $queryName.ToLowerInvariant()
                    $connectionName = ''
                    if ($connectionForQuery.ContainsKey($key)) { $connectionName = [string]$connectionForQuery[$key] }

                    if (-not [string]::IsNullOrWhiteSpace($connectionName)) { $claimed[$connectionName.ToLowerInvariant()] = $true }
                    [void]$items.Add(@{
                        Name = $queryName
                        ConnectionName = $connectionName
                        Kind = 'Power Query'
                        Refreshable = (-not [string]::IsNullOrWhiteSpace($connectionName))
                    })
                }
                catch { }
                finally { Remove-ComReference $query }
            }
        }
        finally { Remove-ComReference $queries }
    }

    foreach ($connection in $connections) {
        $connectionName = [string]$connection.Name
        if ($claimed.ContainsKey($connectionName.ToLowerInvariant())) { continue }
        [void]$items.Add(@{
            Name = $connectionName
            ConnectionName = $connectionName
            Kind = 'Connection'
            Refreshable = $true
        })
    }

    return @($items.ToArray())
}


function Get-WorkbookConnectionRefreshTelemetry {
    <#
        Best-effort observation of one workbook connection. Excel/Power Query
        does not expose the internal M evaluation graph, but refreshable workbook
        connections commonly expose Refreshing and RefreshDate. Those signals are
        enough to show useful progress without changing RefreshAll semantics.
    #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [string]$ConnectionName
    )

    $result = @{
        Known       = $false
        Refreshing  = $false
        RefreshDate = [DateTime]::MinValue
    }
    if ([string]::IsNullOrWhiteSpace($ConnectionName)) { return $result }

    $connections = $null
    $connection  = $null
    $inner       = $null
    try {
        $connections = $Workbook.Connections
        $connection = $connections.Item($ConnectionName)

        $type = 0
        try { $type = [int]$connection.Type } catch { $type = 0 }
        if ($type -eq 1) { $inner = $connection.OLEDBConnection }
        elseif ($type -eq 2) { $inner = $connection.ODBCConnection }

        if ($null -eq $inner) { return $result }
        $result.Known = $true
        try { $result.Refreshing = [bool]$inner.Refreshing } catch { }
        try {
            $refreshDate = [DateTime]$inner.RefreshDate
            if ($refreshDate -gt [DateTime]::MinValue) { $result.RefreshDate = $refreshDate }
        }
        catch { }
        return $result
    }
    catch {
        return $result
    }
    finally {
        Remove-ComReference $inner $connection $connections
    }
}

function New-QueryProgressTracker {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][array]$QueryInfo,
        [string[]]$QueryNames = @()
    )

    $lookup = @{}
    foreach ($item in @($QueryInfo)) {
        $name = [string]$item.Name
        $connectionName = [string]$item.ConnectionName
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $lookup[$name.ToLowerInvariant()] = $item
        }
        if (-not [string]::IsNullOrWhiteSpace($connectionName)) {
            $key = $connectionName.ToLowerInvariant()
            if (-not $lookup.ContainsKey($key)) { $lookup[$key] = $item }
        }
    }

    $requested = @($QueryNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($requested.Count -eq 0) {
        # Only independently refreshable connections can be timed. Connectionless
        # helper queries can still execute as dependencies, but Excel does not
        # publish an individual running/completed signal for them.
        $requested = @($QueryInfo | Where-Object { [bool]$_.Refreshable } | ForEach-Object { [string]$_.Name })
    }

    $items = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($requestedName in $requested) {
        $key = ([string]$requestedName).ToLowerInvariant()
        $info = $null
        if ($lookup.ContainsKey($key)) { $info = $lookup[$key] }

        $displayName = [string]$requestedName
        $connectionName = ''
        if ($null -ne $info) {
            if (-not [string]::IsNullOrWhiteSpace([string]$info.Name)) { $displayName = [string]$info.Name }
            $connectionName = [string]$info.ConnectionName
        }

        $identity = $(if (-not [string]::IsNullOrWhiteSpace($connectionName)) {
            $connectionName.ToLowerInvariant()
        } else {
            $displayName.ToLowerInvariant()
        })
        if ($seen.ContainsKey($identity)) { continue }
        $seen[$identity] = $true

        $baseline = [DateTime]::MinValue
        $known = $false
        if (-not [string]::IsNullOrWhiteSpace($connectionName)) {
            $telemetry = Get-WorkbookConnectionRefreshTelemetry -Workbook $Workbook -ConnectionName $connectionName
            $known = [bool]$telemetry.Known
            $baseline = [DateTime]$telemetry.RefreshDate
        }

        [void]$items.Add(@{
            Name                = $displayName
            ConnectionName      = $connectionName
            BaselineRefreshDate = $baseline
            TelemetryKnown      = $known
            State               = 'Pending'
            StartedAt           = $null
            FinishedAt          = $null
            DurationSeconds     = -1
            ObservedActive      = $false
        })
    }

    return @{
        StartedAt = Get-Date
        Items     = @($items.ToArray())
    }
}

function Find-QueryProgressItem {
    param([hashtable]$Tracker, [string]$Name)
    if ($null -eq $Tracker -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    foreach ($item in @($Tracker.Items)) {
        if ([string]::Equals([string]$item.Name, $Name, [StringComparison]::OrdinalIgnoreCase) -or
            [string]::Equals([string]$item.ConnectionName, $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
    }
    return $null
}

function Set-QueryProgressItemStarted {
    param([hashtable]$Tracker, [string]$Name)
    $item = Find-QueryProgressItem -Tracker $Tracker -Name $Name
    if ($null -eq $item) { return }
    if ($null -eq $item.StartedAt) { $item.StartedAt = Get-Date }
    $item.State = 'Active'
    $item.ObservedActive = $true
}

function Set-QueryProgressItemReturned {
    param(
        [hashtable]$Tracker,
        [Parameter(Mandatory = $true)]$Workbook,
        [string]$Name
    )
    $item = Find-QueryProgressItem -Tracker $Tracker -Name $Name
    if ($null -eq $item) { return }

    $telemetry = Get-WorkbookConnectionRefreshTelemetry -Workbook $Workbook -ConnectionName ([string]$item.ConnectionName)
    if ($telemetry.Known -and $telemetry.Refreshing) {
        # The COM Refresh() call returned but Excel kept the connection running
        # in the background. The normal polling loop will finish the timing.
        return
    }

    if ($null -ne $item.StartedAt) {
        $item.FinishedAt = Get-Date
        $item.DurationSeconds = [Math]::Max(0, [int]($item.FinishedAt - [DateTime]$item.StartedAt).TotalSeconds)
        $item.State = 'Done'
    }
}

function Update-QueryProgressTracker {
    param(
        [hashtable]$Tracker,
        [Parameter(Mandatory = $true)]$Workbook
    )
    if ($null -eq $Tracker) { return }

    $now = Get-Date
    foreach ($item in @($Tracker.Items)) {
        if ([string]::IsNullOrWhiteSpace([string]$item.ConnectionName)) { continue }

        $telemetry = Get-WorkbookConnectionRefreshTelemetry -Workbook $Workbook -ConnectionName ([string]$item.ConnectionName)
        if (-not $telemetry.Known) { continue }
        $item.TelemetryKnown = $true

        if ($telemetry.Refreshing) {
            if ($null -eq $item.StartedAt) { $item.StartedAt = $now }
            $item.State = 'Active'
            $item.ObservedActive = $true
            continue
        }

        $refreshDateChanged = ([DateTime]$telemetry.RefreshDate -gt [DateTime]$item.BaselineRefreshDate -and
            [DateTime]$telemetry.RefreshDate -gt [DateTime]::MinValue)

        if ([string]$item.State -eq 'Active') {
            $item.FinishedAt = $now
            if ($null -ne $item.StartedAt) {
                $item.DurationSeconds = [Math]::Max(0, [int]($now - [DateTime]$item.StartedAt).TotalSeconds)
            }
            $item.State = 'Done'
        }
        elseif ($refreshDateChanged -and [string]$item.State -ne 'Done') {
            # It completed between polling samples. Mark completion but do not
            # invent a precise duration that Excel did not expose.
            $item.FinishedAt = $now
            $item.State = 'Done'
        }
    }
}

function Publish-QueryProgress {
    param(
        [hashtable]$Tracker,
        [Parameter(Mandatory = $true)][hashtable]$Shared
    )
    if ($null -eq $Tracker) { return }

    $items = @($Tracker.Items)
    $total = $items.Count
    if ($total -le 0) {
        $Shared.CurrentJob.QueryTotal = 0
        return
    }

    $done = @($items | Where-Object { [string]$_.State -eq 'Done' })
    $active = @($items | Where-Object { [string]$_.State -eq 'Active' })
    $now = Get-Date

    $currentName = ''
    $currentElapsed = 0
    if ($active.Count -gt 0) {
        $currentName = [string]$active[0].Name
        if ($active.Count -gt 1) { $currentName += (' +{0} more' -f ($active.Count - 1)) }
        if ($null -ne $active[0].StartedAt) {
            $currentElapsed = [Math]::Max(0, [int]($now - [DateTime]$active[0].StartedAt).TotalSeconds)
        }
    }

    $position = 0
    if ($done.Count -ge $total) { $position = $total }
    elseif ($active.Count -gt 0) { $position = [Math]::Min($total, $done.Count + 1) }
    else { $position = [Math]::Min($total, [Math]::Max(1, $done.Count + 1)) }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Query progress reported by Excel. Dependencies may run in parallel, so timings can overlap.')
    foreach ($item in $items) {
        $prefix = switch ([string]$item.State) {
            'Done'   { '[OK]' }
            'Active' { '[>>]' }
            default  { '[..]' }
        }

        $timing = ''
        if ([string]$item.State -eq 'Active' -and $null -ne $item.StartedAt) {
            $seconds = [Math]::Max(0, [int]($now - [DateTime]$item.StartedAt).TotalSeconds)
            $timing = ' - ' + (Format-Elapsed ([TimeSpan]::FromSeconds($seconds)))
        }
        elseif ([int]$item.DurationSeconds -ge 0) {
            $timing = ' - ' + (Format-Elapsed ([TimeSpan]::FromSeconds([int]$item.DurationSeconds)))
        }
        elseif ([string]$item.State -eq 'Done') {
            $timing = ' - completed (individual timing was not exposed by Excel)'
        }

        [void]$lines.Add(('{0} {1}{2}' -f $prefix, [string]$item.Name, $timing))
    }

    $Shared.CurrentJob.QueryTotal          = $total
    $Shared.CurrentJob.QueryCompleted      = $done.Count
    $Shared.CurrentJob.QueryPosition       = $position
    $Shared.CurrentJob.QueryName           = $currentName
    $Shared.CurrentJob.QueryElapsedSeconds = $currentElapsed
    $Shared.CurrentJob.QueryProgressDetail = ($lines.ToArray() -join [Environment]::NewLine)
}

function Complete-QueryProgressTracker {
    param(
        [hashtable]$Tracker,
        [Parameter(Mandatory = $true)][hashtable]$Shared
    )
    if ($null -eq $Tracker) { return @() }

    $now = Get-Date
    foreach ($item in @($Tracker.Items)) {
        if ([string]$item.State -eq 'Active') {
            $item.FinishedAt = $now
            if ($null -ne $item.StartedAt) {
                $item.DurationSeconds = [Math]::Max(0, [int]($now - [DateTime]$item.StartedAt).TotalSeconds)
            }
        }
        $item.State = 'Done'
    }
    Publish-QueryProgress -Tracker $Tracker -Shared $Shared

    $results = New-Object System.Collections.ArrayList
    foreach ($item in @($Tracker.Items)) {
        [void]$results.Add(@{
            name     = [string]$item.Name
            seconds  = [int]$item.DurationSeconds
            observed = [bool]$item.ObservedActive
        })
    }
    return @($results.ToArray())
}

function Invoke-SelectedPowerQueryRefresh {
    <# Refreshes only the explicitly selected Power Query workbook connections. #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string[]]$QueryNames,
        [scriptblock]$OnQueryStarting,
        [scriptblock]$OnQueryReturned
    )

    $requested = @($QueryNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($requested.Count -eq 0) { throw 'No Power Query names were selected.' }

    $info = @(Get-WorkbookPowerQueryInfo -Workbook $Workbook)
    $map = @{}
    foreach ($item in $info) {
        $infoName = [string]$item.Name
        $map[$infoName.ToLowerInvariant()] = $item
        # Earlier versions offered the connection's own name in the picker, so a
        # saved selection may still hold it.
        $connectionName = [string]$item.ConnectionName
        if (-not [string]::IsNullOrWhiteSpace($connectionName)) {
            $connectionKey = $connectionName.ToLowerInvariant()
            if (-not $map.ContainsKey($connectionKey)) { $map[$connectionKey] = $item }
        }
    }

    $missing = New-Object System.Collections.ArrayList
    $targets = New-Object System.Collections.ArrayList
    foreach ($name in $requested) {
        $key = ([string]$name).ToLowerInvariant()
        if (-not $map.ContainsKey($key) -or -not [bool]$map[$key].Refreshable) {
            [void]$missing.Add([string]$name)
        }
        else {
            [void]$targets.Add(@{
                RequestedName  = [string]$name
                DisplayName    = [string]$map[$key].Name
                ConnectionName = [string]$map[$key].ConnectionName
            })
        }
    }
    if ($missing.Count -gt 0) {
        throw ('The following selected queries do not have a refreshable workbook connection: {0}' -f (@($missing.ToArray()) -join ', '))
    }

    $connections = $null
    try { $connections = $Workbook.Connections } catch { $connections = $null }
    if ($null -eq $connections) { throw 'This workbook has no refreshable workbook connections.' }

    try {
        $index = 0
        foreach ($target in @($targets.ToArray())) {
            $index++
            $connection = $null
            $startedAt = Get-Date
            try {
                if ($null -ne $OnQueryStarting) {
                    try { & $OnQueryStarting ([string]$target.DisplayName) $index $targets.Count | Out-Null } catch { }
                }

                $connection = $connections.Item([string]$target.ConnectionName)
                $connection.Refresh()
            }
            finally {
                Remove-ComReference $connection
                if ($null -ne $OnQueryReturned) {
                    try {
                        $elapsed = [Math]::Max(0, [int]((Get-Date) - $startedAt).TotalSeconds)
                        & $OnQueryReturned ([string]$target.DisplayName) $index $targets.Count $elapsed | Out-Null
                    }
                    catch { }
                }
            }
        }
    }
    finally { Remove-ComReference $connections }

    return $requested
}

function Set-WorkbookQueryRefreshStamp {
    <# Persist the refresh-completion time independently of the file save time.
       msoPropertyTypeString (4) keeps the ISO value portable across Office
       versions and lets the dashboard read it directly from custom.xml. #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][DateTime]$When
    )

    $properties = $null
    $property = $null
    $name = 'ExcelQueryTriggerLastQueryRefreshUtc'
    $value = $When.ToUniversalTime().ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
    try {
        $properties = $Workbook.CustomDocumentProperties
        try {
            $property = $properties.Item($name)
            $property.Value = $value
        }
        catch {
            Remove-ComReference $property
            $property = $null
            $property = $properties.Add($name, $false, 4, $value)
        }
        return $true
    }
    catch { return $false }
    finally { Remove-ComReference $property $properties }
}

function Initialize-ExcelProcessIdentityHelper {
    if ('ExcelProcessIdentityHelper' -as [type]) { return $true }
    $source = @'
using System;
using System.Runtime.InteropServices;
public static class ExcelProcessIdentityHelper
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    try {
        Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
        return $true
    }
    catch { return $false }
}

function Get-ExcelProcessIdentity {
    param([Parameter(Mandatory = $true)]$Excel)
    $empty = @{ Pid = 0; Hwnd = 0; StartedAtUtc = $null }
    if (-not (Initialize-ExcelProcessIdentityHelper)) { return $empty }

    try {
        $hwndValue = [int64]$Excel.Hwnd
        if ($hwndValue -eq 0) { return $empty }
        $pidValue = [uint32]0
        [void][ExcelProcessIdentityHelper]::GetWindowThreadProcessId([IntPtr]$hwndValue, [ref]$pidValue)
        if ($pidValue -eq 0) { return $empty }
        $process = Get-Process -Id ([int]$pidValue) -ErrorAction Stop
        if ($process.ProcessName -ne 'EXCEL') { return $empty }
        return @{
            Pid = [int]$pidValue
            Hwnd = $hwndValue
            StartedAtUtc = $process.StartTime.ToUniversalTime()
        }
    }
    catch { return $empty }
}

function Test-ExcelProcessIdentity {
    param([int]$ProcessId, $StartedAtUtc)
    if ($ProcessId -le 0 -or $null -eq $StartedAtUtc) { return $false }
    try {
        $process = Get-Process -Id $ProcessId -ErrorAction Stop
        return ($process.ProcessName -eq 'EXCEL' -and
            $process.StartTime.ToUniversalTime() -eq [DateTime]$StartedAtUtc)
    }
    catch { return $false }
}

function Invoke-ExcelRefresh {
    <#
        Refreshes one workbook. Never throws: every outcome is reported in the
        returned hashtable so the job runner can decide whether to continue.

        Returns @{
            Success; ErrorType; Message; Workbook; StartedAt; FinishedAt;
            ElapsedSeconds; Saved; ConnectionsSwitched; RefreshSeconds
        }
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Action,
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [string]$RuleName = '',
        [scriptblock]$OnStage,
        [scriptblock]$ShouldAbort,
        [bool]$CheckDataSources = $true
    )

    $result = @{
        Success             = $false
        ErrorType           = ''
        Message             = ''
        Workbook            = [string]$Action.path
        StartedAt           = (Get-Date)
        FinishedAt          = $null
        ElapsedSeconds      = 0
        Saved               = $false
        ConnectionsSwitched = 0
        RefreshSeconds      = 0
        RefreshCallSeconds  = 0
        RefreshConfirmationSeconds = 0
        SaveSeconds         = 0
        SaveMode            = 'Normal'
        RefreshMethod       = [string]$Action.refreshMethod
        RefreshedQueries    = @()
        QueryTimings        = @()
        QueryRefreshedAt    = $null
    }

    $reportStage = {
        param([string]$Stage)
        try {
            $Shared.CurrentJob.CanCancel = ($Stage -in @('Pending', 'WaitingForFile', 'CheckingWorkbook', 'OpeningExcel', 'Refreshing', 'RefreshingLong'))
            $Shared.CurrentJob.CanForceTerminate = ($Stage -in @('OpeningExcel', 'Refreshing', 'RefreshingLong'))
            # Dialog inspection is deliberately wider than termination. During a
            # direct save Excel must never be killed, but a modal dialog raised
            # there still has to be seen and closed, otherwise the job waits for
            # an answer nobody can give.
            $Shared.CurrentJob.CanInspectDialogs = ($Stage -in @('OpeningExcel', 'Refreshing', 'RefreshingLong', 'Saving'))
        }
        catch { }
        if ($null -ne $OnStage) { try { & $OnStage $Stage | Out-Null } catch { } }
    }

    $excel        = $null
    $workbooks    = $null
    $workbook     = $null
    $workbookOpen = $false
    $ownedPid     = 0
    $connectionBackgroundStates = @()
    $connectionSettingsRestored = $false
    $workbookCloseSucceeded = $false
    $saveStarted = $false
    $saveCompleted = $false
    $saveOperationStartedAt = $null
    $ownedProcessStartedAtUtc = $null
    $allowWorkbookMacros = ConvertTo-BoolValue $Action.allowWorkbookMacros $false

    try {
        if (-not (Test-Path -LiteralPath $Action.path)) {
            $result.ErrorType = 'WorkbookNotFound'
            $result.Message   = 'Workbook not found or the path is unreachable: {0}' -f $Action.path
            return $result
        }

        & $reportStage 'CheckingWorkbook'
        Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path -Stage 'CheckingWorkbook' `
            -Message 'Checking whether the workbook is already open before starting Excel.'

        $lock = Test-WorkbookLocked -Path $Action.path -Attempts 3 -DelayMilliseconds 300
        if ($lock.Locked) {
            $result.ErrorType = 'WorkbookLocked'
            $owner = [string]$lock.Owner
            if ([string]::IsNullOrWhiteSpace($owner)) { $owner = 'another user or another Excel window' }
            $result.Message = 'Refresh was not started because the workbook is already in use by {0}. No queries were refreshed.' -f $owner
            Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path -Stage 'CheckingWorkbook' `
                -ErrorType 'WorkbookLocked' -Message $result.Message
            return $result
        }

        # The probe above opened the workbook exclusively. An SMB client can
        # hold that handle open for a moment after it is closed, so give it time
        # to drain before Excel opens the same file.
        Start-Sleep -Milliseconds 400

        & $reportStage 'OpeningExcel'
        Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'OpeningExcel' -Message 'Starting a dedicated Excel instance.'

        try {
            $excel = New-Object -ComObject Excel.Application
        }
        catch {
            $result.ErrorType = 'ExcelOpenError'
            $result.Message   = 'Could not start Excel: {0}' -f $_.Exception.Message
            return $result
        }

        $excelIdentity = Get-ExcelProcessIdentity -Excel $excel
        if ([int]$excelIdentity.Pid -gt 0) {
            $ownedPid = [int]$excelIdentity.Pid
            $ownedProcessStartedAtUtc = [DateTime]$excelIdentity.StartedAtUtc
            $Shared.OwnedExcelPid = $ownedPid
            $Shared.OwnedExcelStartedAtUtc = $ownedProcessStartedAtUtc
            $Shared.OwnedExcelWindowHandle = [int64]$excelIdentity.Hwnd
            Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Message ('Owned EXCEL.EXE identity: pid={0}, hwnd={1}, started={2:o}' -f $ownedPid, $excelIdentity.Hwnd, $ownedProcessStartedAtUtc)
        }
        else {
            Write-AppLog -Level 'WARN' -RuleName $RuleName `
                -Message 'The dedicated Excel process identity could not be proven from Excel.Application.Hwnd. Automatic process termination is disabled for this job.'
        }

        try {
            $excel.Visible          = [bool]$Action.visible
            $excel.DisplayAlerts    = $false
            $excel.AskToUpdateLinks = $false
            $excel.ScreenUpdating   = [bool]$Action.visible
        }
        catch { }

        # The dashboard's dialog watchdog only acts on a hidden instance; when
        # the window is shown on purpose the dialogs are there to be read.
        try { $Shared.CurrentJob.ExcelVisible = [bool]$Action.visible } catch { }

        # 3 = msoAutomationSecurityForceDisable; 2 = msoAutomationSecurityByUI.
        # Blocking macros is a safety boundary, so failure to apply ForceDisable
        # must stop the job before Workbooks.Open has a chance to run VBA.
        try {
            if ($allowWorkbookMacros) {
                $excel.AutomationSecurity = 2
                $excel.EnableEvents = $true
                Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path `
                    -Message 'Workbook macros are allowed for this action, subject to Excel Trust Center policy.'
            }
            else {
                $excel.AutomationSecurity = 3
                $excel.EnableEvents = $false
                Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path `
                    -Message 'Workbook macros are blocked for this automated refresh.'
            }
        }
        catch {
            if (-not $allowWorkbookMacros) {
                $result.ErrorType = 'MacroSecurityError'
                $result.Message = 'Excel macro security could not be forced to Disabled, so the workbook was not opened.'
                return $result
            }
            Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path `
                -Message 'Could not explicitly apply Excel Trust Center macro policy; continuing because macros were explicitly allowed for this action.'
        }

        $workbooks = $excel.Workbooks
        $missing   = [System.Reflection.Missing]::Value
        # Open(Filename, UpdateLinks:=0, ReadOnly:=$false, ... IgnoreReadOnlyRecommended:=$true, ... Notify:=$false)
        $workbook = $workbooks.Open($Action.path, 0, $false, $missing, $missing, $missing, $true,
            $missing, $missing, $missing, $false)
        $workbookOpen = $true

        $isReadOnly = $false
        try { $isReadOnly = [bool]$workbook.ReadOnly } catch { $isReadOnly = $false }
        if ($isReadOnly -and (ConvertTo-BoolValue $Action.save $true)) {
            $result.ErrorType = 'WorkbookLocked'
            $result.Message   = 'Refresh was not started because Excel opened the workbook read-only, which means another user or Excel window has the writable copy. No queries were refreshed.'
            Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path -Stage 'CheckingWorkbook' `
                -ErrorType 'WorkbookLocked' -Message $result.Message
            return $result
        }
        if (ConvertTo-BoolValue $Action.save $true) {
            Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path -Stage 'CheckingWorkbook' `
                -Message 'Workbook opened for writing. Excel will keep it open through refresh, save and close.'
        }

        # Before touching the queries: an unreachable data source is what makes
        # the provider raise its own credential wizard, and Excel then sits in a
        # modal dialog inside a COM call where cancellation cannot reach it.
        if ($CheckDataSources -and (ConvertTo-BoolValue $Action.checkConnections $true)) {
            $connectionCheck = Test-WorkbookConnectionTargets -Workbook $workbook
            foreach ($connectionWarning in @($connectionCheck.Warnings)) {
                Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path -Message $connectionWarning
            }
            if (@($connectionCheck.Blocked).Count -gt 0) {
                $result.ErrorType = 'ConnectionUnreachable'
                $result.Message   = 'The workbook was not refreshed because a data source is unreachable. {0}' -f (@($connectionCheck.Blocked) -join ' ')
                return $result
            }
            Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path `
                -Message ('Data source check: {0} server(s) checked, {1} connection(s) not checkable{2}.' -f `
                    $connectionCheck.Checked, @($connectionCheck.Skipped).Count, `
                    $(if (@($connectionCheck.Skipped).Count -gt 0) { ' (' + (@($connectionCheck.Skipped) -join ', ') + ')' } else { '' }))
        }

        if (ConvertTo-BoolValue $Action.disableBackgroundQuery $true) {
            $foregroundState = Set-WorkbookConnectionsForeground -Workbook $workbook
            $result.ConnectionsSwitched = ConvertTo-IntValue $foregroundState.ChangedCount 0 0
            $connectionBackgroundStates = @($foregroundState.States)
            if ($result.ConnectionsSwitched -gt 0) {
                Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path `
                    -Message ('Temporarily switched {0} connection(s) to synchronous refresh.' -f $result.ConnectionsSwitched)
            }
        }

        & $reportStage 'Refreshing'
        $refreshOperationStartedAt = Get-Date
        $refreshCallStartedAt = Get-Date
        Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' -Message 'Refresh started.'

        $queryInfo = @(Get-WorkbookPowerQueryInfo -Workbook $workbook)
        $allQueryNames = @($queryInfo | ForEach-Object { [string]$_.Name })
        $refreshMethod = [string]$Action.refreshMethod
        $queryTracker = $null

        if ($refreshMethod -eq 'SelectedQueries') {
            $selected = @($Action.selectedQueries | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $queryTracker = New-QueryProgressTracker -Workbook $workbook -QueryInfo $queryInfo -QueryNames $selected
            Publish-QueryProgress -Tracker $queryTracker -Shared $Shared

            $onSelectedQueryStarting = {
                param([string]$QueryName, [int]$Index, [int]$Total)
                Set-QueryProgressItemStarted -Tracker $queryTracker -Name $QueryName
                Publish-QueryProgress -Tracker $queryTracker -Shared $Shared
            }.GetNewClosure()

            $onSelectedQueryReturned = {
                param([string]$QueryName, [int]$Index, [int]$Total, [int]$ElapsedSeconds)
                Set-QueryProgressItemReturned -Tracker $queryTracker -Workbook $workbook -Name $QueryName
                Publish-QueryProgress -Tracker $queryTracker -Shared $Shared
            }.GetNewClosure()

            $result.RefreshedQueries = @(Invoke-SelectedPowerQueryRefresh -Workbook $workbook -QueryNames $selected `
                -OnQueryStarting $onSelectedQueryStarting -OnQueryReturned $onSelectedQueryReturned)
            $refreshCallFinishedAt = Get-Date
            $result.RefreshCallSeconds = [Math]::Max(0, [int]($refreshCallFinishedAt - $refreshCallStartedAt).TotalSeconds)
            Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
                -Message ('Selected query refresh call(s) returned after {0}. Queries: {1}' -f (Format-Elapsed ([TimeSpan]::FromSeconds($result.RefreshCallSeconds))), ($result.RefreshedQueries -join ', '))
        }
        else {
            $result.RefreshMethod = 'RefreshAll'
            $result.RefreshedQueries = @($allQueryNames)
            $queryTracker = New-QueryProgressTracker -Workbook $workbook -QueryInfo $queryInfo
            Publish-QueryProgress -Tracker $queryTracker -Shared $Shared

            $queryText = $(if ($allQueryNames.Count -gt 0) { $allQueryNames -join ', ' } else { 'No Power Query names detected (RefreshAll)' })
            # Log before the COM call. RefreshAll itself can block when a workbook
            # uses foreground connections; if that happens the Dashboard still
            # shows that Excel has entered RefreshAll instead of looking frozen.
            Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
                -Message ('Starting RefreshAll. Queries: {0}' -f $queryText)
            $workbook.RefreshAll()
            $refreshCallFinishedAt = Get-Date
            $result.RefreshCallSeconds = [Math]::Max(0, [int]($refreshCallFinishedAt - $refreshCallStartedAt).TotalSeconds)
            Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
                -Message ('RefreshAll returned after {0}; waiting for Excel to confirm all refresh work is finished.' -f (Format-Elapsed ([TimeSpan]::FromSeconds($result.RefreshCallSeconds))))
        }

        $onPoll = {
            param([string]$State, [int]$Seconds)
            if ($null -ne $queryTracker) {
                Update-QueryProgressTracker -Tracker $queryTracker -Workbook $workbook
                Publish-QueryProgress -Tracker $queryTracker -Shared $Shared
            }
            Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path `
                -Message ('Refresh poll: state={0} elapsed={1}s' -f $State, $Seconds) -NoActivity
        }.GetNewClosure()

        $onWarning = {
            param([int]$WarningSeconds)
            & $reportStage 'RefreshingLong'
            Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
                -Message ('Refresh is taking longer than expected ({0}s). Excel will remain open and the application will continue waiting.' -f $WarningSeconds)
        }.GetNewClosure()

        $result.RefreshConfirmationSeconds = Wait-ExcelRefreshCompletion -Excel $excel -Workbook $workbook `
            -TimeoutSeconds (ConvertTo-IntValue $Action.timeoutSeconds 300 5) `
            -ShouldAbort $ShouldAbort -OnPoll $onPoll -OnWarning $onWarning

        if ($null -ne $queryTracker) {
            $result.QueryTimings = @(Complete-QueryProgressTracker -Tracker $queryTracker -Shared $Shared)
            $timedQueries = @($result.QueryTimings | Where-Object { [int]$_.seconds -ge 0 })
            if ($timedQueries.Count -gt 0) {
                $timingText = @($timedQueries | ForEach-Object {
                    '{0}={1}' -f [string]$_.name, (Format-Elapsed ([TimeSpan]::FromSeconds([int]$_.seconds)))
                }) -join ', '
                Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
                    -Message ('Observed query timings: {0}' -f $timingText)
            }
        }

        # Query completion and workbook saving are separate events. Excel does
        # not persist a refresh time for every Power Query layout, so remember
        # the completion time explicitly before the later save operation.
        $result.QueryRefreshedAt = Get-Date

        $result.RefreshSeconds = [Math]::Max(0, [int]((Get-Date) - $refreshOperationStartedAt).TotalSeconds)
        $refreshTotalText = Format-Elapsed ([TimeSpan]::FromSeconds($result.RefreshSeconds))
        $refreshCallText = Format-Elapsed ([TimeSpan]::FromSeconds($result.RefreshCallSeconds))
        $refreshConfirmationText = Format-Elapsed ([TimeSpan]::FromSeconds($result.RefreshConfirmationSeconds))
        $refreshedQueryText = $(if (@($result.RefreshedQueries).Count -gt 0) { @($result.RefreshedQueries) -join ', ' } else { 'none reported' })
        Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Refreshing' `
            -Message ('Refresh completed successfully in {0} (Excel call(s) {1}; completion confirmation {2}). Queries: {3}' -f `
                $refreshTotalText, $refreshCallText, $refreshConfirmationText, $refreshedQueryText)

        # Restore the workbook's original BackgroundQuery settings before save.
        if (@($connectionBackgroundStates).Count -gt 0) {
            $restore = Restore-WorkbookConnectionBackgroundQuery -Workbook $workbook -States $connectionBackgroundStates
            if ((ConvertTo-IntValue $restore.Failed 0 0) -gt 0) {
                $connectionSettingsRestored = $false
                $failedConnectionText = $(if (@($restore.FailedNames).Count -gt 0) {
                    ' ({0})' -f (@($restore.FailedNames) -join ', ')
                } else { '' })
                throw ([System.InvalidOperationException]::new(
                    ('Could not restore BackgroundQuery on {0} connection(s){1}. The workbook will not be saved.' -f `
                        $restore.Failed, $failedConnectionText)))
            }
            else {
                $connectionSettingsRestored = $true
                Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path `
                    -Message 'Original BackgroundQuery settings restored before save.'
            }
        }

        if (ConvertTo-BoolValue $Action.save $true) {
            if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
                throw (New-Object System.OperationCanceledException('Refresh cancelled before the workbook save started.'))
            }
            & $reportStage 'Saving'
            if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
                throw (New-Object System.OperationCanceledException('Refresh cancelled before the workbook save started.'))
            }
            $stampWritten = Set-WorkbookQueryRefreshStamp -Workbook $workbook -When $result.QueryRefreshedAt
            if (-not $stampWritten) {
                # This marker is only a portability aid for the Dashboard. The
                # authoritative refresh time is also persisted in local job
                # history, so a workbook that rejects CustomDocumentProperties
                # must not turn an otherwise successful refresh into a warning.
                Write-AppLog -Level 'DEBUG' -RuleName $RuleName -Workbook $Action.path -Stage 'Saving' `
                    -Message 'Workbook custom refresh timestamp was not written; local job history will retain the query refresh time.'
            }

            $saveOperationStartedAt = Get-Date
            $Shared.CurrentJob.SaveStartedAt = $saveOperationStartedAt
            $Shared.CurrentJob.SaveMode = 'Normal direct save'
            $Shared.CurrentJob.SaveProgressDetail = 'Excel is writing the original workbook; this step cannot be interrupted safely.'
            Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Saving' `
                -Message 'Saving workbook directly. Excel may take time to write a large or complex file; this step cannot be interrupted safely.'
            $saveStarted = $true
            $workbook.Save()
            $saveCompleted = $true
            $result.Saved = $true
            $result.SaveSeconds = [Math]::Max(0, [int]((Get-Date) - $saveOperationStartedAt).TotalSeconds)
            Write-AppLog -Level 'INFO' -RuleName $RuleName -Workbook $Action.path -Stage 'Saving' `
                -Message ('Workbook saved in {0}.' -f (Format-Elapsed ([TimeSpan]::FromSeconds($result.SaveSeconds))))
        }

        if (ConvertTo-BoolValue $Action.close $true -and $workbookOpen) {
            & $reportStage 'Closing'
            $workbook.Close($false)
            $workbookOpen = $false
            $workbookCloseSucceeded = $true
        }

        $result.Success = $true
        $querySummary = $(if (@($result.RefreshedQueries).Count -gt 0) { @($result.RefreshedQueries) -join ', ' } else { 'No Power Query names reported' })
        $result.Message = ('Refresh completed successfully. Queries: {0}' -f $querySummary)
        return $result
    }
    catch [System.OperationCanceledException] {
        $result.ErrorType = 'Cancelled'
        $result.Message   = $_.Exception.Message
        return $result
    }
    catch [System.Runtime.InteropServices.COMException] {
        # Ending the Excel process breaks whatever COM call was blocked. That is
        # the cancellation taking effect, not a refresh failure.
        if ([bool]$Shared.ForcedExcelTermination) {
            $result.ErrorType = 'Cancelled'
            $result.Message   = 'The refresh was stopped by ending the unresponsive Excel instance from the dashboard.'
            return $result
        }
        $result.ErrorType = 'ExcelRefreshError'
        $result.Message   = 'Excel COM error 0x{0:X8}: {1}' -f $_.Exception.HResult, $_.Exception.Message
        return $result
    }
    catch {
        if ([bool]$Shared.ForcedExcelTermination) {
            $result.ErrorType = 'Cancelled'
            $result.Message   = 'The refresh was stopped by ending the unresponsive Excel instance from the dashboard.'
            return $result
        }
        $result.ErrorType = 'UnexpectedError'
        $result.Message   = '{0}: {1}' -f $_.Exception.GetType().Name, $_.Exception.Message
        return $result
    }
    finally {
        # ---- cleanup always runs, on success and on every failure path ----
        & $reportStage 'Cleanup'

        if ($workbookOpen -and $null -ne $workbook) {
            # On an error path this close is deliberately SaveChanges=False. If
            # we temporarily changed BackgroundQuery, restore it in memory first
            # even though the failed workbook will not be saved.
            if (-not $connectionSettingsRestored -and @($connectionBackgroundStates).Count -gt 0) {
                try { [void](Restore-WorkbookConnectionBackgroundQuery -Workbook $workbook -States $connectionBackgroundStates) } catch { }
            }
            try {
                $workbook.Close($false)
                $workbookOpen = $false
                $workbookCloseSucceeded = $true
            }
            catch {
                Write-AppLog -Level 'ERROR' -RuleName $RuleName -Workbook $result.Workbook `
                    -Message 'Excel could not confirm that the workbook closed.'
            }
        }
        elseif (-not $workbookOpen) {
            $workbookCloseSucceeded = $true
        }
        Remove-ComReference $workbook $workbooks
        $workbook  = $null
        $workbooks = $null

        if ($null -ne $excel) {
            try { $excel.DisplayAlerts = $false } catch { }
            try { $excel.Quit() } catch { }
            Remove-ComReference $excel
            $excel = $null
        }

        Clear-ComMemory

        if ($ownedPid -gt 0) {
            # Excel can need several seconds to finish its own shutdown after
            # Quit(), especially after a large refresh. Never race it at 400 ms.
            $exitDeadline = (Get-Date).AddSeconds(5)
            $orphan = $null
            do {
                try { $orphan = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue } catch { $orphan = $null }
                if ($null -eq $orphan) { break }
                Start-Sleep -Milliseconds 250
            } while ((Get-Date) -lt $exitDeadline)

            # Refuse termination unless ownership still matches and either the
            # workbook closed or cancellation occurred before Save began.
            $jobWasCancelled = ($result.ErrorType -eq 'Cancelled') -or [bool]$Shared.CancelCurrentJob -or [bool]$Shared.ForcedExcelTermination
            $identityMatches = Test-ExcelProcessIdentity -ProcessId $ownedPid -StartedAtUtc $ownedProcessStartedAtUtc
            $mayTerminate    = $identityMatches -and ($workbookCloseSucceeded -or ($jobWasCancelled -and -not $saveStarted))

            if ($null -ne $orphan) {
                if ($mayTerminate) {
                    try {
                        Stop-Process -Id $ownedPid -Force -ErrorAction Stop
                        $reason = $(if ($workbookCloseSucceeded) { 'after the workbook had closed' } else { 'because the job was cancelled before saving started' })
                        Write-AppLog -Level 'WARN' -RuleName $RuleName -Workbook $result.Workbook `
                            -Message ('Excel did not exit after 5 seconds; the dedicated instance started by this application (pid {0}) was terminated {1}.' -f $ownedPid, $reason)
                    }
                    catch { }
                }
                else {
                    $reason = $(if (-not $identityMatches) { 'its process identity could not be revalidated' } elseif ($saveStarted -and -not $saveCompleted) { 'a direct save may still be in progress' } else { 'workbook closure could not be confirmed' })
                    Write-AppLog -Level 'ERROR' -RuleName $RuleName -Workbook $result.Workbook `
                        -Message ('Excel process {0} was left running because {1}. Please check it manually.' -f $ownedPid, $reason)
                }
            }
        }

        $Shared.OwnedExcelPid = 0
        $Shared.OwnedExcelStartedAtUtc = $null
        $Shared.OwnedExcelWindowHandle = 0

        $result.FinishedAt     = Get-Date
        $result.ElapsedSeconds = [int]($result.FinishedAt - $result.StartedAt).TotalSeconds
    }
}
