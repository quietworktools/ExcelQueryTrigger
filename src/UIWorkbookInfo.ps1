# ==============================================================================
#  UIWorkbookInfo.ps1  (UI runspace only)
#
#  Two questions the dashboard could not answer before:
#
#    "when was the data last refreshed?"   - not by this tool; at all. On a
#       shared workbook somebody else may have refreshed it an hour ago, and
#       our own Last Run column says nothing about that.
#
#    "is this file shared?"                - because a shared file behaves
#       differently: another person can refresh it, it can be locked when we
#       want it, and it may be syncing while we read it.
#
#  Both are answered by reading the workbook file itself. A .xlsx is a zip, so
#  no Excel instance is started, nothing is locked, and it works while a refresh
#  is running.
#
#    docProps/core.xml                  dcterms:modified, cp:lastModifiedBy
#    docProps/custom.xml                app-recorded query refresh time
#    xl/pivotCache/pivotCacheDefinition*.xml   refreshedDate, refreshedBy
#    xl/revisions/                      present only on a legacy shared workbook
#
#  Reading is cached against the file's own timestamp and rate-limited, because
#  a workbook on a disconnected network share can take seconds just to stat.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:WorkbookInfoCache = @{}
$script:WorkbookInfoBackoff = @{}

function New-EmptyWorkbookInfo {
    return @{
        Ok              = $false
        Exists          = $false
        SavedAt         = [DateTime]::MinValue
        SavedBy         = ''
        DataRefreshedAt = [DateTime]::MinValue
        RefreshedBy     = ''
        Location        = 'Unknown'
        LocationDetail  = ''
        SharedWorkbook  = $false
        Reason          = ''
    }
}

function Get-WorkbookLocationKind {
    <#
        Where the file lives, in the terms that matter operationally.

        Local        nobody else can be in it
        Network      a UNC path or a mapped drive - somebody else can
        OneDrive     a personal sync folder
        SharePoint   a OneDrive for Business / SharePoint sync folder
        Removable    a USB stick or similar, which may simply vanish
    #>
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return @{ Kind = 'Unknown'; Detail = '' } }

    if ($Path.StartsWith('\\')) {
        $server = ''
        $parts = @($Path.TrimStart('\') -split '\\')
        if ($parts.Count -gt 0) { $server = $parts[0] }
        return @{ Kind = 'Network'; Detail = ('On the shared folder \\{0}' -f $server) }
    }

    # OneDrive sync roots come through as ordinary local paths, so the
    # environment has to be asked which of them are really sync folders.
    foreach ($pair in @(
        @{ Var = 'OneDriveCommercial'; Kind = 'SharePoint'; Detail = 'In a OneDrive for Business / SharePoint folder, so it syncs and other people may open it' },
        @{ Var = 'OneDriveConsumer';   Kind = 'OneDrive';   Detail = 'In a personal OneDrive folder, so it syncs' },
        @{ Var = 'OneDrive';           Kind = 'OneDrive';   Detail = 'In a OneDrive folder, so it syncs' })) {

        $root = ''
        try { $root = [string](Get-Item -LiteralPath ('Env:' + $pair.Var) -ErrorAction SilentlyContinue).Value } catch { $root = '' }
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        if ($Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return @{ Kind = [string]$pair.Kind; Detail = [string]$pair.Detail }
        }
    }

    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        if (-not [string]::IsNullOrWhiteSpace($root)) {
            $drive = New-Object System.IO.DriveInfo($root)
            switch ($drive.DriveType) {
                'Network'   { return @{ Kind = 'Network';   Detail = ('On the mapped drive {0}' -f $root.TrimEnd('\')) } }
                'Removable' { return @{ Kind = 'Removable'; Detail = ('On removable media ({0})' -f $root.TrimEnd('\')) } }
            }
        }
    }
    catch { }

    return @{ Kind = 'Local'; Detail = 'On this computer only' }
}

function Read-WorkbookMetadata {
    <#
        Opens the workbook as a zip and reads the three parts that carry the
        answers. A copy is taken first so a file that is open in Excel, or being
        written by OneDrive, can still be read.
    #>
    param([string]$Path, [hashtable]$Info)

    $temp = Copy-FileForReading -Path $Path
    if ([string]::IsNullOrWhiteSpace($temp)) {
        $Info.Reason = 'The file could not be read just now.'
        return $Info
    }

    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($temp)

        foreach ($entry in $archive.Entries) {
            $name = [string]$entry.FullName

            if ($name -eq 'docProps/core.xml') {
                try {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $text = $reader.ReadToEnd(); $reader.Dispose()
                    $xml = New-Object System.Xml.XmlDocument
                    $xml.LoadXml($text)
                    foreach ($node in $xml.DocumentElement.ChildNodes) {
                        if ($node.LocalName -eq 'modified') {
                            $when = [DateTime]::MinValue
                            if ([DateTime]::TryParse([string]$node.InnerText, [ref]$when)) { $Info.SavedAt = $when.ToLocalTime() }
                        }
                        elseif ($node.LocalName -eq 'lastModifiedBy') {
                            $Info.SavedBy = [string]$node.InnerText
                        }
                    }
                }
                catch { }
                continue
            }

            if ($name -eq 'docProps/custom.xml') {
                try {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $text = $reader.ReadToEnd(); $reader.Dispose()
                    $xml = New-Object System.Xml.XmlDocument
                    $xml.LoadXml($text)
                    foreach ($node in $xml.DocumentElement.ChildNodes) {
                        if ($node.LocalName -ne 'property' -or
                            [string]$node.GetAttribute('name') -ne 'ExcelQueryTriggerLastQueryRefreshUtc') { continue }
                        $when = [DateTime]::MinValue
                        if ([DateTime]::TryParse([string]$node.InnerText,
                                [System.Globalization.CultureInfo]::InvariantCulture,
                                [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$when)) {
                            $localWhen = $when.ToLocalTime()
                            if ($localWhen -gt $Info.DataRefreshedAt) {
                                $Info.DataRefreshedAt = $localWhen
                                $Info.RefreshedBy = 'Excel Query Trigger Manager'
                            }
                        }
                    }
                }
                catch { }
                continue
            }

            # A legacy shared workbook keeps a revision log. Its presence is
            # proof that Excel was told several people may edit at once.
            if ($name.StartsWith('xl/revisions/')) { $Info.SharedWorkbook = $true; continue }

            if ($name -match '^xl/pivotCache/pivotCacheDefinition\d*\.xml$') {
                try {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $text = $reader.ReadToEnd(); $reader.Dispose()
                    $xml = New-Object System.Xml.XmlDocument
                    $xml.LoadXml($text)

                    $refreshedDate = [string]$xml.DocumentElement.GetAttribute('refreshedDate')
                    if (-not [string]::IsNullOrWhiteSpace($refreshedDate)) {
                        $serial = 0.0
                        # An OLE Automation date: days since 1899-12-30.
                        if ([double]::TryParse($refreshedDate, [System.Globalization.NumberStyles]::Float,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$serial) -and $serial -gt 0) {
                            $when = [DateTime]::FromOADate($serial)
                            if ($when -gt $Info.DataRefreshedAt) {
                                $Info.DataRefreshedAt = $when
                                $Info.RefreshedBy = [string]$xml.DocumentElement.GetAttribute('refreshedBy')
                            }
                        }
                    }
                }
                catch { }
            }
        }

        $Info.Ok = $true
        return $Info
    }
    catch {
        $Info.Reason = 'This file could not be read as an Office file.'
        return $Info
    }
    finally {
        if ($null -ne $archive) { try { $archive.Dispose() } catch { } }
        try { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Get-WorkbookInfo {
    <#
        Cached per file, keyed on the file's own timestamp and size, so an
        unchanged workbook is never opened twice.

        $Refresh = $false returns whatever is already known and reads nothing.
        That is what the rule list uses, so redrawing never touches the disk.
    #>
    param([string]$Path, [bool]$Refresh = $false)

    if ([string]::IsNullOrWhiteSpace($Path)) { return (New-EmptyWorkbookInfo) }
    $key = $Path.ToLowerInvariant()

    if (-not $Refresh) {
        if ($script:WorkbookInfoCache.ContainsKey($key)) { return $script:WorkbookInfoCache[$key].Info }
        return (New-EmptyWorkbookInfo)
    }

    # A workbook on a share that has gone away can take seconds just to stat.
    # After one slow or failed attempt, leave it alone for a while.
    if ($script:WorkbookInfoBackoff.ContainsKey($key)) {
        if ((Get-Date) -lt $script:WorkbookInfoBackoff[$key]) {
            if ($script:WorkbookInfoCache.ContainsKey($key)) { return $script:WorkbookInfoCache[$key].Info }
            return (New-EmptyWorkbookInfo)
        }
    }

    $started = Get-Date
    $info = New-EmptyWorkbookInfo
    $location = Get-WorkbookLocationKind -Path $Path
    $info.Location       = [string]$location.Kind
    $info.LocationDetail = [string]$location.Detail

    $file = $null
    try { $file = Get-Item -LiteralPath $Path -ErrorAction Stop }
    catch {
        $info.Reason = 'The file could not be found.'
        $script:WorkbookInfoCache[$key] = @{ Stamp = ''; Info = $info }
        $script:WorkbookInfoBackoff[$key] = (Get-Date).AddMinutes(5)
        return $info
    }

    $info.Exists = $true
    $stamp = '{0}|{1}' -f $file.LastWriteTimeUtc.Ticks, $file.Length

    if ($script:WorkbookInfoCache.ContainsKey($key) -and $script:WorkbookInfoCache[$key].Stamp -eq $stamp) {
        return $script:WorkbookInfoCache[$key].Info
    }

    $info = Read-WorkbookMetadata -Path $Path -Info $info
    # Saving is the moment the file itself changed, whoever did it.
    if ($info.SavedAt -eq [DateTime]::MinValue) { $info.SavedAt = $file.LastWriteTime }

    $script:WorkbookInfoCache[$key] = @{ Stamp = $stamp; Info = $info }
    if (((Get-Date) - $started).TotalSeconds -gt 2) {
        $script:WorkbookInfoBackoff[$key] = (Get-Date).AddMinutes(5)
    }
    else {
        [void]$script:WorkbookInfoBackoff.Remove($key)
    }
    return $info
}

function Get-WorkbookDataAgeText {
    <#  When the data in this workbook was last refreshed, in words.  #>
    param([hashtable]$Info)

    if ($null -eq $Info -or -not $Info.Exists) { return '' }
    if ($Info.DataRefreshedAt -gt [DateTime]::MinValue) { return (Get-RelativeTimeText -When $Info.DataRefreshedAt) }
    return 'Not recorded'
}

function Get-WorkbookDataAgeTooltip {
    param([hashtable]$Info)

    if ($null -eq $Info -or -not $Info.Exists) { return '' }
    $lines = New-Object System.Collections.ArrayList

    if ($Info.DataRefreshedAt -gt [DateTime]::MinValue) {
        $who = $(if ([string]::IsNullOrWhiteSpace([string]$Info.RefreshedBy)) { '' } else { ' by ' + [string]$Info.RefreshedBy })
        [void]$lines.Add(('Data last refreshed {0}{1}' -f $Info.DataRefreshedAt.ToString('dddd d MMMM, HH:mm'), $who))
    }
    else {
        [void]$lines.Add('No query refresh time is recorded in this workbook.')
    }

    if ($Info.SavedAt -gt [DateTime]::MinValue) {
        $who = $(if ([string]::IsNullOrWhiteSpace([string]$Info.SavedBy)) { '' } else { ' by ' + [string]$Info.SavedBy })
        [void]$lines.Add(('File last saved {0}{1} (not used for Data updated)' -f $Info.SavedAt.ToString('dddd d MMMM, HH:mm'), $who))
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Info.LocationDetail)) { [void]$lines.Add([string]$Info.LocationDetail) }
    if ($Info.SharedWorkbook) { [void]$lines.Add('Excel has this marked as a shared workbook, so several people can have it open at once.') }

    return ($lines.ToArray() -join [Environment]::NewLine)
}
