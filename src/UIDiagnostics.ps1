# ==============================================================================
#  UIDiagnostics.ps1  (UI runspace only)
#
#  "Test Data Sources" answers the question that a failed refresh does not:
#  which connection in which workbook is the problem, and what is wrong with it.
#
#  It never opens Excel. A workbook is a zip, and its external connections are
#  stored in xl/connections.xml, so the whole check reads the file directly.
#  That matters for three reasons: it works while a refresh is running, it works
#  on a workbook that Excel cannot open at all, and it can never raise the very
#  credential dialog we are trying to diagnose.
# ==============================================================================

Set-StrictMode -Version 1.0

function Get-WorkbookConnectionDefinitions {
    <#
        Reads xl/connections.xml straight out of the workbook.
        Returns @{ Ok; Reason; Connections = @(...) }
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $empty = @{ Ok = $false; Reason = ''; Connections = @() }

    if (-not (Test-Path -LiteralPath $Path)) {
        $empty.Reason = 'The workbook could not be found at this path.'
        return $empty
    }

    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

    $temp = Copy-FileForReading -Path $Path
    if ([string]::IsNullOrWhiteSpace($temp)) {
        $empty.Reason = 'The workbook could not be read. It may be locked by another program.'
        return $empty
    }

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($temp)

        $entry = $null
        foreach ($candidate in $archive.Entries) {
            if ($candidate.FullName -eq 'xl/connections.xml') { $entry = $candidate; break }
        }
        if ($null -eq $entry) {
            $empty.Ok = $true
            $empty.Reason = 'This workbook stores no external connections.'
            return $empty
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        $text = $reader.ReadToEnd()
        $reader.Dispose()

        $xml = New-Object System.Xml.XmlDocument
        $xml.LoadXml($text)

        $found = New-Object System.Collections.ArrayList
        foreach ($node in $xml.DocumentElement.ChildNodes) {
            if ($node.LocalName -ne 'connection') { continue }

            $name = ''
            try { $name = [string]$node.GetAttribute('name') } catch { }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = ('connection {0}' -f ($found.Count + 1)) }

            $refreshOnLoad = $false
            try { $refreshOnLoad = ([string]$node.GetAttribute('refreshOnLoad') -eq '1') } catch { }

            $description = ''
            try { $description = [string]$node.GetAttribute('description') } catch { }

            $connectionString = ''
            $command = ''
            $kind = 'Other'
            foreach ($child in $node.ChildNodes) {
                switch ($child.LocalName) {
                    'dbPr' {
                        $kind = 'Database'
                        try { $connectionString = [string]$child.GetAttribute('connection') } catch { }
                        try { $command = [string]$child.GetAttribute('command') } catch { }
                    }
                    'olapPr' { if ($kind -eq 'Other') { $kind = 'Cube' } }
                    'webPr'  { $kind = 'Web'
                        try { $connectionString = [string]$child.GetAttribute('url') } catch { } }
                    'textPr' { $kind = 'Text file'
                        try { $connectionString = [string]$child.GetAttribute('sourceFile') } catch { } }
                }
            }

            [void]$found.Add(@{
                Name             = $name
                Kind             = $kind
                ConnectionString = $connectionString
                Command          = $command
                RefreshOnLoad    = $refreshOnLoad
                Description      = $description
            })
        }

        return @{ Ok = $true; Reason = ''; Connections = @($found.ToArray()) }
    }
    catch {
        $empty.Reason = ('The workbook could not be read as an Office file: {0}' -f $_.Exception.Message)
        return $empty
    }
    finally {
        if ($null -ne $archive) { try { $archive.Dispose() } catch { } }
        try { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Test-WorkbookHasQueries {
    <#
        Does this workbook have anything that can be refreshed at all?

        Read straight out of the file, so no Excel is needed - which matters,
        because this is asked while adding a file and Excel may not even be
        installed on the machine doing the adding.

        Decided = $false means the file could not be judged (an .xlsb, an
        unreadable copy). The caller must let those through: refusing a file we
        could not read would be worse than adding one that turns out to be
        empty.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $answer = @{ Decided = $false; HasQueries = $false; Count = 0; Reason = '' }

    if (-not (Test-Path -LiteralPath $Path)) {
        $answer.Decided = $true
        $answer.Reason  = 'The file could not be found.'
        return $answer
    }

    $extension = ''
    try { $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant() } catch { $extension = '' }
    if (@('.xlsx', '.xlsm') -notcontains $extension) {
        # .xlsb keeps these parts in a binary form this cannot read.
        $answer.Reason = 'This file type cannot be checked without opening it.'
        return $answer
    }

    $definitions = Get-WorkbookConnectionDefinitions -Path $Path
    if (-not $definitions.Ok) {
        $answer.Reason = [string]$definitions.Reason
        return $answer
    }

    $count = @($definitions.Connections).Count
    if ($count -eq 0) {
        # Refusing a file wrongly is worse than accepting an empty one, so look
        # for the other two things that can be refreshed before saying no: a
        # query table on a sheet, and a PivotTable fed from outside the file.
        # An older workbook can carry either without a WorkbookConnection.
        $count = Get-WorkbookRefreshablePartCount -Path $Path
    }

    $answer.Decided    = $true
    $answer.Count      = $count
    $answer.HasQueries = ($count -gt 0)
    if ($count -eq 0) {
        $answer.Reason = 'This workbook has no queries or data connections, so there would be nothing to refresh.'
    }
    return $answer
}

function Get-WorkbookRefreshablePartCount {
    <#
        Query tables and externally-sourced PivotTables, counted straight out
        of the file. These are refreshable even when the workbook records no
        connection of its own.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $temp = Copy-FileForReading -Path $Path
    if ([string]::IsNullOrWhiteSpace($temp)) { return 0 }
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

    $archive = $null
    $count = 0
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($temp)
        foreach ($entry in $archive.Entries) {
            $name = [string]$entry.FullName
            if ($name -match '^xl/queryTables/queryTable\d*\.xml$') { $count++; continue }
            if ($name -match '^xl/pivotCache/pivotCacheDefinition\d*\.xml$') {
                try {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $text = $reader.ReadToEnd(); $reader.Dispose()
                    if ($text -match '(?i)<cacheSource[^>]*type\s*=\s*"external"') { $count++ }
                }
                catch { }
            }
        }
    }
    catch { return 0 }
    finally {
        if ($null -ne $archive) { try { $archive.Dispose() } catch { } }
        try { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } catch { }
    }
    return $count
}

function Get-DataSourceDiagnostics {
    <#
        One row per connection per workbook, each with a plain-language verdict.
        Severity is 'Error', 'Warning', 'Note' or 'Ok'.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    $rows = New-Object System.Collections.ArrayList

    foreach ($action in @($Rule.actions)) {
        $path = [string]$action.path
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        $leaf = Split-Path -Leaf $path

        $definitions = Get-WorkbookConnectionDefinitions -Path $path
        if (-not $definitions.Ok) {
            [void]$rows.Add(@{ Workbook = $leaf; Connection = '--'; Server = '--'
                Severity = 'Error'; Verdict = $definitions.Reason; Detail = $path })
            continue
        }
        if (@($definitions.Connections).Count -eq 0) {
            [void]$rows.Add(@{ Workbook = $leaf; Connection = '--'; Server = '--'
                Severity = 'Note'; Verdict = 'No external connections are stored in this workbook.'; Detail = $path })
            continue
        }

        foreach ($connection in @($definitions.Connections)) {
            $row = @{
                Workbook   = $leaf
                Connection = [string]$connection.Name
                Server     = '--'
                Severity   = 'Ok'
                Verdict    = ''
                Detail     = ''
            }

            $detail = New-Object System.Collections.ArrayList
            [void]$detail.Add(('Workbook   : {0}' -f $path))
            [void]$detail.Add(('Connection : {0}' -f $connection.Name))
            [void]$detail.Add(('Kind       : {0}' -f $connection.Kind))
            if (-not [string]::IsNullOrWhiteSpace([string]$connection.Command)) {
                [void]$detail.Add(('Command    : {0}' -f $connection.Command))
            }
            [void]$detail.Add(('Refresh on open: {0}' -f $(if ($connection.RefreshOnLoad) { 'Yes' } else { 'No' })))
            [void]$detail.Add('')
            [void]$detail.Add('Connection string:')
            [void]$detail.Add([string]$connection.ConnectionString)

            $connectionString = [string]$connection.ConnectionString

            if ($connectionString -match '(?i)Microsoft\.Mashup') {
                $row.Server   = 'Power Query'
                $row.Severity = 'Note'
                $row.Verdict  = 'A Power Query connection. The real source is inside the query itself, so it cannot be checked from here.'
            }
            else {
                $target = Get-ConnectionServerTarget -ConnectionString $connectionString
                if ($null -eq $target) {
                    $row.Severity = 'Note'
                    $row.Verdict  = 'Not a network server - a file, a web address or the workbook itself.'
                }
                else {
                    $row.Server = [string]$target.Host
                    $probe = Test-ServerReachable -ServerHost $target.Host -Port $target.Port -DefaultPort $target.DefaultPort
                    if (-not $probe.Resolved -and $probe.Conclusive) {
                        $row.Severity = 'Error'
                        $row.Verdict  = ('{0} does not exist on the network. This connection needs to be repointed.' -f $target.Host)
                    }
                    elseif (-not $probe.Resolved) {
                        $row.Severity = 'Warning'
                        $row.Verdict  = ('{0} could not be looked up just now. Check the network or the VPN.' -f $target.Host)
                    }
                    elseif ($probe.Probed -and -not $probe.PortOpen) {
                        $port = $(if ($target.Port -gt 0) { $target.Port } else { $target.DefaultPort })
                        if ($target.Port -gt 0) {
                            $row.Severity = 'Error'
                            $row.Verdict  = ('{0} answers, but nothing is listening on port {1}.' -f $target.Host, $port)
                        }
                        else {
                            $row.Severity = 'Warning'
                            $row.Verdict  = ('{0} answers, but not on the usual port {1}. That is normal for a named instance.' -f $target.Host, $port)
                        }
                    }
                    else {
                        $row.Verdict = ('{0} was reached.' -f $target.Host)
                    }
                }
            }

            # Reaching the server proves nothing about the cube or table behind
            # it, and this is the setting that makes the prompt appear during
            # Workbooks.Open, before anything else has had a chance to run.
            if ($connection.RefreshOnLoad) {
                $row.Verdict = ($row.Verdict + ' This connection also refreshes as the file opens, so any prompt appears before the refresh step.').Trim()
                if ($row.Severity -eq 'Ok') { $row.Severity = 'Note' }
            }

            $row.Detail = ($detail.ToArray() -join [Environment]::NewLine)
            [void]$rows.Add($row)
        }
    }

    # No , prefix: callers wrap this in @(), which would otherwise turn the
    # array back into a single nested element.
    return @($rows.ToArray())
}

function Show-DataSourceDiagnostics {
    <#  Runs the check for one rule and shows the result.  #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    if (@($Rule.actions).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('This rule has no workbooks yet.',
            'Test data sources', 'OK', 'Information') | Out-Null
        return
    }

    $previousCursor = [System.Windows.Forms.Cursor]::Current
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try { $rows = @(Get-DataSourceDiagnostics -Rule $Rule) }
    finally { [System.Windows.Forms.Cursor]::Current = $previousCursor }

    $errors   = @($rows | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnings = @($rows | Where-Object { $_.Severity -eq 'Warning' }).Count

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = ('Test data sources - {0}' -f [string]$Rule.name)
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize      = New-Object System.Drawing.Size(820, 520)
    $form.StartPosition   = 'CenterParent'
    $form.MinimumSize     = New-Object System.Drawing.Size(700, 440)
    # UiFonts contains Regular and Bold. The old Normal key did not exist, so
    # this one dialog silently fell back to the Windows default font.
    $form.Font            = Get-UiFont
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $summary = New-Object System.Windows.Forms.Label
    $summary.Location = New-Object System.Drawing.Point(12, 12)
    $summary.Size     = New-Object System.Drawing.Size(796, 20)
    $summary.Anchor   = 'Top,Left,Right'
    $summary.Text     = $(
        if ($errors -gt 0)        { 'Something here will stop a refresh: {0} problem(s), {1} thing(s) worth a look.' -f $errors, $warnings }
        elseif ($warnings -gt 0)  { 'Nothing looks broken, but {0} thing(s) are worth a look.' -f $warnings }
        else                      { 'Every data source that can be checked from here was reached.' })
    $summary.ForeColor = $(
        if ($errors -gt 0)       { [System.Drawing.Color]::FromArgb(180, 40, 35) }
        elseif ($warnings -gt 0) { [System.Drawing.Color]::FromArgb(170, 110, 10) }
        else                     { [System.Drawing.Color]::FromArgb(25, 115, 50) })
    $form.Controls.Add($summary)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location      = New-Object System.Drawing.Point(12, 38)
    $listView.Size          = New-Object System.Drawing.Size(796, 260)
    $listView.Anchor        = 'Top,Left,Right,Bottom'
    $listView.View          = 'Details'
    $listView.FullRowSelect = $true
    $listView.GridLines     = $true
    $listView.MultiSelect   = $false
    $listView.HideSelection = $false
    [void]$listView.Columns.Add('Workbook', 170)
    [void]$listView.Columns.Add('Connection', 150)
    [void]$listView.Columns.Add('Server', 120)
    [void]$listView.Columns.Add('Result', 330)
    $form.Controls.Add($listView)

    foreach ($row in $rows) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$row.Workbook)
        [void]$item.SubItems.Add([string]$row.Connection)
        [void]$item.SubItems.Add([string]$row.Server)
        [void]$item.SubItems.Add([string]$row.Verdict)
        $item.Tag = $row
        switch ([string]$row.Severity) {
            'Error'   { $item.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 35) }
            'Warning' { $item.ForeColor = [System.Drawing.Color]::FromArgb(170, 110, 10) }
            'Note'    { $item.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90) }
            default   { $item.ForeColor = [System.Drawing.Color]::FromArgb(25, 115, 50) }
        }
        [void]$listView.Items.Add($item)
    }

    $detailBox = New-Object System.Windows.Forms.TextBox
    $detailBox.Location   = New-Object System.Drawing.Point(12, 306)
    $detailBox.Size       = New-Object System.Drawing.Size(796, 158)
    $detailBox.Anchor     = 'Left,Right,Bottom'
    $detailBox.Multiline  = $true
    $detailBox.ReadOnly   = $true
    $detailBox.ScrollBars = 'Both'
    $detailBox.WordWrap   = $false
    $detailBox.BackColor  = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $detailBox.Text       = 'Select a row to see its full connection string.'
    $form.Controls.Add($detailBox)

    $listView.Add_SelectedIndexChanged({
        if ($listView.SelectedItems.Count -eq 0) { return }
        $detailBox.Text = [string]$listView.SelectedItems[0].Tag.Detail
    }.GetNewClosure())

    $copyButton = New-FormAutoButton -Text 'Copy report' -MinimumWidth 120
    $form.Controls.Add($copyButton)
    $copyButton.Location = New-Object System.Drawing.Point(12, 476)
    $copyButton.Anchor   = 'Left,Bottom'
    $copyButton.Add_Click({
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add(('Data source check - {0}' -f [string]$Rule.name))
        [void]$lines.Add((Get-Date).ToString('dd/MM/yyyy HH:mm'))
        [void]$lines.Add('')
        foreach ($row in $rows) {
            [void]$lines.Add(('[{0}] {1} / {2} / {3}' -f $row.Severity, $row.Workbook, $row.Connection, $row.Server))
            [void]$lines.Add(('    {0}' -f $row.Verdict))
        }
        try { [System.Windows.Forms.Clipboard]::SetText(($lines.ToArray() -join [Environment]::NewLine)) } catch { }
    }.GetNewClosure())

    $closeButton = New-FormAutoButton -Text 'Close' -MinimumWidth 100
    $form.Controls.Add($closeButton)
    $closeButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 12 - $closeButton.Width), 476)
    $closeButton.Anchor   = 'Right,Bottom'
    $closeButton.Add_Click({ $form.Close() }.GetNewClosure())
    $form.CancelButton = $closeButton

    if ($listView.Items.Count -gt 0) { $listView.Items[0].Selected = $true }

    Set-FormWithinWorkingArea -Form $form
    if ($null -ne $script:UiControls.Form) { [void]$form.ShowDialog($script:UiControls.Form) }
    else { [void]$form.ShowDialog() }
    $form.Dispose()
}
