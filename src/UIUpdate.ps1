# ============================================================================== 
#  UIUpdate.ps1  (UI runspace only)
#
#  GitHub is used for one purpose only: reading the newest public release of
#  Excel Query Trigger Manager. No token, issue, feedback or write API exists in
#  the application.
# ============================================================================== 

Set-StrictMode -Version 1.0

$script:UpdateRepository = 'quietworktools/ExcelQueryTrigger'
$script:UpdateLastCheck   = [DateTime]::MinValue
$script:UpdateCheckTask   = $null

function Get-UpdateHeaders {
    return @{
        'Accept'               = 'application/vnd.github+json'
        'User-Agent'           = ('ExcelQueryTrigger/{0}' -f (Get-AppVersion))
        'X-GitHub-Api-Version' = '2022-11-28'
    }
}

function ConvertTo-ComparableVersion {
    param([string]$Text)

    $match = [regex]::Match(([string]$Text).Trim(), '^v?(\d+(?:\.\d+){0,3})$')
    if (-not $match.Success) { return $null }
    try { return [version]$match.Groups[1].Value } catch { return $null }
}

function Get-NormalizedVersionText {
    param([string]$Text)
    $version = ConvertTo-ComparableVersion $Text
    if ($null -eq $version) { return '' }
    return $version.ToString()
}

function ConvertFrom-GitHubRelease {
    <#  Converts a GitHub release response into the small object the UI needs. #>
    param($Release)

    $result = @{ Ok = $false; Available = $false; Version = ''; Notes = ''; ZipUrl = ''; ChecksumUrl = ''; PageUrl = ''; AssetName = ''; Message = '' }
    if ($null -eq $Release) {
        $result.Message = 'The update service returned no release information.'
        return $result
    }

    $tag     = [string]$Release.tag_name
    $latest  = ConvertTo-ComparableVersion $tag
    $current = ConvertTo-ComparableVersion (Get-AppVersion)
    if ($null -eq $latest) {
        $result.Message = ('The latest release tag "{0}" is not a version number.' -f $tag)
        return $result
    }
    if ($null -eq $current) {
        $result.Message = 'This copy has an invalid internal version number.'
        return $result
    }

    $result.Ok        = $true
    $result.Version   = $tag
    $result.Notes     = [string]$Release.body
    $result.PageUrl   = [string]$Release.html_url
    $result.Available = ($latest -gt $current)

    $normalized = Get-NormalizedVersionText $tag
    $expectedName = 'ExcelQueryTrigger-v{0}.zip' -f $normalized
    foreach ($asset in @($Release.assets)) {
    $assetName = [string]$asset.name
    if ([string]::Equals($assetName, $expectedName, [StringComparison]::OrdinalIgnoreCase)) {
        $result.AssetName = $assetName
        $result.ZipUrl    = [string]$asset.browser_download_url
    }
    elseif ([string]::Equals($assetName, 'SHA256SUMS.txt', [StringComparison]::OrdinalIgnoreCase)) {
        $result.ChecksumUrl = [string]$asset.browser_download_url
    }
}

if (-not $result.Available) { $result.Message = 'This is the newest version.' }
elseif ([string]::IsNullOrWhiteSpace($result.ZipUrl)) {
    $result.Message = ('Version {0} is available, but its expected package ({1}) is not attached.' -f $tag, $expectedName)
}
elseif ([string]::IsNullOrWhiteSpace($result.ChecksumUrl)) {
    $result.Message = ('Version {0} is available, but SHA256SUMS.txt is not attached.' -f $tag)
}
return $result
}

function Test-TrustedReleaseUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    try { $uri = [Uri]$Url } catch { return $false }
    if ($uri.Scheme -ne 'https') { return $false }
    return ($uri.Host -in @('github.com', 'objects.githubusercontent.com', 'release-assets.githubusercontent.com'))
}

function ConvertTo-ReleaseNotesDisplayText {
    <#
        A WinForms TextBox needs CRLF line endings. GitHub returns LF-only
        Markdown, which otherwise appears as one long paragraph on Windows.
        Render the small Markdown subset used by RELEASE-NOTES.md as readable
        plain text: headings become text and list markers become real bullets.
    #>
    param([string]$Markdown)

    if ([string]::IsNullOrWhiteSpace($Markdown)) { return 'No release notes were provided.' }

    $output = New-Object System.Collections.ArrayList
    $previousBlank = $true
    $bulletPrefix = ([char]0x2022).ToString() + ' '
    foreach ($sourceLine in @($Markdown -split "`r?`n")) {
        $line = [string]$sourceLine
        $line = $line.TrimEnd()
        if ([string]::IsNullOrWhiteSpace($line)) {
            if (-not $previousBlank) { [void]$output.Add('') }
            $previousBlank = $true
            continue
        }

        $line = $line -replace '^\s*#{1,6}\s+', ''
        $line = $line -replace '^\s*[-*+]\s+', $bulletPrefix
        $line = $line -replace '\*\*([^*]+)\*\*', '$1'
        $line = $line -replace '`([^`]+)`', '$1'
        $line = $line -replace '\[([^\]]+)\]\([^\)]+\)', '$1'
        [void]$output.Add($line)
        $previousBlank = $false
    }
    while ($output.Count -gt 0 -and [string]::IsNullOrWhiteSpace([string]$output[$output.Count - 1])) {
        $output.RemoveAt($output.Count - 1)
    }
    return (@($output.ToArray()) -join [Environment]::NewLine)
}

function Get-UpdateProxyArgs {
    <#
        Windows PowerShell does not authenticate to a corporate proxy on its
        own, and most offices have one in the way. Nothing is passed when there
        is no proxy, because -Proxy with an empty value fails outright.
    #>
    param([string]$Uri = 'https://api.github.com')

    try {
        $proxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy([Uri]$Uri)
        if ($null -ne $proxyUri -and $proxyUri.AbsoluteUri -ne ([Uri]$Uri).AbsoluteUri) {
            return @{ Proxy = $proxyUri; ProxyUseDefaultCredentials = $true }
        }
    }
    catch { }
    return @{}
}

function Get-UpdateFailureText {
    <#
        Why the check did not answer, said usefully.

        A bare 403 is nearly always the hourly limit rather than a refusal:
        unauthenticated calls get sixty an hour per public address, and an
        office shares one address. GitHub reports which it is in
        X-RateLimit-Remaining and in the response body, so both are read
        instead of guessed at, and the raw answer goes to the log so a report
        afterwards can say what actually happened.
    #>
    param($ErrorRecord)

    $status    = 0
    $body      = ''
    $remaining = ''
    $response  = $null
    try { $response = $ErrorRecord.Exception.Response } catch { $response = $null }

    if ($null -ne $response) {
        try { $status = [int]$response.StatusCode } catch { }
        try { $remaining = [string]$response.Headers['X-RateLimit-Remaining'] } catch { }
        try {
            $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
            $body = $reader.ReadToEnd()
            $reader.Dispose()
        }
        catch { }
    }

    $limited = ($remaining -eq '0') -or ($body -match '(?i)rate limit')

    $message =
        if ($limited)            { 'This network has checked for updates too many times in the last hour. The limit is shared by everyone in the office and it clears within the hour.' }
        elseif ($status -eq 404) { 'No release has been published yet.' }
        elseif ($status -eq 403) { 'The update service refused the request. On an office network this is usually the proxy.' }
        elseif ($status -eq 401) { 'The update service did not accept this copy of the application.' }
        elseif ($status -ge 500) { 'The update service is having trouble at the moment. Try again later.' }
        else                     { 'The update service could not be reached. There may be no internet access from here.' }

    try {
        Write-AppLog -Level 'WARN' -Message ('Update check failed. HTTP {0}. {1} {2}' -f `
            $status, [string]$ErrorRecord.Exception.Message, $body.Trim())
    }
    catch { }

    return $message
}

function Get-UpdateInfo {
    <#  Manual checks may wait; automatic checks use the asynchronous path below. #>
    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch { }
    try {
        $proxyArgs = Get-UpdateProxyArgs
        $release = Invoke-RestMethod @proxyArgs -Method Get `
            -Uri ('https://api.github.com/repos/{0}/releases/latest' -f $script:UpdateRepository) `
            -Headers (Get-UpdateHeaders) -TimeoutSec 20 -ErrorAction Stop
        return (ConvertFrom-GitHubRelease -Release $release)
    }
    catch {
        return @{ Ok = $false; Available = $false; Version = ''; Notes = ''; ZipUrl = ''
            PageUrl = ''; AssetName = ''; Message = (Get-UpdateFailureText -ErrorRecord $_) }
    }
}

function Get-PackageVersion {
    param([string]$Root)
    $common = Join-Path $Root 'src\Common.ps1'
    if (-not (Test-Path -LiteralPath $common -PathType Leaf)) { return '' }
    try {
        $match = [regex]::Match(([System.IO.File]::ReadAllText($common)), "AppVersion\s*=\s*'([^']+)'")
        if ($match.Success) { return [string]$match.Groups[1].Value }
    }
    catch { }
    return ''
}

function Find-ValidatedUpdateRoot {
    param([string]$Unpacked, [string]$ExpectedVersion)

    $roots = New-Object System.Collections.ArrayList
    [void]$roots.Add($Unpacked)
    foreach ($directory in @(Get-ChildItem -LiteralPath $Unpacked -Directory -ErrorAction SilentlyContinue)) {
        [void]$roots.Add($directory.FullName)
    }

    $candidates = New-Object System.Collections.ArrayList
    foreach ($root in @($roots.ToArray())) {
        $required = @('Setup.cmd', 'ExcelQueryTrigger.ps1', 'tools\Setup.ps1', 'src\Common.ps1')
        $complete = $true
        foreach ($relative in $required) {
            if (-not (Test-Path -LiteralPath (Join-Path $root $relative) -PathType Leaf)) { $complete = $false; break }
        }
        if ($complete) { [void]$candidates.Add($root) }
    }
    if ($candidates.Count -ne 1) { return @{ Ok = $false; Root = ''; Message = 'The downloaded package does not have one valid application root.' } }

    $root = [string]$candidates[0]
    $setups = @(Get-ChildItem -LiteralPath $Unpacked -Recurse -Filter 'Setup.cmd' -File -ErrorAction SilentlyContinue)
    if ($setups.Count -ne 1 -or -not [string]::Equals($setups[0].FullName, (Join-Path $root 'Setup.cmd'), [StringComparison]::OrdinalIgnoreCase)) {
        return @{ Ok = $false; Root = ''; Message = 'The downloaded package contains an unexpected Setup.cmd.' }
    }

    $packageVersion = Get-PackageVersion -Root $root
    $expected = Get-NormalizedVersionText $ExpectedVersion
    $actual   = Get-NormalizedVersionText $packageVersion
    if ($actual -eq '' -or $actual -ne $expected) {
        return @{ Ok = $false; Root = ''; Message = ('The package says it is version {0}, not {1}.' -f $packageVersion, $ExpectedVersion) }
    }
    return @{ Ok = $true; Root = $root; Message = '' }
}

function Install-Update {
    <#  Downloads and validates the exact release package, then starts its setup.
        Setup waits for this process to exit; it never kills the running app. #>
    param([hashtable]$Update)

    if ([string]::IsNullOrWhiteSpace([string]$Update.ZipUrl) -or [string]::IsNullOrWhiteSpace([string]$Update.ChecksumUrl)) {
    try { Start-Process ([string]$Update.PageUrl) | Out-Null } catch { }
    return @{ Ok = $false; Message = [string]$Update.Message }
}
if (-not (Test-TrustedReleaseUrl -Url ([string]$Update.ZipUrl)) -or
    -not (Test-TrustedReleaseUrl -Url ([string]$Update.ChecksumUrl))) {
    return @{ Ok = $false; Message = 'The update service returned an unexpected download URL.' }
}

$staging = Join-Path ([System.IO.Path]::GetTempPath()) ('eqt-update-{0}' -f ([guid]::NewGuid().ToString('N')))
$zipPath = Join-Path $staging 'update.zip'
$checksumPath = Join-Path $staging 'SHA256SUMS.txt'
    try {
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
        try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch { }
        $proxyArgs = Get-UpdateProxyArgs -Uri ([string]$Update.ZipUrl)
    Invoke-WebRequest @proxyArgs -UseBasicParsing -Uri ([string]$Update.ZipUrl) -OutFile $zipPath `
        -Headers (Get-UpdateHeaders) -TimeoutSec 180 -ErrorAction Stop
    $checksumProxyArgs = Get-UpdateProxyArgs -Uri ([string]$Update.ChecksumUrl)
    Invoke-WebRequest @checksumProxyArgs -UseBasicParsing -Uri ([string]$Update.ChecksumUrl) -OutFile $checksumPath `
        -Headers (Get-UpdateHeaders) -TimeoutSec 60 -ErrorAction Stop

    $checksumText = [System.IO.File]::ReadAllText($checksumPath)
    $expectedAssetName = [string]$Update.AssetName
    $escapedAssetName = [regex]::Escape($expectedAssetName)
    $checksumMatch = [regex]::Match($checksumText, ('(?im)^([0-9a-f]{{64}})\s+\*?{0}\s*$' -f $escapedAssetName))
    if (-not $checksumMatch.Success) {
        throw [System.InvalidOperationException]::new('SHA256SUMS.txt does not contain the expected release package hash.')
    }
    $expectedHash = $checksumMatch.Groups[1].Value.ToLowerInvariant()
    $actualHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not [string]::Equals($expectedHash, $actualHash, [StringComparison]::OrdinalIgnoreCase)) {
        throw [System.Security.SecurityException]::new('The downloaded update package failed SHA-256 verification.')
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $unpacked = Join-Path $staging 'files'
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $unpacked)
        $validated = Find-ValidatedUpdateRoot -Unpacked $unpacked -ExpectedVersion ([string]$Update.Version)
        if (-not $validated.Ok) { throw [System.InvalidOperationException]::new([string]$validated.Message) }

        try { Get-ChildItem -LiteralPath $validated.Root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue } catch { }
        Write-AppLog -Level 'INFO' -Message ('Update {0} downloaded and SHA-256 verified ({1}).' -f [string]$Update.Version, $actualHash)

        $setupScript = Join-Path ([string]$validated.Root) 'tools\Setup.ps1'
        $powershellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}" -WaitForPid {1} -UpdateMode' -f $setupScript, $PID
        Start-Process -FilePath $powershellExe -ArgumentList $arguments -WorkingDirectory ([string]$validated.Root) | Out-Null
        return @{ Ok = $true; Message = 'The verified installer will open as soon as this copy closes.' }
    }
    catch {
        try { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
        return @{ Ok = $false; Message = ('The update could not be prepared: {0}' -f $_.Exception.Message) }
    }
}

function Show-AboutDialog {
    param($Owner = $null)

    $form = New-Object System.Windows.Forms.Form
    $form.Text                = 'About'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize          = New-Object System.Drawing.Size(560, 476)
    $form.FormBorderStyle     = 'FixedDialog'
    $form.StartPosition       = 'CenterParent'
    $form.MaximizeBox         = $false
    $form.MinimizeBox         = $false
    $form.Font                = (Get-UiFont)
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = 'Excel Query Trigger Manager'; $lblName.Font = (Get-UiFont 13 'Bold')
    $lblName.Location = New-Object System.Drawing.Point(22, 18); $lblName.Size = New-Object System.Drawing.Size(516, 30)
    $form.Controls.Add($lblName)

    $versionGroup = New-Object System.Windows.Forms.GroupBox
    $versionGroup.Text = 'Version and updates'
    $versionGroup.Location = New-Object System.Drawing.Point(22, 56)
    $versionGroup.Size = New-Object System.Drawing.Size(516, 144)
    $form.Controls.Add($versionGroup)

    $lblVersionCaption = New-Object System.Windows.Forms.Label
    $lblVersionCaption.Text = 'Installed version'
    $lblVersionCaption.Location = New-Object System.Drawing.Point(18, 24)
    $lblVersionCaption.Size = New-Object System.Drawing.Size(190, 18)
    $lblVersionCaption.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
    $versionGroup.Controls.Add($lblVersionCaption)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = (Get-AppVersion); $lblVersion.Font = (Get-UiFont 15 'Bold')
    $lblVersion.Location = New-Object System.Drawing.Point(18, 43)
    $lblVersion.Size = New-Object System.Drawing.Size(190, 30); $versionGroup.Controls.Add($lblVersion)

    $btnCheck = New-FormAutoButton -Text 'Check for updates' -MinimumWidth 160
    $btnCheck.Location = New-Object System.Drawing.Point(330, 28); $versionGroup.Controls.Add($btnCheck)

    $lblStatus = New-Object System.Windows.Forms.TextBox
    $lblStatus.Location = New-Object System.Drawing.Point(18, 76); $lblStatus.Size = New-Object System.Drawing.Size(480, 56)
    $lblStatus.Multiline = $true; $lblStatus.ReadOnly = $true; $lblStatus.WordWrap = $true
    $lblStatus.BorderStyle = 'None'; $lblStatus.ScrollBars = 'Vertical'; $lblStatus.BackColor = $versionGroup.BackColor
    $lblStatus.Text = 'Update status has not been checked yet.'
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105); $versionGroup.Controls.Add($lblStatus)

    $notesGroup = New-Object System.Windows.Forms.GroupBox
    $notesGroup.Text = 'Release notes'
    $notesGroup.Location = New-Object System.Drawing.Point(22, 212)
    $notesGroup.Size = New-Object System.Drawing.Size(516, 190)
    $form.Controls.Add($notesGroup)

    $txtNotes = New-Object System.Windows.Forms.TextBox
    $txtNotes.Location = New-Object System.Drawing.Point(12, 24); $txtNotes.Size = New-Object System.Drawing.Size(492, 152)
    $txtNotes.Multiline = $true; $txtNotes.ReadOnly = $true; $txtNotes.ScrollBars = 'Vertical'
    $txtNotes.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $txtNotes.Text = 'Release notes will appear here when a newer version is available.'
    $notesGroup.Controls.Add($txtNotes)

    $btnInstall = New-FormAutoButton -Text 'Install it' -MinimumWidth 120
    $btnInstall.Location = New-Object System.Drawing.Point(22, 422); $btnInstall.Visible = $false; $form.Controls.Add($btnInstall)
    $btnClose = New-FormAutoButton -Text 'Close' -MinimumWidth 90
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnClose.Location = New-Object System.Drawing.Point(448, 422)
    $form.Controls.Add($btnClose); $form.CancelButton = $btnClose

    # Each GetNewClosure() below gets its own dynamic module.  A script-scoped
    # scalar therefore is not shared between the Check and Install handlers.
    # Keep the selected update in one reference object captured by both.
    $aboutState = @{ Update = $null }
    $btnCheck.Add_Click({
        $previousCursor = [System.Windows.Forms.Cursor]::Current
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = 'Checking...'; [System.Windows.Forms.Application]::DoEvents()
        try { $update = Get-UpdateInfo } finally { [System.Windows.Forms.Cursor]::Current = $previousCursor }
        $aboutState.Update = $update
        if (-not $update.Ok) {
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170,110,10)
            $lblStatus.Text = [string]$update.Message
            $txtNotes.Text = 'Release information could not be loaded.'
            $btnInstall.Visible = $false
            return
        }
        if (-not $update.Available) {
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30,130,85)
            $lblStatus.Text = 'This is the newest version.'
            $txtNotes.Text = ('Version {0} is currently installed. No update is required.' -f (Get-AppVersion))
            $btnInstall.Visible = $false
            return
        }
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(24,95,165); $lblStatus.Text = ('Version {0} is available.' -f [string]$update.Version)
        $txtNotes.Text = ConvertTo-ReleaseNotesDisplayText -Markdown ([string]$update.Notes)
        $btnInstall.Visible = (-not [string]::IsNullOrWhiteSpace([string]$update.ZipUrl))
        if (-not $btnInstall.Visible) { $lblStatus.Text = [string]$update.Message }
    }.GetNewClosure())

    $btnInstall.Add_Click({
        if ($null -eq $aboutState.Update) {
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170,110,10)
            $lblStatus.Text = 'Check for updates again before installing.'
            return
        }
        if ([bool]$script:UiShared.CurrentJob.Active) {
            [System.Windows.Forms.MessageBox]::Show('A workbook refresh is running. Install the update after it finishes.', 'Install update', 'OK', 'Information') | Out-Null
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ('Install version {0}?' -f [string]$aboutState.Update.Version) + [Environment]::NewLine + [Environment]::NewLine +
            'The package is downloaded and checked first. Your rules, history and logs are kept.' + [Environment]::NewLine +
            'After that, this app closes, Setup finishes automatically, and the updated Dashboard opens.',
            'Install update', 'YesNo', 'Question')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $previousCursor = [System.Windows.Forms.Cursor]::Current
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = 'Downloading and checking the package...'; [System.Windows.Forms.Application]::DoEvents()
        try { $outcome = Install-Update -Update $aboutState.Update } finally { [System.Windows.Forms.Cursor]::Current = $previousCursor }
        if (-not $outcome.Ok) { $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170,110,10); $lblStatus.Text = [string]$outcome.Message; return }
        # The user already approved the update above. Do not add another OK
        # dialog between download and Setup: close this window and the Dashboard
        # so the automated updater can replace the running files immediately.
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30,130,85)
        $lblStatus.Text = 'Ready. Closing this version and installing automatically...'
        [System.Windows.Forms.Application]::DoEvents()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
        Stop-Application
    }.GetNewClosure())

    Set-FormWithinWorkingArea -Form $form
    if ($null -eq $Owner) { $Owner = $script:UiControls.Form }
    [void]$form.ShowDialog($Owner)
    $form.Dispose()
}

function Update-VersionLabel {
    param([hashtable]$Update)
    $label = $script:UiControls.Version
    if ($null -eq $label) { return }
    if ($null -ne $Update -and $Update.Ok -and $Update.Available) {
        $label.Text = ('Version {0}   -   {1} is available' -f (Get-AppVersion), [string]$Update.Version)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(24,95,165); $label.Font = (Get-UiFont 9 'Bold')
    }
    else {
        $label.Text = ('Version {0}' -f (Get-AppVersion)); $label.ForeColor = [System.Drawing.Color]::FromArgb(115,115,115)
        $label.Font = (Get-UiFont 9 'Regular')
    }
}

function Get-LastUpdateCheck {
    $value = ''
    try { $value = [string]$script:UiConfig.appSettings.lastUpdateCheckUtc } catch { $value = '' }
    $parsed = [DateTime]::MinValue
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        try { $parsed = [DateTime]::Parse($value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind) } catch { $parsed = [DateTime]::MinValue }
    }
    return $parsed
}

function Save-LastUpdateCheck {
    $now = [DateTime]::UtcNow
    $script:UpdateLastCheck = $now
    try {
        $script:UiConfig.appSettings.lastUpdateCheckUtc = $now.ToString('o')
        Save-AppConfiguration -Path $script:UiPaths.ConfigPath -Config $script:UiConfig
    }
    catch { }
}

function Invoke-BackgroundUpdateCheck {
    <#  Called by the UI heartbeat. It starts once, then only polls the handle. #>
    $automatic = $true
    try { $automatic = ConvertTo-BoolValue $script:UiConfig.appSettings.checkForUpdatesAutomatically $true } catch { }
    if (-not $automatic) {
        Stop-BackgroundUpdateCheck
        return
    }

    if ($null -ne $script:UpdateCheckTask) {
        if (-not $script:UpdateCheckTask.Handle.IsCompleted) { return }
        $task = $script:UpdateCheckTask; $script:UpdateCheckTask = $null
        try {
            $output = @($task.PowerShell.EndInvoke($task.Handle))
            if ($output.Count -gt 0 -and [bool]$output[0].Ok) {
                $update = ConvertFrom-GitHubRelease -Release $output[0].Release
                if ($update.Ok -and $update.Available) {
                    Write-AppLog -Level 'INFO' -Message ('Version {0} is available. Click the version number at the bottom left.' -f [string]$update.Version)
                }
                Update-VersionLabel -Update $update
            }
        }
        catch { }
        finally { try { $task.PowerShell.Dispose() } catch { } }
        return
    }

    $last = Get-LastUpdateCheck
    if ($last -gt [DateTime]::MinValue -and ([DateTime]::UtcNow - $last.ToUniversalTime()).TotalHours -lt 24) { return }
    Save-LastUpdateCheck

    $uri = 'https://api.github.com/repos/{0}/releases/latest' -f $script:UpdateRepository
    $headers = Get-UpdateHeaders
    $proxyArgs = Get-UpdateProxyArgs -Uri $uri
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
        param($RequestUri, $RequestHeaders, $RequestProxyArgs)
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $release = Invoke-RestMethod @RequestProxyArgs -Method Get -Uri $RequestUri `
                -Headers $RequestHeaders -TimeoutSec 20 -ErrorAction Stop
            [pscustomobject]@{ Ok = $true; Release = $release; Message = '' }
        }
        catch { [pscustomobject]@{ Ok = $false; Release = $null; Message = [string]$_.Exception.Message } }
    }).AddArgument($uri).AddArgument($headers).AddArgument($proxyArgs)
    $script:UpdateCheckTask = @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}

function Stop-BackgroundUpdateCheck {
    if ($null -eq $script:UpdateCheckTask) { return }
    try { $script:UpdateCheckTask.PowerShell.Stop() } catch { }
    try { $script:UpdateCheckTask.PowerShell.Dispose() } catch { }
    $script:UpdateCheckTask = $null
}
