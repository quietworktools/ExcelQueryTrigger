# ============================================================================== 
#  UISplash.ps1 - small, dependency-free startup progress window
# ============================================================================== 

Set-StrictMode -Version 1.0

function Show-StartupSplash {
    param([string]$AppRoot)

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Excel Query Trigger Manager'
    # This is intentionally roomier than a decorative splash. Startup can be
    # held by a slow network workbook, and the user needs to see which concrete
    # step or file is being handled without opening the log.
    $form.ClientSize      = New-Object System.Drawing.Size(640, 430)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.ControlBox      = $false
    $form.ShowInTaskbar   = $true
    # Keep startup progress visible even when launched by Setup or Explorer.
    $form.TopMost         = $true
    try { $form.Font = New-Object System.Drawing.Font('Segoe UI', 9) } catch { }

    $iconPath = Join-Path $AppRoot 'assets\ExcelQueryTrigger.ico'
    if (Test-Path -LiteralPath $iconPath) { try { $form.Icon = New-Object System.Drawing.Icon($iconPath) } catch { } }

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Excel Query Trigger Manager'
    $title.Location = New-Object System.Drawing.Point(24, 20)
    $title.Size = New-Object System.Drawing.Size(570, 30)
    try { $title.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 14, [System.Drawing.FontStyle]::Bold) } catch { }
    $form.Controls.Add($title)

    $message = New-Object System.Windows.Forms.Label
    $message.Text = 'Starting...'
    $message.Location = New-Object System.Drawing.Point(26, 66)
    $message.Size = New-Object System.Drawing.Size(588, 24)
    $form.Controls.Add($message)

    $detail = New-Object System.Windows.Forms.Label
    $detail.Text = ''
    $detail.Location = New-Object System.Drawing.Point(26, 92)
    $detail.Size = New-Object System.Drawing.Size(588, 38)
    $detail.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $form.Controls.Add($detail)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(26, 138)
    $progress.Size = New-Object System.Drawing.Size(588, 18)
    $progress.Minimum = 0; $progress.Maximum = 100; $progress.Value = 2
    $form.Controls.Add($progress)

    $activityLabel = New-Object System.Windows.Forms.Label
    $activityLabel.Text = 'Overall startup progress'
    $activityLabel.Location = New-Object System.Drawing.Point(26, 170)
    $activityLabel.Size = New-Object System.Drawing.Size(588, 20)
    try { $activityLabel.Font = New-Object System.Drawing.Font($form.Font.FontFamily, $form.Font.Size, [System.Drawing.FontStyle]::Bold) } catch { }
    $form.Controls.Add($activityLabel)

    $activity = New-Object System.Windows.Forms.TextBox
    $activity.Location = New-Object System.Drawing.Point(26, 194)
    $activity.Size = New-Object System.Drawing.Size(588, 150)
    $activity.Multiline = $true
    $activity.ReadOnly = $true
    $activity.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $activity.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $activity.Text = 'Starting the application...'
    $form.Controls.Add($activity)

    $startupNote = New-Object System.Windows.Forms.Label
    $startupNote.Text = 'The Dashboard opens only after monitoring is ready. Slow network locations may extend startup.'
    $startupNote.Location = New-Object System.Drawing.Point(26, 354)
    $startupNote.Size = New-Object System.Drawing.Size(588, 22)
    $startupNote.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $form.Controls.Add($startupNote)

    $version = New-Object System.Windows.Forms.Label
    $version.Text = ('Version {0}' -f (Get-AppVersion))
    $version.Location = New-Object System.Drawing.Point(26, 385)
    $version.Size = New-Object System.Drawing.Size(300, 20)
    $version.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $form.Controls.Add($version)

    $form.Tag = @{
        Message  = $message
        Detail   = $detail
        Progress          = $progress
        Activity          = $activity
        ShownAt           = Get-Date
    }
    $form.Show()
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    return $form
}

function Update-StartupSplash {
    param(
        $Splash,
        [string]$Message,
        [int]$Percent = -1,
        [string]$Detail = '',
        [AllowNull()][string]$ActivityText = $null
    )
    if ($null -eq $Splash -or $Splash.IsDisposed) { return }
    try {
        if (-not [string]::IsNullOrWhiteSpace($Message)) { $Splash.Tag.Message.Text = $Message }
        $Splash.Tag.Detail.Text = $Detail
        if ($null -ne $ActivityText) {
            $Splash.Tag.Activity.Text = $ActivityText
            $Splash.Tag.Activity.SelectionStart = $Splash.Tag.Activity.TextLength
            $Splash.Tag.Activity.ScrollToCaret()
        }
        if ($Percent -ge 0) { $Splash.Tag.Progress.Value = [Math]::Max(0, [Math]::Min(100, $Percent)) }
        $Splash.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch { }
}

function Close-StartupSplash {
    param($Splash)
    if ($null -eq $Splash) { return }
    try { if (-not $Splash.IsDisposed) { $Splash.Close(); $Splash.Dispose() } } catch { }
}
