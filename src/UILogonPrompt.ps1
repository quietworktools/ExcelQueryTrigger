# ==============================================================================
#  UILogonPrompt.ps1  (UI runspace only)
#
#  Straight after logon the machine is busy: OneDrive, Outlook, the security
#  agent and the network drives are all still waking up. Starting an Excel
#  refresh into that is exactly the wrong moment, so a rule set to run "at
#  logon" asks first by default.
#
#  The dialog does not refresh anything itself. It posts the same RunNow command
#  the Run Now button posts, so logon runs go through the existing queue and the
#  existing refresh engine like everything else.
# ==============================================================================

Set-StrictMode -Version 1.0

function Get-LogonRuleList {
    <#  Enabled rules whose trigger is "At Windows logon".  #>
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    $result = New-Object System.Collections.ArrayList
    foreach ($rule in @($Config.rules)) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (Test-TriggerIsLogon ([string]$rule.trigger.type)) { [void]$result.Add($rule) }
    }
    $found = @($result.ToArray())
    return ,$found
}

function Get-RuleWorkbookSummary {
    param([hashtable]$Rule)

    $names = New-Object System.Collections.ArrayList
    foreach ($action in @($Rule.actions)) {
        $path = [string]$action.path
        if (-not [string]::IsNullOrWhiteSpace($path)) { [void]$names.Add((Split-Path -Leaf $path)) }
    }
    if ($names.Count -eq 0) { return 'no workbooks configured' }
    return ($names -join ', ')
}

function Show-LogonRefreshPrompt {
    <#
        .SYNOPSIS
            Asks which logon rules to run. Returns the rule ids to run, or an
            empty array when the user skips.
        .DESCRIPTION
            Availability is checked after the window is on screen, one workbook
            at a time, because a probe against a share that is not connected yet
            can take a few seconds and must not delay the dialog appearing.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][array]$Rules = @()
    )

    if ($null -eq $Rules) { $Rules = @() }
    if ($Rules.Count -eq 0) { return ,@() }

    $font     = Get-UiFont
    $fontBold = Get-UiFont -Style 'Bold'

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Excel Query Trigger Manager'
    # Fixed-coordinate WinForms layout: Windows compatibility scaling is used
    # so fonts and controls scale together on high-DPI/Retina displays.
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize      = New-Object System.Drawing.Size(560, 400)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterScreen'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = $font
    $form.TopMost         = $true      # it appears while the user is elsewhere
    if ($null -ne $script:UiIcons -and $null -ne $script:UiIcons.Running) {
        try { $form.Icon = $script:UiIcons.Running } catch { }
    }

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text     = 'Excel refreshes are set up to run when you log in.'
    $lblTitle.Font     = $fontBold
    $lblTitle.Location = New-Object System.Drawing.Point(16, 16)
    $lblTitle.Size     = New-Object System.Drawing.Size(528, 24)
    $form.Controls.Add($lblTitle)

    $lblHint = New-FormWrappedLabel `
        -Text 'Tick the ones you want to run now. Nothing is refreshed until you choose.' `
        -X 16 -Y 40 -Width 528
    $form.Controls.Add($lblHint)

    $list = New-Object System.Windows.Forms.ListView
    $list.Location      = New-Object System.Drawing.Point(16, 68)
    $list.Size          = New-Object System.Drawing.Size(528, 240)
    $list.View          = 'Details'
    $list.CheckBoxes    = $true
    $list.FullRowSelect = $true
    $list.GridLines     = $true
    $list.HideSelection = $false
    $list.ShowItemToolTips = $true
    [void]$list.Columns.Add('Rule', 190)
    [void]$list.Columns.Add('Workbooks', 220)
    [void]$list.Columns.Add('Available', 110)
    $form.Controls.Add($list)

    foreach ($rule in $Rules) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$rule.name)
        [void]$item.SubItems.Add((Get-RuleWorkbookSummary -Rule $rule))
        [void]$item.SubItems.Add('Checking...')
        $item.Checked     = $true
        $item.Tag         = $rule
        $item.ToolTipText = (@(@($rule.actions) | ForEach-Object { [string]$_.path }) -join [Environment]::NewLine)
        [void]$list.Items.Add($item)
    }

    $buttons = @{}
    $x = 16
    foreach ($spec in @(
            @{ Key = 'Selected'; Text = 'Refresh Selected'; Width = 140 },
            @{ Key = 'All';      Text = 'Refresh All';      Width = 120 },
            @{ Key = 'Skip';     Text = 'Skip';             Width = 100 })) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text     = $spec.Text
        $button.Location = New-Object System.Drawing.Point($x, 322)
        $button.Size     = New-Object System.Drawing.Size($spec.Width, 34)
        $form.Controls.Add($button)
        $buttons[$spec.Key] = $button
        $x += $spec.Width + 8
    }
    $form.AcceptButton = $buttons['Selected']
    $form.CancelButton = $buttons['Skip']

    $lblNote = New-FormWrappedLabel `
        -Text 'Skipping only affects this logon. Folder monitoring and scheduled rules keep running.' `
        -X 16 -Y 364 -Width 528
    $form.Controls.Add($lblNote)

    $selection = New-Object System.Collections.ArrayList

    # ---- availability probe, one item per tick ------------------------------
    $probeIndex = 0
    $probeTimer = New-Object System.Windows.Forms.Timer
    $probeTimer.Interval = 60
    $probeTimer.Add_Tick({
        if ($probeIndex -ge $list.Items.Count) {
            $probeTimer.Stop()
            return
        }
        $item = $list.Items[$probeIndex]
        $probeIndex++

        $missing = 0
        $total   = 0
        foreach ($action in @($item.Tag.actions)) {
            $path = [string]$action.path
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $total++
            $exists = $false
            try { $exists = Test-Path -LiteralPath $path } catch { $exists = $false }
            if (-not $exists) { $missing++ }
        }

        if ($total -eq 0) {
            $item.SubItems[2].Text = 'Nothing to do'
            $item.Checked = $false
        }
        elseif ($missing -eq 0) {
            $item.SubItems[2].Text = 'Ready'
        }
        elseif ($missing -eq $total) {
            # Not an error: the drive is often simply not connected yet. The rule
            # stays selectable, it is just unticked so it is a deliberate choice.
            $item.SubItems[2].Text = 'Not available'
            $item.ForeColor = [System.Drawing.Color]::FromArgb(180, 60, 40)
            $item.Checked   = $false
        }
        else {
            $item.SubItems[2].Text = ('{0} of {1} missing' -f $missing, $total)
            $item.ForeColor = [System.Drawing.Color]::FromArgb(180, 60, 40)
        }
    }.GetNewClosure())

    $form.Add_Shown({ $probeTimer.Start() }.GetNewClosure())

    $buttons['Selected'].Add_Click({
        foreach ($item in $list.Items) {
            if ($item.Checked) { [void]$selection.Add([string]$item.Tag.id) }
        }
        $form.Close()
    }.GetNewClosure())

    $buttons['All'].Add_Click({
        foreach ($item in $list.Items) { [void]$selection.Add([string]$item.Tag.id) }
        $form.Close()
    }.GetNewClosure())

    $buttons['Skip'].Add_Click({ $form.Close() }.GetNewClosure())

    try {
        Set-FormWithinWorkingArea -Form $form
        [void]$form.ShowDialog()
    }
    finally {
        try { $probeTimer.Stop() } catch { }
        try { $probeTimer.Dispose() } catch { }
        $form.Dispose()
    }

    $chosen = @($selection.ToArray())
    return ,$chosen
}

function Invoke-LogonRefreshDecision {
    <#
        .SYNOPSIS
            Runs once, some seconds after a logon start: queues the rules set to
            refresh automatically, and asks about the rest.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$Config
    )

    $rules = Get-LogonRuleList -Config $Config
    if ($rules.Count -eq 0) { return }

    $automatic = @($rules | Where-Object { [string]$_.trigger.logonBehavior -eq 'Automatic' })
    $ask       = @($rules | Where-Object { [string]$_.trigger.logonBehavior -ne 'Automatic' })

    foreach ($rule in $automatic) {
        Write-AppLog -Level 'INFO' -RuleName $rule.name -Message 'Logon rule set to refresh automatically.'
        Send-EngineCommand -Shared $Shared -Type 'RunNow' -Payload @{
            RuleId = [string]$rule.id; Source = 'Logon'; CheckRecentRefresh = $true
        }
    }

    if ($ask.Count -eq 0) { return }

    $chosen = Show-LogonRefreshPrompt -Rules $ask
    if ($chosen.Count -eq 0) {
        Write-AppLog -Level 'INFO' -Message 'Logon refresh prompt: nothing selected. Monitoring continues.'
        return
    }

    foreach ($ruleId in $chosen) {
        Send-EngineCommand -Shared $Shared -Type 'RunNow' -Payload @{ RuleId = $ruleId; Source = 'Logon' }
    }
    Write-AppLog -Level 'INFO' -Message ('Logon refresh prompt: {0} rule(s) queued.' -f $chosen.Count)
}
