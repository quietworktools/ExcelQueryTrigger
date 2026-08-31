# ==============================================================================
#  UISettings.ps1  (UI runspace only)
#
#  The settings a person actually changes are few; the rest are knobs that exist
#  because the engine needs a number somewhere. This dialog separates the two:
#
#     Basics         what happens when Windows starts, and where the window goes
#     Notifications  what it tells you, and when
#     Advanced       the timing numbers, macro policy, logging, folders
#
#  Every setting carries a plain sentence underneath saying what it does and
#  what happens if you leave it alone. The keys and the save behaviour are
#  unchanged - only the presentation and the wording.
#
#  "Start with Windows" is applied immediately, because the registry write can
#  fail under a locked-down policy and that has to be visible now rather than at
#  the next sign-in.
# ==============================================================================

Set-StrictMode -Version 1.0

function New-SettingCheck {
    <#  A checkbox with its explanation underneath. Returns the checkbox.  #>
    param(
        [Parameter(Mandatory = $true)]$Parent,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Explanation,
        [Parameter(Mandatory = $true)][int]$Y,
        [bool]$Checked = $false
    )

    $check = New-Object System.Windows.Forms.CheckBox
    $check.Text     = $Text
    $check.Location = New-Object System.Drawing.Point(16, $Y)
    $check.Size     = New-Object System.Drawing.Size(468, 22)
    $check.Checked  = $Checked
    $Parent.Controls.Add($check)

    $note = New-FormWrappedLabel -Text $Explanation -X 35 -Y ($Y + 26) -Width 445
    $Parent.Controls.Add($note)

    return $check
}

function New-SettingNumber {
    <#  A labelled number box with its unit and an explanation underneath.  #>
    param(
        [Parameter(Mandatory = $true)]$Parent,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Unit,
        [Parameter(Mandatory = $true)][string]$Explanation,
        [Parameter(Mandatory = $true)][int]$Y,
        [int]$Min = 0,
        [int]$Max = 3600,
        [int]$Value = 0
    )

    $Parent.Controls.Add((New-FormLabel $Text 16 ($Y + 3) 300))
    $numeric = New-FormNumeric 322 $Y 84 $Min $Max $Value
    $Parent.Controls.Add($numeric)
    $Parent.Controls.Add((New-FormLabel $Unit 412 ($Y + 3) 72))

    $note = New-FormWrappedLabel -Text $Explanation -X 18 -Y ($Y + 28) -Width 466
    $Parent.Controls.Add($note)

    return $numeric
}

function Show-SettingsDialog {
    <#  Returns the updated appSettings hashtable, or $null on cancel.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$AppSettings,
        [Parameter(Mandatory = $true)][hashtable]$Paths
    )

    $working  = Merge-DefaultValues -Default (Get-DefaultAppSettings) -Value $AppSettings
    $defaults = Get-DefaultAppSettings

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Settings'
    # Fixed-coordinate WinForms layout: Windows compatibility scaling is used
    # so fonts and controls scale together on high-DPI/Retina displays.
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize      = New-Object System.Drawing.Size(540, 560)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterParent'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = (Get-UiFont)
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, 12)
    $tabs.Size     = New-Object System.Drawing.Size(516, 456)
    $form.Controls.Add($tabs)

    $tabBasics   = New-Object System.Windows.Forms.TabPage; $tabBasics.Text   = '  Basics  '
    $tabNotify   = New-Object System.Windows.Forms.TabPage; $tabNotify.Text   = '  Notifications  '
    $tabAdv      = New-Object System.Windows.Forms.TabPage; $tabAdv.Text      = '  Advanced  '
    $tabs.TabPages.AddRange(@($tabBasics, $tabNotify, $tabAdv))
    foreach ($page in @($tabBasics, $tabNotify, $tabAdv)) {
        $page.BackColor  = [System.Drawing.SystemColors]::Window
        $page.AutoScroll = $true
    }

    $checkBoxes = @{}
    $numerics   = @{}

    # ------------------------------------------------------------- Basics ----

    $checkBoxes['startWithWindows'] = New-SettingCheck -Parent $tabBasics -Y 16 `
        -Text 'Start it when I sign in to Windows' `
        -Explanation 'Recommended. It cannot refresh anything while it is closed, and it does not catch up on what it missed.' `
        -Checked (ConvertTo-BoolValue $working['startWithWindows'] $false)

    $checkBoxes['startMinimized'] = New-SettingCheck -Parent $tabBasics -Y 94 `
        -Text 'Start without opening the window' `
        -Explanation 'It starts quietly and gets on with it. Turn this off if you would rather see the dashboard every time.' `
        -Checked (ConvertTo-BoolValue $working['startMinimized'] $true)

    $checkBoxes['minimizeToTray'] = New-SettingCheck -Parent $tabBasics -Y 172 `
        -Text 'Keep it by the clock instead of on the taskbar' `
        -Explanation 'Closing or minimising the window leaves the icon by the clock. Click that icon once to bring the window back.' `
        -Checked (ConvertTo-BoolValue $working['minimizeToTray'] $true)

    $checkBoxes['checkForUpdatesAutomatically'] = New-SettingCheck -Parent $tabBasics -Y 250 `
        -Text 'Automatically check for updates' `
        -Explanation 'Recommended. Once a day, it quietly checks the latest public GitHub release. Turn this off to use only the manual Check for updates button.' `
        -Checked (ConvertTo-BoolValue $working['checkForUpdatesAutomatically'] $true)

    $tabBasics.Controls.Add((New-FormLabel 'Text style' 16 336 120))
    $lblFontIntro = New-FormWrappedLabel `
        -Text 'Leave this alone unless the text is hard to read on your screen. The size is fixed, because the layout is built around it.' `
        -X 18 -Y 358 -Width 466
    $tabBasics.Controls.Add($lblFontIntro)

    $cboFont = New-Object System.Windows.Forms.ComboBox
    $cboFont.Location      = New-Object System.Drawing.Point(18, 418)
    $cboFont.Size          = New-Object System.Drawing.Size(250, 24)
    $cboFont.DropDownStyle = 'DropDownList'
    [void]$cboFont.Items.Add('Automatic (recommended)')
    try {
        $fontCollection = New-Object System.Drawing.Text.InstalledFontCollection
        foreach ($fontName in @($fontCollection.Families | ForEach-Object { $_.Name } | Sort-Object -Unique)) {
            [void]$cboFont.Items.Add($fontName)
        }
        $fontCollection.Dispose()
    }
    catch { }
    $currentFont = [string]$working['uiFontName']
    if ([string]::IsNullOrWhiteSpace($currentFont)) { $cboFont.SelectedIndex = 0 }
    else {
        $fontIndex = $cboFont.Items.IndexOf($currentFont)
        if ($fontIndex -ge 0) { $cboFont.SelectedIndex = $fontIndex }
        else { [void]$cboFont.Items.Add($currentFont); $cboFont.SelectedItem = $currentFont }
    }
    $tabBasics.Controls.Add($cboFont)

    $lblFontHint = New-FormLabel 'Applies next time it starts' 278 421 206
    $lblFontHint.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
    $tabBasics.Controls.Add($lblFontHint)

    # ------------------------------------------------------ Notifications ----

    $lblNotifyIntro = New-FormWrappedLabel `
        -Text 'These are the small pop-ups near the clock. Windows Focus assist can hold them back, so treat them as a convenience rather than the record - the activity list on the dashboard always has everything.' `
        -X 16 -Y 12 -Width 468
    $tabNotify.Controls.Add($lblNotifyIntro)

    $checkBoxes['showErrorNotifications'] = New-SettingCheck -Parent $tabNotify -Y 92 `
        -Text 'Tell me when a refresh fails' `
        -Explanation 'Recommended. This is the one worth keeping on: a failed refresh means the workbook still holds the old numbers.' `
        -Checked (ConvertTo-BoolValue $working['showErrorNotifications'] $true)

    $checkBoxes['showSuccessNotifications'] = New-SettingCheck -Parent $tabNotify -Y 174 `
        -Text 'Tell me when a refresh succeeds' `
        -Explanation 'Reassuring at first, noisy once you trust it. The Last Run column says the same thing without interrupting you.' `
        -Checked (ConvertTo-BoolValue $working['showSuccessNotifications'] $true)

    $checkBoxes['showTriggerNotifications'] = New-SettingCheck -Parent $tabNotify -Y 256 `
        -Text 'Tell me when something sets a rule off' `
        -Explanation 'Off by default. Useful for a day or two while you check that a new rule is watching the right folder.' `
        -Checked (ConvertTo-BoolValue $working['showTriggerNotifications'] $false)

    $checkBoxes['showTriggeredResultPopup'] = New-SettingCheck -Parent $tabNotify -Y 338 `
        -Text 'Show a full result window after an automatic refresh' `
        -Explanation 'A window listing every workbook and query, which waits for you to close it. Off unless you need that detail each time.' `
        -Checked (ConvertTo-BoolValue $working['showTriggeredResultPopup'] $false)

    # ----------------------------------------------------------- Advanced ----

    $lblAdvIntro = New-FormWrappedLabel `
        -Text 'The defaults suit almost everyone. Change these only if something specific is going wrong.' `
        -X 16 -Y 10 -Width 468
    $tabAdv.Controls.Add($lblAdvIntro)

    $numerics['defaultDebounceSeconds'] = New-SettingNumber -Parent $tabAdv -Y 58 `
        -Text 'Wait after a file changes' -Unit 'seconds' -Min 0 -Max 3600 `
        -Explanation 'A program writing a file touches it several times. Waiting lets it finish, so the rule runs once instead of five times.' `
        -Value (ConvertTo-IntValue $working['defaultDebounceSeconds'] 5 0)

    $numerics['defaultCooldownSeconds'] = New-SettingNumber -Parent $tabAdv -Y 136 `
        -Text 'Then leave that rule alone for' -Unit 'seconds' -Min 0 -Max 86400 `
        -Explanation 'Stops a folder that keeps changing from queueing the same rule over and over again.' `
        -Value (ConvertTo-IntValue $working['defaultCooldownSeconds'] 30 0)

    $numerics['defaultRefreshWarningSeconds'] = New-SettingNumber -Parent $tabAdv -Y 214 `
        -Text 'Call a refresh slow after' -Unit 'seconds' -Min 5 -Max 7200 `
        -Explanation 'A warning, not a time limit - it keeps waiting for Excel. Use Cancel Job on the dashboard to actually stop one.' `
        -Value (ConvertTo-IntValue $working['defaultRefreshWarningSeconds'] 300 5)

    $numerics['startupPromptDelaySeconds'] = New-SettingNumber -Parent $tabAdv -Y 292 `
        -Text 'After signing in, wait before asking' -Unit 'seconds' -Min 0 -Max 900 `
        -Explanation 'Applies to sign-in rules only. Long enough that your network drives are ready before it asks.' `
        -Value (ConvertTo-IntValue $working['startupPromptDelaySeconds'] 30 0)

    $numerics['watcherHealthCheckSeconds'] = New-SettingNumber -Parent $tabAdv -Y 370 `
        -Text 'Re-check that folder watching works every' -Unit 'seconds' -Min 10 -Max 3600 `
        -Explanation 'A network share that drops out can silently stop being watched. This notices, and starts watching again.' `
        -Value (ConvertTo-IntValue $working['watcherHealthCheckSeconds'] 60 10)

    $numerics['logRetentionDays'] = New-SettingNumber -Parent $tabAdv -Y 448 `
        -Text 'Keep the log for' -Unit 'days' -Min 1 -Max 3650 `
        -Explanation 'Older entries are deleted automatically. The log is plain text, in the folder you can open below.' `
        -Value (ConvertTo-IntValue $working['logRetentionDays'] 30 1)

    $checkBoxes['allowWorkbookMacrosByDefault'] = New-SettingCheck -Parent $tabAdv -Y 526 `
        -Text 'Let workbooks run their macros' `
        -Explanation 'Off is safest, and off is the default. Turn it on only if a workbook builds its connections when it opens. A single rule can allow it without changing this.' `
        -Checked (ConvertTo-BoolValue $working['allowWorkbookMacrosByDefault'] $false)

    $checkBoxes['debugLogging'] = New-SettingCheck -Parent $tabAdv -Y 616 `
        -Text 'Record extra detail in the log' `
        -Explanation 'Turn this on while chasing a problem, and off again afterwards. It makes the log much larger.' `
        -Checked (ConvertTo-BoolValue $working['debugLogging'] $false)

    $btnOpenConfig = New-FormAutoButton -Text 'Open settings folder' -MinimumWidth 160
    $tabAdv.Controls.Add($btnOpenConfig)
    $btnOpenConfig.Location = New-Object System.Drawing.Point(18, 706)
    $btnOpenConfig.Add_Click({ Open-FolderInExplorer -Folder $Paths.ConfigDir }.GetNewClosure())

    $btnOpenLogs = New-FormAutoButton -Text 'Open log folder' -MinimumWidth 140
    $tabAdv.Controls.Add($btnOpenLogs)
    $btnOpenLogs.Location = New-Object System.Drawing.Point(($btnOpenConfig.Right + 10), 706)
    $btnOpenLogs.Add_Click({ Open-FolderInExplorer -Folder $Paths.LogDir }.GetNewClosure())

    $lblHow = New-FormWrappedLabel `
        -Text 'Folder watching does not scan anything. Windows tells the application when a file changes, so how many files sit in the folder makes no difference to your machine.' `
        -X 18 -Y 752 -Width 466 -Color ([System.Drawing.Color]::FromArgb(70, 100, 75))
    $tabAdv.Controls.Add($lblHow)

    # ------------------------------------------------------------ buttons ----

    $btnDefaults = New-FormAutoButton -Text 'Restore defaults' -MinimumWidth 140
    $form.Controls.Add($btnDefaults)
    $btnDefaults.Location = New-Object System.Drawing.Point(12, 480)

    $btnOk = New-FormAutoButton -Text 'Save' -MinimumWidth 96
    $form.Controls.Add($btnOk)
    $btnOk.Location = New-Object System.Drawing.Point(332, 480)

    $btnCancel = New-FormAutoButton -Text 'Cancel' -MinimumWidth 96
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $btnCancel.Location = New-Object System.Drawing.Point(434, 480)
    $form.CancelButton = $btnCancel

    $btnDefaults.Add_Click({
        $answer = [System.Windows.Forms.MessageBox]::Show(
            'Put every setting on every tab back to its default?' + [Environment]::NewLine + [Environment]::NewLine +
            'Your rules are not touched.',
            'Restore defaults', 'YesNo', 'Question')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        foreach ($key in @($checkBoxes.Keys)) {
            # Windows decides this one; leaving it alone keeps the checkbox
            # honest about what is actually registered.
            if ($key -eq 'startWithWindows') { continue }
            $checkBoxes[$key].Checked = (ConvertTo-BoolValue $defaults[$key] $false)
        }
        foreach ($key in @($numerics.Keys)) {
            $value = ConvertTo-IntValue $defaults[$key] 0 0
            if ($value -lt $numerics[$key].Minimum) { $value = [int]$numerics[$key].Minimum }
            if ($value -gt $numerics[$key].Maximum) { $value = [int]$numerics[$key].Maximum }
            $numerics[$key].Value = $value
        }
        $cboFont.SelectedIndex = 0
    }.GetNewClosure())

    $btnOk.Add_Click({
        $result = Merge-DefaultValues -Default (Get-DefaultAppSettings) -Value $working

        foreach ($key in @($checkBoxes.Keys)) { $result[$key] = $checkBoxes[$key].Checked }
        foreach ($key in @($numerics.Keys))   { $result[$key] = [int]$numerics[$key].Value }
        if ((ConvertTo-BoolValue $result['checkForUpdatesAutomatically'] $true) -and
            -not (ConvertTo-BoolValue $working['checkForUpdatesAutomatically'] $true)) {
            # Enabling it should cause a fresh check on the next UI heartbeat,
            # even when an old timestamp remains from before it was disabled.
            $result['lastUpdateCheckUtc'] = ''
        }
        # Preserve the existing font size; only the font family is user-selectable.
        $result['uiFontSize'] = ConvertTo-IntValue $working['uiFontSize'] 9 8
        $result['uiFontName'] = if ($cboFont.SelectedIndex -le 0) { '' } else { [string]$cboFont.SelectedItem }

        # Apply the startup entry now so a failure is visible immediately.
        $startupWanted = [bool]$result['startWithWindows']
        if ($startupWanted -ne (Test-StartupRegistration)) {
            $startupResult = Set-StartupRegistration -Paths $Paths -Enabled $startupWanted
            if (-not $startupResult.Success) {
                [System.Windows.Forms.MessageBox]::Show($startupResult.Message, 'Windows startup', 'OK', 'Warning') | Out-Null
                $result['startWithWindows'] = (Test-StartupRegistration)
            }
        }

        $form.Tag          = $result
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }.GetNewClosure())

    # Reflect reality rather than the file: the registry is the source of truth.
    $checkBoxes['startWithWindows'].Checked = (Test-StartupRegistration)

    Set-FormWithinWorkingArea -Form $form
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $form.Tag }
    return $null
}
