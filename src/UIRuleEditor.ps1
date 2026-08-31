# ==============================================================================
#  UIRuleEditor.ps1  (UI runspace only)
#  Add / Edit Rule dialog + the Excel action dialog it owns.
#  Everything a user needs to build a new automation lives here - no one should
#  ever have to open a .ps1 file to add a trigger.
# ==============================================================================

Set-StrictMode -Version 1.0

function New-FormLabel {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 120, [int]$Height = 20)
    $label = New-Object System.Windows.Forms.Label
    $label.Text     = $Text
    $label.Location = New-Object System.Drawing.Point($X, ($Y + 2))
    $label.Size     = New-Object System.Drawing.Size($Width, $Height)
    return $label
}

function New-FormWrappedLabel {
    <#
        A paragraph label whose height follows its text. Fixed-height labels
        silently lose their second/third line with some Windows font and DPI
        combinations, so explanatory copy should always use this helper.
    #>
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 450,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(105, 105, 105)
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Text
    $label.Location  = New-Object System.Drawing.Point($X, $Y)
    $label.AutoSize  = $true
    $label.MinimumSize = New-Object System.Drawing.Size($Width, 0)
    $label.MaximumSize = New-Object System.Drawing.Size($Width, 0)
    $label.ForeColor = $Color
    $label.AutoEllipsis = $false
    return $label
}

function New-FormNumeric {
    param([int]$X, [int]$Y, [int]$Width = 70, [int]$Minimum = 0, [int]$Maximum = 86400, [int]$Value = 0)
    $numeric = New-Object System.Windows.Forms.NumericUpDown
    $numeric.Location = New-Object System.Drawing.Point($X, $Y)
    $numeric.Size     = New-Object System.Drawing.Size($Width, 24)
    $numeric.Minimum  = $Minimum
    $numeric.Maximum  = $Maximum
    if ($Value -lt $Minimum) { $Value = $Minimum }
    if ($Value -gt $Maximum) { $Value = $Maximum }
    $numeric.Value = $Value
    return $numeric
}

function Show-RuleEditor {
    <#
        Returns a normalised rule hashtable, or $null when the user cancels.
        The rule passed in is never modified: edits are applied to a copy so a
        cancelled dialog leaves the live configuration untouched.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [hashtable]$AppSettings
    )

    $working = ConvertTo-NormalizedRule $Rule
    $actions = New-Object System.Collections.ArrayList
    foreach ($action in @($working.actions)) { [void]$actions.Add($action) }

    $font = Get-UiFont

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Trigger Rule'
    # Fixed-coordinate WinForms layout: Windows compatibility scaling is used
    # so fonts and controls scale together on high-DPI/Retina displays.
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize      = New-Object System.Drawing.Size(680, 660)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterParent'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = $font

    # ---- basic ------------------------------------------------------------
    $gbBasic = New-Object System.Windows.Forms.GroupBox
    $gbBasic.Text     = 'Basic'
    $gbBasic.Location = New-Object System.Drawing.Point(12, 10)
    $gbBasic.Size     = New-Object System.Drawing.Size(616, 82)
    $form.Controls.Add($gbBasic)

    $gbBasic.Controls.Add((New-FormLabel 'Rule Name:' 12 22 124))
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(140, 22)
    $txtName.Size     = New-Object System.Drawing.Size(456, 24)
    $txtName.Text     = [string]$working.name
    $gbBasic.Controls.Add($txtName)

    $chkEnabled = New-Object System.Windows.Forms.CheckBox
    $chkEnabled.Text     = 'Enabled'
    $chkEnabled.Location = New-Object System.Drawing.Point(140, 50)
    $chkEnabled.Size     = New-Object System.Drawing.Size(200, 22)
    $chkEnabled.Checked  = (ConvertTo-BoolValue $working.enabled $true)
    $gbBasic.Controls.Add($chkEnabled)

    # ---- trigger ----------------------------------------------------------
    # The three trigger kinds need completely different fields, so each lives in
    # its own panel and only one is ever visible. Radio buttons inside a panel
    # are also grouped by that panel, which is what makes the schedule and logon
    # options independent of each other.
    $gbTrigger = New-Object System.Windows.Forms.GroupBox
    $gbTrigger.Text     = 'When should this run?'
    $gbTrigger.Location = New-Object System.Drawing.Point(12, 98)
    $gbTrigger.Size     = New-Object System.Drawing.Size(616, 158)
    $form.Controls.Add($gbTrigger)

    $gbTrigger.Controls.Add((New-FormLabel 'Trigger:' 12 22 124))
    $cboType = New-Object System.Windows.Forms.ComboBox
    $cboType.Location      = New-Object System.Drawing.Point(140, 22)
    $cboType.Size          = New-Object System.Drawing.Size(280, 24)
    $cboType.DropDownStyle = 'DropDownList'
    foreach ($type in (Get-TriggerTypeList)) { [void]$cboType.Items.Add((Get-TriggerTypeLabel $type)) }
    $cboType.SelectedIndex = [Math]::Max(0, [array]::IndexOf((Get-TriggerTypeList), [string]$working.trigger.type))
    $gbTrigger.Controls.Add($cboType)

    $chkAskBefore = New-Object System.Windows.Forms.CheckBox
    $chkAskBefore.Text     = 'Always ask before refresh'
    $chkAskBefore.Location = New-Object System.Drawing.Point(422, 22)
    $chkAskBefore.Size     = New-Object System.Drawing.Size(182, 24)
    $chkAskBefore.Checked  = (ConvertTo-BoolValue $working.askBeforeRefresh $false)
    $chkAskBefore.Tag      = 'AskBeforeRefresh'
    $gbTrigger.Controls.Add($chkAskBefore)
    $askTip = New-Object System.Windows.Forms.ToolTip
    $askTip.SetToolTip($chkAskBefore, 'For file and scheduled triggers, ask for confirmation before the job is queued. Manual runs still start immediately.')

    $panels = @{}
    foreach ($key in @('File', 'Schedule', 'Logon', 'Manual')) {
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Location = New-Object System.Drawing.Point(12, 50)
        $panel.Size     = New-Object System.Drawing.Size(592, 98)
        $panel.Visible  = $false
        $gbTrigger.Controls.Add($panel)
        $panels[$key] = $panel
    }

    # ---- file / folder panel ----
    $lblPath = New-FormLabel 'Folder:' 0 4 120
    $panels['File'].Controls.Add($lblPath)
    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(128, 2)
    $txtPath.Size     = New-Object System.Drawing.Size(360, 24)
    $txtPath.Text     = [string]$working.trigger.path
    $panels['File'].Controls.Add($txtPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = 'Browse'
    $btnBrowse.Location = New-Object System.Drawing.Point(496, 0)
    $btnBrowse.Size     = New-Object System.Drawing.Size(90, 28)
    $panels['File'].Controls.Add($btnBrowse)

    $lblFilter = New-FormLabel 'File type:' 0 36 120
    $panels['File'].Controls.Add($lblFilter)

    # Common file types are offered as presets, while Custom keeps full backward
    # compatibility with existing wildcard filters. Excel files intentionally use
    # several patterns; TriggerManager turns that into a broad watcher plus a cheap
    # filename match, so the folder is still event-driven and is never polled.
    $cboFilterPreset = New-Object System.Windows.Forms.ComboBox
    $cboFilterPreset.Location      = New-Object System.Drawing.Point(128, 34)
    $cboFilterPreset.Size          = New-Object System.Drawing.Size(190, 24)
    $cboFilterPreset.DropDownStyle = 'DropDownList'
    foreach ($label in @('CSV (*.csv)', 'Text (*.txt)', 'Excel files', 'Any file (*.*)', 'Custom')) {
        [void]$cboFilterPreset.Items.Add($label)
    }
    $panels['File'].Controls.Add($cboFilterPreset)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(326, 34)
    $txtFilter.Size     = New-Object System.Drawing.Size(260, 24)
    $txtFilter.Text     = [string]$working.trigger.filter
    $panels['File'].Controls.Add($txtFilter)

    $filterPresets = @{
        'CSV (*.csv)'  = '*.csv'
        'Text (*.txt)' = '*.txt'
        'Excel files'  = '*.xlsx;*.xlsm;*.xlsb;*.xls'
        'Any file (*.*)' = '*.*'
    }

    $applyFilterPreset = {
        $selected = [string]$cboFilterPreset.SelectedItem
        if ($selected -eq 'Custom') {
            $txtFilter.Enabled = $true
            return
        }
        if ($filterPresets.ContainsKey($selected)) {
            $txtFilter.Text = [string]$filterPresets[$selected]
            $txtFilter.Enabled = $false
        }
    }.GetNewClosure()

    $currentFilter = [string]$working.trigger.filter
    $matchedPreset = 'Custom'
    foreach ($key in @($filterPresets.Keys)) {
        if ([string]::Equals([string]$filterPresets[$key], $currentFilter, [StringComparison]::OrdinalIgnoreCase)) {
            $matchedPreset = $key
            break
        }
    }
    $cboFilterPreset.SelectedItem = $matchedPreset
    & $applyFilterPreset
    $cboFilterPreset.Add_SelectedIndexChanged($applyFilterPreset)

    $lblContains = New-FormLabel 'Name contains:' 0 68 120
    $panels['File'].Controls.Add($lblContains)
    $txtContains = New-Object System.Windows.Forms.TextBox
    $txtContains.Location = New-Object System.Drawing.Point(128, 66)
    $txtContains.Size     = New-Object System.Drawing.Size(120, 24)
    $txtContains.Text     = [string]$working.trigger.contains
    $panels['File'].Controls.Add($txtContains)

    $lblExclude = New-FormLabel 'Name must not contain:' 264 68 170
    $panels['File'].Controls.Add($lblExclude)
    $txtExclude = New-Object System.Windows.Forms.TextBox
    $txtExclude.Location = New-Object System.Drawing.Point(440, 66)
    $txtExclude.Size     = New-Object System.Drawing.Size(146, 24)
    $txtExclude.Text     = [string]$working.trigger.exclude
    $panels['File'].Controls.Add($txtExclude)

    # ---- schedule panel ----
    $panels['Schedule'].Controls.Add((New-FormLabel 'Run at:' 0 4 120))
    $txtTime = New-Object System.Windows.Forms.TextBox
    $txtTime.Location = New-Object System.Drawing.Point(128, 2)
    $txtTime.Size     = New-Object System.Drawing.Size(60, 24)
    $txtTime.Text     = [string]$working.trigger.scheduleTime
    $panels['Schedule'].Controls.Add($txtTime)

    $lblTimeHint = New-FormLabel '24-hour clock, for example 07:30' 196 4 260
    $lblTimeHint.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
    $panels['Schedule'].Controls.Add($lblTimeHint)

    $panels['Schedule'].Controls.Add((New-FormLabel 'Repeat:' 0 36 120))
    $repeatButtons = @{}
    $x = 128
    foreach ($spec in @(
            @{ Key = 'Daily';    Text = 'Every day';     Width = 96 },
            @{ Key = 'Weekdays'; Text = 'Mon to Fri';    Width = 106 },
            @{ Key = 'Custom';   Text = 'Selected days'; Width = 120 })) {
        $radio = New-Object System.Windows.Forms.RadioButton
        $radio.Text     = $spec.Text
        $radio.Location = New-Object System.Drawing.Point($x, 34)
        $radio.Size     = New-Object System.Drawing.Size($spec.Width, 22)
        $panels['Schedule'].Controls.Add($radio)
        $repeatButtons[$spec.Key] = $radio
        $x += $spec.Width
    }

    $dayBoxes = @{}
    $x = 128
    foreach ($day in (Get-WeekDayNameList)) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text     = $day.Substring(0, 3)
        $checkBox.Location = New-Object System.Drawing.Point($x, 66)
        $checkBox.Size     = New-Object System.Drawing.Size(56, 22)
        $checkBox.Checked  = (@($working.trigger.scheduleDays) -contains $day)
        $panels['Schedule'].Controls.Add($checkBox)
        $dayBoxes[$day] = $checkBox
        $x += 60
    }

    # ---- logon panel ----
    $rdoLogonAsk = New-Object System.Windows.Forms.RadioButton
    $rdoLogonAsk.Text     = 'Ask me before refreshing  (recommended)'
    $rdoLogonAsk.Location = New-Object System.Drawing.Point(0, 4)
    $rdoLogonAsk.Size     = New-Object System.Drawing.Size(400, 22)
    $panels['Logon'].Controls.Add($rdoLogonAsk)

    $rdoLogonAuto = New-Object System.Windows.Forms.RadioButton
    $rdoLogonAuto.Text     = 'Refresh automatically, without asking'
    $rdoLogonAuto.Location = New-Object System.Drawing.Point(0, 30)
    $rdoLogonAuto.Size     = New-Object System.Drawing.Size(400, 22)
    $panels['Logon'].Controls.Add($rdoLogonAuto)

    if ([string]$working.trigger.logonBehavior -eq 'Automatic') { $rdoLogonAuto.Checked = $true }
    else { $rdoLogonAsk.Checked = $true }

    $lblLogonHint = New-FormWrappedLabel 'Runs shortly after you log in to Windows. The delay is set in Settings, so the refresh does not compete with everything else that starts at logon.' 0 60 580
    $panels['Logon'].Controls.Add($lblLogonHint)

    # ---- manual panel ----
    $lblManualHint = New-FormWrappedLabel 'This rule only runs when you press Run Now, or Run All Manual Rules from the tray icon.' 0 8 580
    $panels['Manual'].Controls.Add($lblManualHint)

    # ---- recent-refresh protection ----------------------------------------
    # This is a normal operating choice, not a technical tuning value, so it
    # stays visible instead of being hidden under Advanced settings.
    $gbRecentRefresh = New-Object System.Windows.Forms.GroupBox
    $gbRecentRefresh.Text     = 'Recent refresh protection'
    $gbRecentRefresh.Location = New-Object System.Drawing.Point(12, 262)
    $gbRecentRefresh.Size     = New-Object System.Drawing.Size(616, 82)
    $form.Controls.Add($gbRecentRefresh)

    $recentMinutes = ConvertTo-IntValue $working.trigger.recentRefreshPromptMinutes 0 0
    $chkRecentRefresh = New-Object System.Windows.Forms.CheckBox
    $chkRecentRefresh.Text     = 'Ask before refreshing a workbook whose queries were updated recently'
    $chkRecentRefresh.Location = New-Object System.Drawing.Point(16, 22)
    $chkRecentRefresh.Size     = New-Object System.Drawing.Size(560, 24)
    $chkRecentRefresh.Checked  = ($recentMinutes -gt 0)
    $gbRecentRefresh.Controls.Add($chkRecentRefresh)

    $lblRecentWindow = New-FormLabel 'Treat queries as recent if updated within:' 38 52 282
    $gbRecentRefresh.Controls.Add($lblRecentWindow)
    $recentUnit = $(if ($recentMinutes -ge 60 -and ($recentMinutes % 60) -eq 0) { 'hours' } else { 'minutes' })
    $recentValue = $(if ($recentUnit -eq 'hours') { [int]($recentMinutes / 60) } else { $recentMinutes })
    if ($recentValue -le 0) { $recentValue = 30 }
    $numRecentRefresh = New-FormNumeric 324 50 72 1 10080 $recentValue
    $gbRecentRefresh.Controls.Add($numRecentRefresh)

    $cboRecentUnit = New-Object System.Windows.Forms.ComboBox
    $cboRecentUnit.Location      = New-Object System.Drawing.Point(404, 50)
    $cboRecentUnit.Size          = New-Object System.Drawing.Size(96, 24)
    $cboRecentUnit.DropDownStyle = 'DropDownList'
    [void]$cboRecentUnit.Items.Add('minutes')
    [void]$cboRecentUnit.Items.Add('hours')
    $cboRecentUnit.SelectedItem = $recentUnit
    $gbRecentRefresh.Controls.Add($cboRecentUnit)

    $recentTip = New-Object System.Windows.Forms.ToolTip
    $recentTip.SetToolTip($chkRecentRefresh, 'For automatic runs, ask only when the workbook records a recent query refresh. Saving the file alone does not count. Run Now is not affected.')

    # ---- advanced (collapsed by default) ----------------------------------
    # Everything here has a sensible default, so it stays folded away until
    # someone actually needs it. The dialog shrinks to match.
    $btnAdvanced = New-Object System.Windows.Forms.Button
    $btnAdvanced.Location  = New-Object System.Drawing.Point(12, 262)
    $btnAdvanced.Size      = New-Object System.Drawing.Size(270, 28)
    $btnAdvanced.FlatStyle = 'Flat'
    $btnAdvanced.TextAlign = 'MiddleLeft'
    $form.Controls.Add($btnAdvanced)

    $gbAdvanced = New-Object System.Windows.Forms.GroupBox
    $gbAdvanced.Text     = 'Advanced  (defaults are fine for most rules)'
    $gbAdvanced.Location = New-Object System.Drawing.Point(12, 296)
    $gbAdvanced.Size     = New-Object System.Drawing.Size(616, 136)
    $gbAdvanced.Visible  = $false
    $form.Controls.Add($gbAdvanced)

    $lblDebounce = New-FormLabel 'Combine repeated changes within (sec):' 16 26 300
    $gbAdvanced.Controls.Add($lblDebounce)
    $numDebounce = New-FormNumeric 320 24 70 0 3600 (ConvertTo-IntValue $working.trigger.debounceSeconds 5 0)
    $gbAdvanced.Controls.Add($numDebounce)

    $lblCooldown = New-FormLabel 'Do not run this rule again for (sec):' 16 54 300
    $gbAdvanced.Controls.Add($lblCooldown)
    $numCooldown = New-FormNumeric 320 52 70 0 86400 (ConvertTo-IntValue $working.trigger.cooldownSeconds 30 0)
    $gbAdvanced.Controls.Add($numCooldown)

    $chkReady = New-Object System.Windows.Forms.CheckBox
    $chkReady.Text     = 'Wait until the source file is fully written'
    $chkReady.Location = New-Object System.Drawing.Point(16, 80)
    $chkReady.Size     = New-Object System.Drawing.Size(400, 22)
    $chkReady.Checked  = (ConvertTo-BoolValue $working.trigger.waitForReady $true)
    $gbAdvanced.Controls.Add($chkReady)

    $lblReadyInterval = New-FormLabel 'Check every (sec):' 36 108 130
    $gbAdvanced.Controls.Add($lblReadyInterval)
    $numReadyInterval = New-FormNumeric 170 106 60 1 300 (ConvertTo-IntValue $working.trigger.readyCheckIntervalSeconds 2 1)
    $gbAdvanced.Controls.Add($numReadyInterval)

    $lblReadyTimeout = New-FormLabel 'Give up after (sec):' 256 108 140
    $gbAdvanced.Controls.Add($lblReadyTimeout)
    $numReadyTimeout = New-FormNumeric 400 106 70 1 3600 (ConvertTo-IntValue $working.trigger.readyTimeoutSeconds 60 1)
    $gbAdvanced.Controls.Add($numReadyTimeout)

    # ---- actions ----------------------------------------------------------
    $gbActions = New-Object System.Windows.Forms.GroupBox
    $gbActions.Text     = 'Workbooks to refresh  (processed from top to bottom)'
    $gbActions.Location = New-Object System.Drawing.Point(12, 402)
    $gbActions.Size     = New-Object System.Drawing.Size(656, 200)
    $form.Controls.Add($gbActions)

    $lvActions = New-Object System.Windows.Forms.ListView
    $lvActions.Location      = New-Object System.Drawing.Point(10, 20)
    $lvActions.Size          = New-Object System.Drawing.Size(478, 150)
    $lvActions.View          = 'Details'
    $lvActions.FullRowSelect = $true
    $lvActions.MultiSelect   = $false
    $lvActions.HideSelection = $false
    $lvActions.GridLines     = $true
    $lvActions.ShowItemToolTips = $true
    [void]$lvActions.Columns.Add('Workbook', 190)
    [void]$lvActions.Columns.Add('Queries', 82)
    [void]$lvActions.Columns.Add('Warn after', 86)
    [void]$lvActions.Columns.Add('If it fails', 106)
    $gbActions.Controls.Add($lvActions)

    # A new rule has an empty list, so "Add a workbook..." is the only thing
    # that can be done next. It says so, it is taller than the rest, it is the
    # dialog's accept button until there is something in the list, and the
    # others are simply switched off until they have a row to act on.
    $actionButtons = @{}

    $btnActionAdd = New-Object System.Windows.Forms.Button
    $btnActionAdd.Text     = 'Add a workbook...'
    $btnActionAdd.Location = New-Object System.Drawing.Point(498, 20)
    $btnActionAdd.Size     = New-Object System.Drawing.Size(140, 34)
    $btnActionAdd.Font     = (Get-UiFont 9 'Bold')
    $gbActions.Controls.Add($btnActionAdd)
    $actionButtons['Add'] = $btnActionAdd

    $y = 62
    foreach ($caption in @('Edit', 'Remove', 'Up', 'Down')) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text     = $caption
        $button.Location = New-Object System.Drawing.Point(498, $y)
        $button.Size     = New-Object System.Drawing.Size(140, 28)
        $button.Enabled  = $false
        $gbActions.Controls.Add($button)
        $actionButtons[$caption] = $button
        $y += 30
    }

    $lblActionsHint = New-Object System.Windows.Forms.Label
    $lblActionsHint.Location  = New-Object System.Drawing.Point(10, 174)
    $lblActionsHint.Size      = New-Object System.Drawing.Size(480, 18)
    $lblActionsHint.ForeColor = [System.Drawing.Color]::FromArgb(150, 60, 45)
    $lblActionsHint.Text      = 'Add at least one workbook before saving.'
    $gbActions.Controls.Add($lblActionsHint)

    # ---- dialog buttons ---------------------------------------------------
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text     = 'Save'
    $btnOk.Location = New-Object System.Drawing.Point(472, 602)
    $btnOk.Size     = New-Object System.Drawing.Size(92, 30)
    $form.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Cancel'
    $btnCancel.Location     = New-Object System.Drawing.Point(574, 602)
    $btnCancel.Size         = New-Object System.Drawing.Size(92, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    # ---- local behaviour --------------------------------------------------
    $syncActionButtons = {
        $count    = $lvActions.Items.Count
        $index    = $(if ($lvActions.SelectedIndices.Count -gt 0) { $lvActions.SelectedIndices[0] } else { -1 })
        $selected = ($index -ge 0)

        $actionButtons['Edit'].Enabled   = $selected
        $actionButtons['Remove'].Enabled = $selected
        $actionButtons['Up'].Enabled     = ($selected -and $index -gt 0)
        $actionButtons['Down'].Enabled   = ($selected -and $index -lt ($count - 1))

        $lblActionsHint.Visible = ($count -eq 0)
        $btnOk.Enabled          = ($count -gt 0)
        # Until there is a workbook, Enter should add one rather than try to
        # save a rule that cannot be saved.
        $form.AcceptButton = $(if ($count -eq 0) { $actionButtons['Add'] } else { $null })
    }.GetNewClosure()

    $refreshActionList = {
        $lvActions.BeginUpdate()
        try {
            $lvActions.Items.Clear()
            foreach ($action in $actions) {
                $item = New-Object System.Windows.Forms.ListViewItem((Split-Path -Leaf ([string]$action.path)))
                $refreshText = $(if ([string]$action.refreshMethod -eq 'SelectedQueries') { '{0} selected' -f @($action.selectedQueries).Count } else { 'All' })
                [void]$item.SubItems.Add($refreshText)
                [void]$item.SubItems.Add(('{0}s' -f (ConvertTo-IntValue $action.timeoutSeconds 300 5)))
                [void]$item.SubItems.Add($(if (ConvertTo-BoolValue $action.continueOnError $true) { 'Continue' } else { 'Stop' }))
                $item.ToolTipText = [string]$action.path
                [void]$lvActions.Items.Add($item)
            }
        }
        finally {
            $lvActions.EndUpdate()
        }
        & $syncActionButtons
    }.GetNewClosure()

    $layout = @{ Expanded = $false; Available = $true; RecentVisible = $true }

    $applyLayout = {
        # A hashtable, not a plain variable: closures copy variables, so the
        # click handler and the trigger-type handler have to share one object
        # to see each other's changes.
        $expanded = ($layout.Expanded -and $layout.Available)

        $recentVisible = [bool]$layout.RecentVisible
        $gbRecentRefresh.Visible = $recentVisible
        $advancedButtonTop = $(if ($recentVisible) { 350 } else { 262 })
        $advancedGroupTop  = $advancedButtonTop + 34
        $btnAdvanced.Location = New-Object System.Drawing.Point(12, $advancedButtonTop)
        $gbAdvanced.Location  = New-Object System.Drawing.Point(12, $advancedGroupTop)

        $btnAdvanced.Visible = $layout.Available
        $btnAdvanced.Text    = $(if ($expanded) { '  Advanced settings   ^' } else { '  Advanced settings   v' })
        $gbAdvanced.Visible  = $expanded

        $actionsTop = $advancedGroupTop
        if ($expanded) { $actionsTop = $advancedGroupTop + 146 }
        elseif (-not $layout.Available) { $actionsTop = $(if ($recentVisible) { 352 } else { 266 }) }

        $gbActions.Location = New-Object System.Drawing.Point(12, $actionsTop)
        $buttonsTop = $actionsTop + 210
        $btnOk.Location     = New-Object System.Drawing.Point(472, $buttonsTop)
        $btnCancel.Location = New-Object System.Drawing.Point(574, $buttonsTop)
        $form.ClientSize    = New-Object System.Drawing.Size(680, ($buttonsTop + 44))
    }.GetNewClosure()

    $btnAdvanced.Add_Click({
        $layout.Expanded = -not $layout.Expanded
        & $applyLayout
    }.GetNewClosure())

    $applyRepeatMode = {
        # The radio buttons are a shortcut for the seven checkboxes, which stay
        # the single source of truth for what actually gets saved.
        $weekdays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
        if ($repeatButtons['Daily'].Checked) {
            foreach ($day in (Get-WeekDayNameList)) { $dayBoxes[$day].Checked = $true }
        }
        elseif ($repeatButtons['Weekdays'].Checked) {
            foreach ($day in (Get-WeekDayNameList)) { $dayBoxes[$day].Checked = ($weekdays -contains $day) }
        }
        $custom = $repeatButtons['Custom'].Checked
        foreach ($day in (Get-WeekDayNameList)) { $dayBoxes[$day].Enabled = $custom }
    }.GetNewClosure()

    foreach ($key in @($repeatButtons.Keys)) { $repeatButtons[$key].Add_CheckedChanged($applyRepeatMode) }

    $applyTriggerType = {
        $selectedType = (Get-TriggerTypeList)[$cboType.SelectedIndex]
        $usesWatcher  = Test-TriggerUsesWatcher $selectedType
        $usesFolder   = Test-TriggerUsesFolder $selectedType

        $visiblePanel = 'Manual'
        if ($usesWatcher)                             { $visiblePanel = 'File' }
        elseif (Test-TriggerIsScheduled $selectedType) { $visiblePanel = 'Schedule' }
        elseif (Test-TriggerIsLogon $selectedType)     { $visiblePanel = 'Logon' }
        foreach ($key in @($panels.Keys)) { $panels[$key].Visible = ($key -eq $visiblePanel) }

        $lblPath.Text = $(if ($usesFolder) { 'Folder:' } else { 'File:' })
        foreach ($control in @($lblFilter, $cboFilterPreset, $txtFilter, $lblContains, $txtContains, $lblExclude, $txtExclude)) {
            $control.Enabled = $usesFolder
        }
        if ($usesFolder) { & $applyFilterPreset }

        # Advanced contains file-watcher timing only. Do not show meaningless
        # disabled controls for schedule/logon rules; on file triggers every
        # applicable field remains editable.
        $recentApplicable = ($usesWatcher -or (Test-TriggerIsScheduled $selectedType) -or
            ((Test-TriggerIsLogon $selectedType) -and $rdoLogonAuto.Checked))
        $layout.RecentVisible = $recentApplicable
        $layout.Available = $usesWatcher
        $chkRecentRefresh.Enabled = $recentApplicable
        $lblRecentWindow.Enabled = ($recentApplicable -and $chkRecentRefresh.Checked)
        $numRecentRefresh.Enabled = ($recentApplicable -and $chkRecentRefresh.Checked)
        $cboRecentUnit.Enabled = ($recentApplicable -and $chkRecentRefresh.Checked)
        # Manual runs are already an explicit user action. Logon rules keep their
        # dedicated Ask/Automatic choice below, so this generic switch applies to
        # file and scheduled triggers only.
        $chkAskBefore.Enabled = ($usesWatcher -or (Test-TriggerIsScheduled $selectedType))
        if (Test-TriggerIsLogon $selectedType) {
            $chkAskBefore.Checked = $false
        }
        & $applyLayout
        $numReadyInterval.Enabled = ($usesWatcher -and $chkReady.Checked)
        $numReadyTimeout.Enabled  = ($usesWatcher -and $chkReady.Checked)
        $lblReadyInterval.Enabled = ($usesWatcher -and $chkReady.Checked)
        $lblReadyTimeout.Enabled  = ($usesWatcher -and $chkReady.Checked)
    }.GetNewClosure()

    $cboType.Add_SelectedIndexChanged($applyTriggerType)
    $chkReady.Add_CheckedChanged($applyTriggerType)
    $chkRecentRefresh.Add_CheckedChanged($applyTriggerType)
    $rdoLogonAsk.Add_CheckedChanged($applyTriggerType)
    $rdoLogonAuto.Add_CheckedChanged($applyTriggerType)

    $btnBrowse.Add_Click({
        $selectedType = (Get-TriggerTypeList)[$cboType.SelectedIndex]
        if (Test-TriggerUsesFolder $selectedType) {
            $picked = Show-FolderPicker -InitialPath $txtPath.Text -Owner $form `
                -Title 'Select the folder to monitor' -OkLabel 'Monitor This Folder'
        }
        else {
            $picked = Show-FilePicker -InitialPath $txtPath.Text -Owner $form `
                -Title 'Select the file to monitor'
        }
        if (-not [string]::IsNullOrWhiteSpace($picked)) { $txtPath.Text = $picked }
    }.GetNewClosure())

    $actionButtons['Add'].Add_Click({
        $action = Show-ActionEditor -Action (Get-DefaultAction) -AppSettings $AppSettings
        if ($null -ne $action) {
            if ((Find-DuplicateWorkbookAction -Actions @($actions.ToArray()) -CandidatePath ([string]$action.path)) -ge 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'That workbook is already in this rule. Select it and use Edit instead.',
                    'Duplicate workbook', 'OK', 'Information') | Out-Null
                return
            }
            [void]$actions.Add($action)
            & $refreshActionList
        }
    }.GetNewClosure())

    $editAction = {
        if ($lvActions.SelectedIndices.Count -eq 0) { return }
        $index = $lvActions.SelectedIndices[0]
        $action = Show-ActionEditor -Action $actions[$index] -AppSettings $AppSettings
        if ($null -ne $action) {
            if ((Find-DuplicateWorkbookAction -Actions @($actions.ToArray()) -CandidatePath ([string]$action.path) -ExcludeIndex $index) -ge 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'That workbook is already in this rule.',
                    'Duplicate workbook', 'OK', 'Information') | Out-Null
                return
            }
            $actions[$index] = $action
            & $refreshActionList
        }
    }.GetNewClosure()

    $actionButtons['Edit'].Add_Click($editAction)
    $lvActions.Add_DoubleClick($editAction)
    $lvActions.Add_SelectedIndexChanged($syncActionButtons)

    $actionButtons['Remove'].Add_Click({
        if ($lvActions.SelectedIndices.Count -eq 0) { return }
        $index = $lvActions.SelectedIndices[0]
        $actions.RemoveAt($index)
        & $refreshActionList
    }.GetNewClosure())

    $actionButtons['Up'].Add_Click({
        if ($lvActions.SelectedIndices.Count -eq 0) { return }
        $index = $lvActions.SelectedIndices[0]
        if ($index -le 0) { return }
        $item = $actions[$index]
        $actions.RemoveAt($index)
        $actions.Insert($index - 1, $item)
        & $refreshActionList
        $lvActions.Items[$index - 1].Selected = $true
    }.GetNewClosure())

    $actionButtons['Down'].Add_Click({
        if ($lvActions.SelectedIndices.Count -eq 0) { return }
        $index = $lvActions.SelectedIndices[0]
        if ($index -ge ($actions.Count - 1)) { return }
        $item = $actions[$index]
        $actions.RemoveAt($index)
        $actions.Insert($index + 1, $item)
        & $refreshActionList
        $lvActions.Items[$index + 1].Selected = $true
    }.GetNewClosure())

    $btnOk.Add_Click({
        $selectedType = (Get-TriggerTypeList)[$cboType.SelectedIndex]

        $recentRefreshPromptMinutes = 0
        $recentApplicable = ((Test-TriggerUsesWatcher $selectedType) -or
            (Test-TriggerIsScheduled $selectedType) -or
            ((Test-TriggerIsLogon $selectedType) -and $rdoLogonAuto.Checked))
        if ($chkRecentRefresh.Checked -and $recentApplicable) {
            $recentRefreshPromptMinutes = [int]$numRecentRefresh.Value
            if ([string]$cboRecentUnit.SelectedItem -eq 'hours') { $recentRefreshPromptMinutes *= 60 }
        }

        $candidate = ConvertTo-NormalizedRule @{
            id      = [string]$working.id
            name    = $txtName.Text.Trim()
            enabled = $chkEnabled.Checked
            askBeforeRefresh = $(if ((Test-TriggerUsesWatcher $selectedType) -or (Test-TriggerIsScheduled $selectedType)) { $chkAskBefore.Checked } else { $false })
            trigger = @{
                type                      = $selectedType
                path                      = $txtPath.Text.Trim()
                filter                    = $txtFilter.Text.Trim()
                contains                  = $txtContains.Text.Trim()
                exclude                   = $txtExclude.Text.Trim()
                debounceSeconds           = [int]$numDebounce.Value
                cooldownSeconds           = [int]$numCooldown.Value
                waitForReady              = $chkReady.Checked
                readyCheckIntervalSeconds = [int]$numReadyInterval.Value
                readyTimeoutSeconds       = [int]$numReadyTimeout.Value
                recentRefreshPromptMinutes = $recentRefreshPromptMinutes
                scheduleTime              = $txtTime.Text.Trim()
                scheduleDays              = @(@(Get-WeekDayNameList) | Where-Object { $dayBoxes[$_].Checked })
                logonBehavior             = $(if ($rdoLogonAuto.Checked) { 'Automatic' } else { 'Ask' })
            }
            actions = @($actions.ToArray())
        }

        $validation = Test-RuleConfiguration -Rule $candidate
        if ($validation.Errors.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                ('This rule cannot be saved yet:' + [Environment]::NewLine + [Environment]::NewLine + ($validation.Errors -join [Environment]::NewLine)),
                'Check the rule', 'OK', 'Warning') | Out-Null
            return
        }

        if ($validation.Warnings.Count -gt 0) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                (($validation.Warnings -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine +
                 'The rule can still be saved and will start working once the path is reachable. Save it?'),
                'Warning', 'YesNo', 'Warning')
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $form.Tag          = $candidate
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }.GetNewClosure())

    $defaultTrigger = Get-DefaultTrigger
    $layout.Expanded = (
        ((ConvertTo-IntValue $working.trigger.debounceSeconds 5 0) -ne $defaultTrigger.debounceSeconds) -or
        ((ConvertTo-IntValue $working.trigger.cooldownSeconds 30 0) -ne $defaultTrigger.cooldownSeconds) -or
        ((ConvertTo-BoolValue $working.trigger.waitForReady $true) -ne $defaultTrigger.waitForReady) -or
        ((ConvertTo-IntValue $working.trigger.readyCheckIntervalSeconds 2 1) -ne $defaultTrigger.readyCheckIntervalSeconds) -or
        ((ConvertTo-IntValue $working.trigger.readyTimeoutSeconds 60 1) -ne $defaultTrigger.readyTimeoutSeconds)
    )

    $selectedDays = @($working.trigger.scheduleDays)
    $weekdayNames = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
    $isWeekdaysOnly = ($selectedDays.Count -eq 5)
    if ($isWeekdaysOnly) {
        foreach ($day in $weekdayNames) { if ($selectedDays -notcontains $day) { $isWeekdaysOnly = $false; break } }
    }
    if ($selectedDays.Count -eq 7)   { $repeatButtons['Daily'].Checked = $true }
    elseif ($isWeekdaysOnly)         { $repeatButtons['Weekdays'].Checked = $true }
    else                             { $repeatButtons['Custom'].Checked = $true }

    & $refreshActionList
    & $applyRepeatMode
    & $applyTriggerType

    Set-FormWithinWorkingArea -Form $form
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $form.Tag }
    return $null
}

function Show-ActionEditor {
    <#  One Excel workbook refresh action. Returns a hashtable or $null.  #>
    param(
        [Parameter(Mandatory = $true)]$Action,
        [hashtable]$AppSettings
    )

    $working = Merge-DefaultValues -Default (Get-DefaultAction) -Value (ConvertTo-HashtableDeep $Action)
    if ($null -ne $AppSettings -and [string]::IsNullOrWhiteSpace([string]$working.path)) {
        $working['timeoutSeconds'] = ConvertTo-IntValue $AppSettings['defaultRefreshWarningSeconds'] 300 5
        $working['allowWorkbookMacros'] = ConvertTo-BoolValue $AppSettings['allowWorkbookMacrosByDefault'] $false
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Excel Refresh Action'
    # Every coordinate in this dialog was written for a 96 DPI screen. Font
    # scaling (the WinForms default) would re-scale them by whatever runtime UI
    # font happens to be available and make labels overlap the fields next to
    # them; DPI scaling multiplies them by the real display scale instead, which
    # is exactly what is wanted. On a 100% display the factor is 1 and nothing
    # moves, so this is safe either way.
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize      = New-Object System.Drawing.Size(560, 506)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterParent'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = (Get-UiFont)

    $form.Controls.Add((New-FormLabel 'Workbook:' 14 18 100))
    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(120, 18)
    $txtPath.Size     = New-Object System.Drawing.Size(330, 24)
    $txtPath.Text     = [string]$working.path
    $form.Controls.Add($txtPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = 'Browse'
    $btnBrowse.Location = New-Object System.Drawing.Point(456, 15)
    $btnBrowse.Size     = New-Object System.Drawing.Size(90, 28)
    $form.Controls.Add($btnBrowse)

    $form.Controls.Add((New-FormLabel 'Refresh:' 14 50 100))
    $cboMethod = New-Object System.Windows.Forms.ComboBox
    $cboMethod.Location      = New-Object System.Drawing.Point(120, 50)
    $cboMethod.Size          = New-Object System.Drawing.Size(200, 22)
    $cboMethod.DropDownStyle = 'DropDownList'
    [void]$cboMethod.Items.Add('All queries')
    [void]$cboMethod.Items.Add('Selected queries')
    $cboMethod.SelectedIndex = $(if ([string]$working.refreshMethod -eq 'SelectedQueries') { 1 } else { 0 })
    $form.Controls.Add($cboMethod)

    $btnQueries = New-Object System.Windows.Forms.Button
    $btnQueries.Text = 'Choose queries...'
    $btnQueries.Location = New-Object System.Drawing.Point(330, 47)
    $btnQueries.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnQueries)

    $lblQuerySummary = New-Object System.Windows.Forms.Label
    $lblQuerySummary.Location = New-Object System.Drawing.Point(120, 78)
    $lblQuerySummary.Size = New-Object System.Drawing.Size(426, 20)
    $lblQuerySummary.ForeColor = [System.Drawing.Color]::FromArgb(95,95,95)
    $form.Controls.Add($lblQuerySummary)
    # Event handlers run in their own PowerShell scopes. Keep query selection in
    # one shared mutable object so a selection made in Choose queries survives
    # after that click handler returns.
    $queryState = @{
        Selected      = @($working.selectedQueries | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        SelectionPath = ''
    }
    if (@($queryState.Selected).Count -gt 0) { $queryState.SelectionPath = [string]$working.path }
    $workbookCheckState = @{ Path = ''; Decided = $false; HasQueries = $false; Count = 0; Reason = '' }

    $form.Controls.Add((New-FormLabel 'Warn if still running after:' 14 106 180))
    $numTimeout = New-FormNumeric 200 106 80 5 7200 (ConvertTo-IntValue $working.timeoutSeconds 300 5)
    $form.Controls.Add($numTimeout)
    $form.Controls.Add((New-FormLabel 'seconds' 286 106 60))

    $warnNote = New-Object System.Windows.Forms.Label
    $warnNote.Location  = New-Object System.Drawing.Point(120, 134)
    $warnNote.Size      = New-Object System.Drawing.Size(426, 42)
    $warnNote.Text      = 'Shows a warning in the Dashboard and log.' + [Environment]::NewLine +
        'It does not stop the refresh; the application keeps waiting.'
    $warnNote.ForeColor = [System.Drawing.Color]::FromArgb(95, 95, 95)
    $form.Controls.Add($warnNote)

    $fixedBehavior = New-Object System.Windows.Forms.Label
    $fixedBehavior.Location  = New-Object System.Drawing.Point(120, 184)
    $fixedBehavior.Size      = New-Object System.Drawing.Size(426, 42)
    $fixedBehavior.Text      = 'Excel stays hidden during refresh.' + [Environment]::NewLine +
        'After success, the workbook is saved and closed.'
    $fixedBehavior.ForeColor = [System.Drawing.Color]::FromArgb(55, 100, 145)
    $form.Controls.Add($fixedBehavior)

    # A long UNC path would otherwise open scrolled to its end.
    $form.Add_Shown({
        $txtPath.SelectionStart  = 0
        $txtPath.SelectionLength = 0
    }.GetNewClosure())

    $checkSpecs = @(
        @{ Key = 'continueOnError';        Text = 'If this workbook fails, continue to the next workbook';   Y = 236 },
        @{ Key = 'disableBackgroundQuery'; Text = 'Wait for supported connections to finish before saving'; Y = 264 },
        @{ Key = 'allowWorkbookMacros';     Text = 'Allow workbook macros (advanced)';                        Y = 360 }
    )
    $checkBoxes = @{}
    foreach ($spec in $checkSpecs) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text     = $spec.Text
        $checkBox.Location = New-Object System.Drawing.Point(120, $spec.Y)
        $checkBox.Size     = New-Object System.Drawing.Size(420, 22)
        $defaultChecked = $true
        if ($spec.Key -eq 'allowWorkbookMacros') { $defaultChecked = $false }
        $checkBox.Checked  = (ConvertTo-BoolValue $working[$spec.Key] $defaultChecked)
        $form.Controls.Add($checkBox)
        $checkBoxes[$spec.Key] = $checkBox
    }

    $syncNote = New-Object System.Windows.Forms.Label
    $syncNote.Location  = New-Object System.Drawing.Point(142, 290)
    $syncNote.Size      = New-Object System.Drawing.Size(404, 60)
    $syncNote.Text      = 'Recommended.' + [Environment]::NewLine +
        'Waits for supported connections before saving.' + [Environment]::NewLine +
        'Prevents incomplete results from being saved.'
    $syncNote.ForeColor = [System.Drawing.Color]::FromArgb(95, 95, 95)
    $form.Controls.Add($syncNote)

    $macroNote = New-Object System.Windows.Forms.Label
    $macroNote.Location  = New-Object System.Drawing.Point(142, 386)
    $macroNote.Size      = New-Object System.Drawing.Size(404, 48)
    $macroNote.Text      = 'Leave off unless trusted VBA is required.' + [Environment]::NewLine +
        'Enabling it can run code stored in the workbook.'
    $macroNote.ForeColor = [System.Drawing.Color]::FromArgb(165, 90, 25)
    $form.Controls.Add($macroNote)

    $syncTip = New-Object System.Windows.Forms.ToolTip
    $syncTip.SetToolTip($checkBoxes['disableBackgroundQuery'], 'Recommended. The application temporarily waits for supported OLE DB/ODBC connections, then restores their original settings before save.')
    $syncTip.SetToolTip($checkBoxes['allowWorkbookMacros'], 'Advanced. Keep this off unless you trust the workbook and it requires VBA to prepare the refresh.')

    $applyMacroAvailability = {
        $extension = ''
        try { $extension = [System.IO.Path]::GetExtension($txtPath.Text).ToLowerInvariant() } catch { }
        $macroCapable = ($extension -eq '.xlsm' -or $extension -eq '.xlsb')
        $checkBoxes['allowWorkbookMacros'].Enabled = $macroCapable
        if (-not $macroCapable) { $checkBoxes['allowWorkbookMacros'].Checked = $false }
        $macroNote.ForeColor = $(if ($macroCapable) {
            [System.Drawing.Color]::FromArgb(165, 90, 25)
        } else {
            [System.Drawing.Color]::FromArgb(105, 105, 105)
        })
        $macroNote.Text = $(if ($macroCapable) {
            'Leave off unless trusted VBA is required.' + [Environment]::NewLine +
                'Enabling it can run code stored in the workbook.'
        } else {
            'Available only for macro-capable .xlsm or .xlsb workbooks.'
        })
    }.GetNewClosure()
    $txtPath.Add_TextChanged($applyMacroAvailability)
    & $applyMacroAvailability

    $validateWorkbookContent = {
        param([string]$Path, [bool]$ShowFailure = $true)

        $identity = ConvertTo-RuleComparisonPath $Path
        if ([string]::IsNullOrWhiteSpace($identity)) { return $false }
        if ((ConvertTo-RuleComparisonPath ([string]$workbookCheckState.Path)) -eq $identity) {
            if ($workbookCheckState.Decided -and -not $workbookCheckState.HasQueries) { return $false }
            return $true
        }

        $previousCursor = [System.Windows.Forms.Cursor]::Current
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        try { $check = Test-WorkbookHasQueries -Path $Path }
        finally { [System.Windows.Forms.Cursor]::Current = $previousCursor }

        $workbookCheckState.Path       = $Path
        $workbookCheckState.Decided    = [bool]$check.Decided
        $workbookCheckState.HasQueries = [bool]$check.HasQueries
        $workbookCheckState.Count      = [int]$check.Count
        $workbookCheckState.Reason     = [string]$check.Reason
        if ($workbookCheckState.Decided -and -not $workbookCheckState.HasQueries) {
            if ($ShowFailure) {
                [System.Windows.Forms.MessageBox]::Show(
                    ([string]$workbookCheckState.Reason + [Environment]::NewLine + [Environment]::NewLine +
                        'Choose a workbook that contains a query, data connection, QueryTable, or external PivotTable.'),
                    'Nothing to refresh', 'OK', 'Warning') | Out-Null
            }
            return $false
        }
        return $true
    }.GetNewClosure()

    $updateQueryControls = {
        $selectedMode = ($cboMethod.SelectedIndex -eq 1)
        $btnQueries.Enabled = $selectedMode
        $lblQuerySummary.ForeColor = [System.Drawing.Color]::FromArgb(95,95,95)
        $sameCheckedPath = ((ConvertTo-RuleComparisonPath ([string]$workbookCheckState.Path)) -eq (ConvertTo-RuleComparisonPath $txtPath.Text))
        if (-not $selectedMode -and $sameCheckedPath -and $workbookCheckState.Decided -and $workbookCheckState.HasQueries) {
            $lblQuerySummary.Text = ('All queries/connections will refresh ({0} found).' -f [int]$workbookCheckState.Count)
            $lblQuerySummary.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 55)
        }
        elseif (-not $selectedMode -and $sameCheckedPath -and -not $workbookCheckState.Decided) {
            $lblQuerySummary.Text = 'Selected. This file type will be checked when Excel opens it.'
            $lblQuerySummary.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
        }
        elseif (-not $selectedMode) {
            $lblQuerySummary.Text = 'Choose a workbook. Its refreshable content will be checked.'
            $lblQuerySummary.ForeColor = [System.Drawing.Color]::FromArgb(95,95,95)
        }
        elseif (@($queryState.Selected).Count -eq 0) { $lblQuerySummary.Text = 'No queries selected yet.' }
        elseif (@($queryState.Selected).Count -le 3) { $lblQuerySummary.Text = (@($queryState.Selected) -join ', ') }
        else { $lblQuerySummary.Text = ('{0} queries selected: {1}...' -f @($queryState.Selected).Count, ((@($queryState.Selected)[0..2]) -join ', ')) }
    }.GetNewClosure()
    $cboMethod.Add_SelectedIndexChanged($updateQueryControls)
    $btnQueries.Add_Click({
        $currentPath = $txtPath.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($currentPath)) {
            [System.Windows.Forms.MessageBox]::Show('Choose a workbook first.','Power Query list','OK','Information') | Out-Null
            return
        }
        $chosen = Show-QuerySelectionDialog -WorkbookPath $currentPath -SelectedQueries @($queryState.Selected) `
            -AllowWorkbookMacros:$checkBoxes['allowWorkbookMacros'].Checked -Owner $form
        if ($null -ne $chosen) { $queryState.Selected = @($chosen); $queryState.SelectionPath = $currentPath; & $updateQueryControls }
    }.GetNewClosure())
    & $updateQueryControls

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text     = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(352, 464)
    $btnOk.Size     = New-Object System.Drawing.Size(92, 30)
    $form.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Cancel'
    $btnCancel.Location     = New-Object System.Drawing.Point(454, 464)
    $btnCancel.Size         = New-Object System.Drawing.Size(92, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $macroDecisionPath = ''
    $macroDecisionMade = $false

    $showMacroChoice = {
        param([string]$Path)
        $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
        if ($extension -ne '.xlsm' -and $extension -ne '.xlsb') { return 'NotApplicable' }

        $message = 'This workbook can contain macros.' + [Environment]::NewLine + [Environment]::NewLine +
            'Automated refresh blocks workbook macros by default for safety.' + [Environment]::NewLine +
            'Some workbooks use Workbook_Open or other VBA to prepare Power Query connections, parameters, or source paths. Those workbooks may not refresh correctly when macros are blocked.' + [Environment]::NewLine + [Environment]::NewLine +
            'Yes  - allow macros for this action (Excel Trust Center policy still applies)' + [Environment]::NewLine +
            'No   - keep macros blocked (recommended)' + [Environment]::NewLine +
            'Cancel - do not change the selected workbook'
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Macro-enabled workbook', 'YesNoCancel', 'Warning')
        if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) { return 'Allow' }
        if ($choice -eq [System.Windows.Forms.DialogResult]::No) { return 'Block' }
        return 'Cancel'
    }

    $btnBrowse.Add_Click({
        $picked = Show-FilePicker -InitialPath $txtPath.Text -Owner $form `
            -Title 'Select the workbook to refresh' `
            -Filter 'Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*'
        if ([string]::IsNullOrWhiteSpace($picked)) { return }
        if (-not (& $validateWorkbookContent $picked $true)) { return }

        $macroChoice = & $showMacroChoice $picked
        if ($macroChoice -eq 'Cancel') { return }
        if ($macroChoice -eq 'Allow') { $checkBoxes['allowWorkbookMacros'].Checked = $true }
        elseif ($macroChoice -eq 'Block') { $checkBoxes['allowWorkbookMacros'].Checked = $false }
        if (-not [string]::Equals($txtPath.Text, $picked, [System.StringComparison]::OrdinalIgnoreCase)) {
            $queryState.Selected = @()
            $queryState.SelectionPath = ''
        }
        $txtPath.Text = $picked
        & $updateQueryControls
        if ($macroChoice -ne 'NotApplicable') {
            $macroDecisionPath = $picked
            $macroDecisionMade = $true
        }
        else {
            $macroDecisionPath = ''
            $macroDecisionMade = $false
        }
    }.GetNewClosure())

    $btnOk.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Please choose a workbook.', 'Excel action', 'OK', 'Warning') | Out-Null
            return
        }

        # A path can be typed/pasted instead of selected with Browse. Give the
        # same safety choice before saving a macro-capable action.
        $currentPath = $txtPath.Text.Trim()
        if (-not (& $validateWorkbookContent $currentPath $true)) { return }
        & $updateQueryControls
        $extension = [System.IO.Path]::GetExtension($currentPath).ToLowerInvariant()
        $alreadyDecided = ($macroDecisionMade -and [string]::Equals($macroDecisionPath, $currentPath, [System.StringComparison]::OrdinalIgnoreCase))
        if (($extension -eq '.xlsm' -or $extension -eq '.xlsb') -and -not $checkBoxes['allowWorkbookMacros'].Checked -and -not $alreadyDecided) {
            $macroChoice = & $showMacroChoice $currentPath
            if ($macroChoice -eq 'Cancel') { return }
            if ($macroChoice -eq 'Allow') { $checkBoxes['allowWorkbookMacros'].Checked = $true }
        }

        $result = Get-DefaultAction
        if ($cboMethod.SelectedIndex -eq 1 -and @($queryState.Selected).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Selected queries is enabled, but no queries are selected.','Excel action','OK','Warning') | Out-Null
            return
        }
        if ($cboMethod.SelectedIndex -eq 1 -and -not [string]::Equals([string]$queryState.SelectionPath, $currentPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            [System.Windows.Forms.MessageBox]::Show('The workbook path changed after the query list was selected. Please choose the queries again for this workbook.','Excel action','OK','Warning') | Out-Null
            return
        }
        $result['path']                   = $txtPath.Text.Trim()
        $result['refreshMethod']          = $(if ($cboMethod.SelectedIndex -eq 1) { 'SelectedQueries' } else { 'RefreshAll' })
        $result['selectedQueries']        = @($queryState.Selected)
        $result['timeoutSeconds']         = [int]$numTimeout.Value
        # These are the application's normal, predictable refresh semantics;
        # they are intentionally not exposed as per-workbook switches.
        $result['save']                   = $true
        $result['close']                  = $true
        $result['visible']                = $false
        $result['continueOnError']        = $checkBoxes['continueOnError'].Checked
        $result['disableBackgroundQuery'] = $checkBoxes['disableBackgroundQuery'].Checked
        $result['allowWorkbookMacros']     = $checkBoxes['allowWorkbookMacros'].Checked

        $form.Tag          = $result
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }.GetNewClosure())

    Set-FormWithinWorkingArea -Form $form
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $form.Tag }
    return $null
}
