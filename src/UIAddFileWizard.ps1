# ==============================================================================
#  UIAddFileWizard.ps1  (UI runspace only)
#
#  Three focused screens:
#
#     1. when should it update
#     2. which Excel workbooks should refresh, and in what order
#     3. name and review the complete rule
#
#  Folder triggers expand in place. Workbook actions reuse the full action
#  editor, so each workbook has its own query selection and refresh settings.
#  Multiple actions run from top to bottom under the one trigger.
#
#  The result is an ordinary rule, saved to the same rules.json. Nothing about
#  it is special, so it can be edited afterwards like any other.
# ==============================================================================

Set-StrictMode -Version 1.0

function New-WizardHeading {
    param([string]$Text, [int]$Y)
    $label = New-Object System.Windows.Forms.Label
    $label.Text     = $Text
    $label.Location = New-Object System.Drawing.Point(28, $Y)
    $label.Size     = New-Object System.Drawing.Size(520, 28)
    $label.Font     = (Get-UiFont 13 'Regular')
    return $label
}

function New-WizardNote {
    param([string]$Text, [int]$Y, [int]$Height = 34)
    $label = New-FormWrappedLabel -Text $Text -X 28 -Y $Y -Width 520
    $label.MinimumSize = New-Object System.Drawing.Size(520, $Height)
    return $label
}

function New-WizardChoice {
    <#
        A radio button written as a choice with a consequence, not a label.
        The second line is what actually helps someone decide.
    #>
    param($Parent, [string]$Title, [string]$Detail, [int]$Y, [bool]$Checked = $false)

    $radio = New-Object System.Windows.Forms.RadioButton
    $radio.Text     = $Title
    $radio.Location = New-Object System.Drawing.Point(30, $Y)
    $radio.Size     = New-Object System.Drawing.Size(500, 24)
    $radio.Font     = (Get-UiFont 10 'Bold')
    $radio.Checked  = $Checked
    $Parent.Controls.Add($radio)

    $note = New-FormWrappedLabel -Text $Detail -X 52 -Y ($Y + 24) -Width 480
    $Parent.Controls.Add($note)
    $radio.Tag = $note

    return $radio
}

function Invoke-AddFileWizard {
    <#  Adds one rule, or does nothing. Saves and refreshes both views.  #>

    $form = New-Object System.Windows.Forms.Form
    $form.Text                = 'New trigger rule'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize          = New-Object System.Drawing.Size(580, 400)
    $form.FormBorderStyle     = 'FixedDialog'
    $form.StartPosition       = 'CenterParent'
    $form.MaximizeBox         = $false
    $form.MinimizeBox         = $false
    $form.Font                = (Get-UiFont)
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $step1 = New-Object System.Windows.Forms.Panel
    $step2 = New-Object System.Windows.Forms.Panel
    $step3 = New-Object System.Windows.Forms.Panel
    foreach ($panel in @($step1, $step2, $step3)) {
        $panel.Location = New-Object System.Drawing.Point(0, 0)
        $panel.Size     = New-Object System.Drawing.Size(580, 340)
        $panel.Visible  = $false
        $form.Controls.Add($panel)
    }

    # Keep mutable wizard data in shared reference objects. Event closures have
    # independent script scopes, so scalar variables are not reliable state.
    $wizardState = @{
        Step         = 1
        ShowStep     = $null
        SuggestedName = ''
        NameEdited    = $false
        UpdatingName  = $false
        RecentApplicable = $true
        Chrome         = @{}
        # GetNewClosure gives the click handler its own script scope. Keep the
        # actual configuration reference here so Add it modifies the live
        # object instead of looking for an empty $script:UiConfig in that scope.
        Config   = $script:UiConfig
    }
    $workbookState = @{ Actions = @() }

    # Physical panel step1 is displayed as logical step 2. Keeping the panel
    # variables stable avoids rebuilding the trigger controls below.
    $step1.Controls.Add((New-WizardHeading 'Which workbooks should refresh?' 20))
    $step1.Controls.Add((New-WizardNote 'Add one or more workbooks. They refresh from top to bottom when this trigger runs.' 54 36))

    $lvWorkbooks = New-Object System.Windows.Forms.ListView
    $lvWorkbooks.Location      = New-Object System.Drawing.Point(28, 98)
    $lvWorkbooks.Size          = New-Object System.Drawing.Size(524, 190)
    $lvWorkbooks.View          = 'Details'
    $lvWorkbooks.FullRowSelect = $true
    $lvWorkbooks.MultiSelect   = $false
    $lvWorkbooks.HideSelection = $false
    [void]$lvWorkbooks.Columns.Add('Order', 54)
    [void]$lvWorkbooks.Columns.Add('Workbook', 298)
    [void]$lvWorkbooks.Columns.Add('Queries', 148)
    $step1.Controls.Add($lvWorkbooks)

    $btnAddWorkbook = New-FormAutoButton -Text 'Add workbook...' -MinimumWidth 126
    $btnEditWorkbook = New-FormAutoButton -Text 'Edit...' -MinimumWidth 82
    $btnRemoveWorkbook = New-FormAutoButton -Text 'Remove' -MinimumWidth 82
    $btnWorkbookUp = New-FormAutoButton -Text 'Move up' -MinimumWidth 86
    $btnWorkbookDown = New-FormAutoButton -Text 'Move down' -MinimumWidth 92
    foreach ($button in @($btnAddWorkbook, $btnEditWorkbook, $btnRemoveWorkbook, $btnWorkbookUp, $btnWorkbookDown)) {
        $step1.Controls.Add($button)
    }
    $btnAddWorkbook.Location    = New-Object System.Drawing.Point(28, 300)
    $btnEditWorkbook.Location   = New-Object System.Drawing.Point(162, 300)
    $btnRemoveWorkbook.Location = New-Object System.Drawing.Point(252, 300)
    $btnWorkbookUp.Location     = New-Object System.Drawing.Point(342, 300)
    $btnWorkbookDown.Location   = New-Object System.Drawing.Point(436, 300)

    $lblWorkbookHint = New-WizardNote 'Each workbook can refresh all queries or its own selected queries.' 340 34
    $step1.Controls.Add($lblWorkbookHint)

    $updateWorkbookButtons = {
        $hasSelection = ($lvWorkbooks.SelectedIndices.Count -gt 0)
        $index = $(if ($hasSelection) { [int]$lvWorkbooks.SelectedIndices[0] } else { -1 })
        $btnEditWorkbook.Enabled = $hasSelection
        $btnRemoveWorkbook.Enabled = $hasSelection
        $btnWorkbookUp.Enabled = ($hasSelection -and $index -gt 0)
        $btnWorkbookDown.Enabled = ($hasSelection -and $index -lt (@($workbookState.Actions).Count - 1))
        $lblWorkbookHint.Text = $(if (@($workbookState.Actions).Count -eq 0) {
            'Add at least one workbook. Query selection is configured separately for each workbook.'
        } else {
            '{0} workbook(s). They will refresh in the order shown.' -f @($workbookState.Actions).Count
        })
    }.GetNewClosure()

    $refreshWorkbookList = {
        $selectedIndex = $(if ($lvWorkbooks.SelectedIndices.Count -gt 0) { [int]$lvWorkbooks.SelectedIndices[0] } else { -1 })
        $lvWorkbooks.BeginUpdate()
        try {
            $lvWorkbooks.Items.Clear()
            for ($i = 0; $i -lt @($workbookState.Actions).Count; $i++) {
                $action = $workbookState.Actions[$i]
                $item = New-Object System.Windows.Forms.ListViewItem([string]($i + 1))
                [void]$item.SubItems.Add((Split-Path -Leaf ([string]$action.path)))
                $queryText = $(if ([string]$action.refreshMethod -eq 'SelectedQueries') {
                    '{0} selected' -f @($action.selectedQueries).Count
                } else { 'All' })
                [void]$item.SubItems.Add($queryText)
                $item.Tag = $action
                [void]$lvWorkbooks.Items.Add($item)
            }
            if ($selectedIndex -ge 0 -and $selectedIndex -lt $lvWorkbooks.Items.Count) {
                $lvWorkbooks.Items[$selectedIndex].Selected = $true
            }
        }
        finally { $lvWorkbooks.EndUpdate() }
        & $updateWorkbookButtons
        if ($null -ne $wizardState.ShowStep) { & $wizardState.ShowStep }
    }.GetNewClosure()

    $addWorkbook = {
        $action = Show-ActionEditor -Action (Get-DefaultAction) -AppSettings $wizardState.Config['appSettings']
        if ($null -eq $action) { return }
        if ((Find-DuplicateWorkbookAction -Actions @($workbookState.Actions) -CandidatePath ([string]$action.path)) -ge 0) {
            [System.Windows.Forms.MessageBox]::Show('That workbook is already in this rule. Select it and use Edit instead.',
                'Add workbook', 'OK', 'Information') | Out-Null
            return
        }
        $items = New-Object System.Collections.ArrayList
        foreach ($existing in @($workbookState.Actions)) { [void]$items.Add($existing) }
        [void]$items.Add($action)
        $workbookState.Actions = @($items.ToArray())
        if (@($workbookState.Actions).Count -eq 1 -and -not [bool]$wizardState.NameEdited) {
            $wizardState.SuggestedName = [System.IO.Path]::GetFileNameWithoutExtension([string]$action.path)
        }
        & $refreshWorkbookList
        if ($lvWorkbooks.Items.Count -gt 0) { $lvWorkbooks.Items[$lvWorkbooks.Items.Count - 1].Selected = $true }
    }.GetNewClosure()
    $btnAddWorkbook.Add_Click($addWorkbook)

    $editWorkbook = {
        if ($lvWorkbooks.SelectedIndices.Count -eq 0) { return }
        $index = [int]$lvWorkbooks.SelectedIndices[0]
        $updated = Show-ActionEditor -Action $workbookState.Actions[$index] -AppSettings $wizardState.Config['appSettings']
        if ($null -eq $updated) { return }
        if ((Find-DuplicateWorkbookAction -Actions @($workbookState.Actions) -CandidatePath ([string]$updated.path) -ExcludeIndex $index) -ge 0) {
            [System.Windows.Forms.MessageBox]::Show('That workbook is already in this rule.', 'Edit workbook', 'OK', 'Information') | Out-Null
            return
        }
        $workbookState.Actions[$index] = $updated
        & $refreshWorkbookList
        $lvWorkbooks.Items[$index].Selected = $true
    }.GetNewClosure()
    $btnEditWorkbook.Add_Click($editWorkbook)
    $lvWorkbooks.Add_DoubleClick($editWorkbook)

    $lvWorkbooks.Add_SelectedIndexChanged({ & $updateWorkbookButtons }.GetNewClosure())
    $btnRemoveWorkbook.Add_Click({
        if ($lvWorkbooks.SelectedIndices.Count -eq 0) { return }
        $index = [int]$lvWorkbooks.SelectedIndices[0]
        $items = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt @($workbookState.Actions).Count; $i++) {
            if ($i -ne $index) { [void]$items.Add($workbookState.Actions[$i]) }
        }
        $workbookState.Actions = @($items.ToArray())
        & $refreshWorkbookList
    }.GetNewClosure())

    $moveWorkbook = {
        param([int]$Direction)
        if ($lvWorkbooks.SelectedIndices.Count -eq 0) { return }
        $from = [int]$lvWorkbooks.SelectedIndices[0]
        $to = $from + $Direction
        if ($to -lt 0 -or $to -ge @($workbookState.Actions).Count) { return }
        $temp = $workbookState.Actions[$from]
        $workbookState.Actions[$from] = $workbookState.Actions[$to]
        $workbookState.Actions[$to] = $temp
        & $refreshWorkbookList
        $lvWorkbooks.Items[$to].Selected = $true
    }.GetNewClosure()
    $btnWorkbookUp.Add_Click({ & $moveWorkbook -1 }.GetNewClosure())
    $btnWorkbookDown.Add_Click({ & $moveWorkbook 1 }.GetNewClosure())
    & $refreshWorkbookList

    # ---------------------------------------------------------------- step 2 --

    $step2.Controls.Add((New-WizardHeading 'What should trigger the update?' 26))

    $step2.Controls.Add((New-FormLabel 'Trigger:' 28 70 92))
    $cboTriggerType = New-Object System.Windows.Forms.ComboBox
    $cboTriggerType.Location      = New-Object System.Drawing.Point(124, 68)
    $cboTriggerType.Size          = New-Object System.Drawing.Size(380, 24)
    $cboTriggerType.DropDownStyle = 'DropDownList'
    foreach ($type in (Get-TriggerTypeList)) {
        [void]$cboTriggerType.Items.Add((Get-TriggerTypeLabel $type))
    }
    $cboTriggerType.SelectedIndex = 3
    $step2.Controls.Add($cboTriggerType)

    $lblTriggerHelp = New-WizardNote '' 102 36
    $step2.Controls.Add($lblTriggerHelp)

    # All six trigger types remain available. Selecting one opens only the
    # fields that have meaning for that trigger, like a large accordion panel.
    $triggerPanels = @{}
    foreach ($key in @('Folder', 'Specific', 'Schedule', 'Logon', 'Manual')) {
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Location = New-Object System.Drawing.Point(28, 140)
        $panel.Size     = New-Object System.Drawing.Size(524, 230)
        $panel.Visible  = $false
        $step2.Controls.Add($panel)
        $triggerPanels[$key] = $panel
    }

    $gbFolderDetails = New-Object System.Windows.Forms.GroupBox
    $gbFolderDetails.Text     = 'Which files in the folder should count?'
    $gbFolderDetails.Location = New-Object System.Drawing.Point(0, 0)
    $gbFolderDetails.Size     = New-Object System.Drawing.Size(524, 184)
    $triggerPanels['Folder'].Controls.Add($gbFolderDetails)

    $gbFolderDetails.Controls.Add((New-FormLabel 'Folder:' 12 24 104))
    $txtFolder = New-Object System.Windows.Forms.TextBox
    $txtFolder.Location = New-Object System.Drawing.Point(120, 22)
    $txtFolder.Size     = New-Object System.Drawing.Size(264, 24)
    $txtFolder.ReadOnly = $true
    $gbFolderDetails.Controls.Add($txtFolder)

    $btnFolder = New-FormAutoButton -Text 'Choose...' -MinimumWidth 88
    $gbFolderDetails.Controls.Add($btnFolder)
    $btnFolder.Location = New-Object System.Drawing.Point(392, 20)
    $btnFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Choose the folder to watch'
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtFolder.Text = $dialog.SelectedPath }
    }.GetNewClosure())

    $gbFolderDetails.Controls.Add((New-FormLabel 'File type:' 12 58 104))
    $cboIncomingType = New-Object System.Windows.Forms.ComboBox
    $cboIncomingType.Location      = New-Object System.Drawing.Point(120, 56)
    $cboIncomingType.Size          = New-Object System.Drawing.Size(170, 24)
    $cboIncomingType.DropDownStyle = 'DropDownList'
    foreach ($label in @('CSV (*.csv)', 'Text (*.txt)', 'Excel files', 'Any file (*.*)', 'Custom pattern')) {
        [void]$cboIncomingType.Items.Add($label)
    }
    $cboIncomingType.SelectedItem = 'Any file (*.*)'
    $gbFolderDetails.Controls.Add($cboIncomingType)

    $txtIncomingFilter = New-Object System.Windows.Forms.TextBox
    $txtIncomingFilter.Location = New-Object System.Drawing.Point(298, 56)
    $txtIncomingFilter.Size     = New-Object System.Drawing.Size(182, 24)
    $txtIncomingFilter.Text     = '*.*'
    $txtIncomingFilter.Enabled  = $false
    $gbFolderDetails.Controls.Add($txtIncomingFilter)

    $incomingPresets = @{
        'CSV (*.csv)'     = '*.csv'
        'Text (*.txt)'    = '*.txt'
        'Excel files'     = '*.xlsx;*.xlsm;*.xlsb;*.xls'
        'Any file (*.*)'  = '*.*'
    }
    $syncIncomingType = {
        $selected = [string]$cboIncomingType.SelectedItem
        if ($selected -eq 'Custom pattern') { $txtIncomingFilter.Enabled = $true }
        elseif ($incomingPresets.ContainsKey($selected)) {
            $txtIncomingFilter.Text = [string]$incomingPresets[$selected]
            $txtIncomingFilter.Enabled = $false
        }
    }.GetNewClosure()
    $cboIncomingType.Add_SelectedIndexChanged($syncIncomingType)

    $gbFolderDetails.Controls.Add((New-FormLabel 'Name contains:' 12 92 104))
    $txtIncomingContains = New-Object System.Windows.Forms.TextBox
    $txtIncomingContains.Location = New-Object System.Drawing.Point(120, 90)
    $txtIncomingContains.Size     = New-Object System.Drawing.Size(360, 24)
    $gbFolderDetails.Controls.Add($txtIncomingContains)

    $gbFolderDetails.Controls.Add((New-FormLabel 'Ignore if name contains:' 12 126 160))
    $txtIncomingExclude = New-Object System.Windows.Forms.TextBox
    $txtIncomingExclude.Location = New-Object System.Drawing.Point(176, 124)
    $txtIncomingExclude.Size     = New-Object System.Drawing.Size(304, 24)
    $gbFolderDetails.Controls.Add($txtIncomingExclude)

    $lblIncomingHint = New-FormLabel 'Leave the name fields blank when the file type alone is enough.' 120 154 360 20
    $lblIncomingHint.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
    $gbFolderDetails.Controls.Add($lblIncomingHint)

    # Specific-file change trigger: this is the incoming/source file to watch,
    # not the workbook chosen on step 1.
    $triggerPanels['Specific'].Controls.Add((New-FormLabel 'File to watch:' 0 6 112))
    $txtWatchedFile = New-Object System.Windows.Forms.TextBox
    $txtWatchedFile.Location = New-Object System.Drawing.Point(120, 4)
    $txtWatchedFile.Size     = New-Object System.Drawing.Size(300, 24)
    $txtWatchedFile.ReadOnly = $true
    $triggerPanels['Specific'].Controls.Add($txtWatchedFile)
    $btnWatchedFile = New-FormAutoButton -Text 'Browse...' -MinimumWidth 92
    $btnWatchedFile.Location = New-Object System.Drawing.Point(428, 2)
    $triggerPanels['Specific'].Controls.Add($btnWatchedFile)
    $specificHint = New-WizardNote 'The workbook refreshes whenever this exact file is modified or replaced.' 0 42
    $specificHint.Location = New-Object System.Drawing.Point(0, 42)
    $specificHint.Size = New-Object System.Drawing.Size(520, 40)
    $triggerPanels['Specific'].Controls.Add($specificHint)
    $btnWatchedFile.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'All files|*.*'
        $dialog.Title = 'Choose the file whose change should trigger the refresh'
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtWatchedFile.Text = $dialog.FileName
        }
    }.GetNewClosure())

    # Scheduled trigger details.
    $triggerPanels['Schedule'].Controls.Add((New-FormLabel 'Run at:' 0 6 92))
    $timePicker = New-Object System.Windows.Forms.DateTimePicker
    $timePicker.Format       = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $timePicker.CustomFormat = 'HH:mm'
    $timePicker.ShowUpDown   = $true
    $timePicker.Location     = New-Object System.Drawing.Point(96, 4)
    $timePicker.Size         = New-Object System.Drawing.Size(90, 24)
    $timePicker.Value        = [datetime]::Today.AddHours(7).AddMinutes(30)
    $triggerPanels['Schedule'].Controls.Add($timePicker)
    $triggerPanels['Schedule'].Controls.Add((New-FormLabel '24-hour clock' 198 6 180))
    $triggerPanels['Schedule'].Controls.Add((New-FormLabel 'Days:' 0 44 92))
    $dayNames = @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday')
    $dayChecks = @{}
    $dayX = 72
    foreach ($day in $dayNames) {
        $check = New-Object System.Windows.Forms.CheckBox
        $check.Text     = $day.Substring(0, 3)
        $check.Location = New-Object System.Drawing.Point($dayX, 42)
        $check.Size     = New-Object System.Drawing.Size(60, 22)
        $check.Checked  = ($day -ne 'Saturday' -and $day -ne 'Sunday')
        $triggerPanels['Schedule'].Controls.Add($check)
        $dayChecks[$day] = $check
        $dayX += 62
    }
    $scheduleRequirement = New-FormWrappedLabel `
        -Text 'Scheduled refreshes run only while this PC is turned on and you are signed in to Windows. Missed times are not replayed.' `
        -X 0 -Y 82 -Width 510 -Color ([System.Drawing.Color]::FromArgb(105, 105, 105))
    $triggerPanels['Schedule'].Controls.Add($scheduleRequirement)

    # Windows-logon trigger details.
    $rdoLogonAsk = New-WizardChoice -Parent ($triggerPanels['Logon']) -Y 0 -Checked $true `
        -Title 'Ask me before refreshing  (recommended)' `
        -Detail 'The dashboard asks after the Windows logon delay.'
    $rdoLogonAuto = New-WizardChoice -Parent ($triggerPanels['Logon']) -Y 62 `
        -Title 'Refresh automatically, without asking' `
        -Detail 'The job starts after the delay configured in Settings.'

    # This is a rule-wide choice, but it belongs beside the workbook selection:
    # the user is deciding whether these target workbooks are recent enough to
    # warrant a confirmation before Excel opens.
    $gbRecentRefresh = New-Object System.Windows.Forms.GroupBox
    $gbRecentRefresh.Text = 'Recent refresh protection'
    $gbRecentRefresh.Location = New-Object System.Drawing.Point(28, 376)
    $gbRecentRefresh.Size = New-Object System.Drawing.Size(524, 98)
    $step1.Controls.Add($gbRecentRefresh)

    $chkRecentRefresh = New-Object System.Windows.Forms.CheckBox
    $chkRecentRefresh.Text = 'Ask if queries were refreshed within'
    $chkRecentRefresh.Location = New-Object System.Drawing.Point(14, 22)
    $chkRecentRefresh.Size = New-Object System.Drawing.Size(292, 24)
    $gbRecentRefresh.Controls.Add($chkRecentRefresh)

    $numRecentRefresh = New-Object System.Windows.Forms.NumericUpDown
    $numRecentRefresh.Location = New-Object System.Drawing.Point(310, 21)
    $numRecentRefresh.Size = New-Object System.Drawing.Size(70, 24)
    $numRecentRefresh.Minimum = 1
    $numRecentRefresh.Maximum = 10080
    $numRecentRefresh.Value = 30
    $numRecentRefresh.Enabled = $false
    $gbRecentRefresh.Controls.Add($numRecentRefresh)

    $cboRecentUnit = New-Object System.Windows.Forms.ComboBox
    $cboRecentUnit.Location = New-Object System.Drawing.Point(388, 20)
    $cboRecentUnit.Size = New-Object System.Drawing.Size(112, 24)
    $cboRecentUnit.DropDownStyle = 'DropDownList'
    [void]$cboRecentUnit.Items.Add('minutes')
    [void]$cboRecentUnit.Items.Add('hours')
    $cboRecentUnit.SelectedIndex = 0
    $cboRecentUnit.Enabled = $false
    $gbRecentRefresh.Controls.Add($cboRecentUnit)

    $recentRefreshHint = New-FormWrappedLabel `
        -Text ('Automatic runs only.' + [Environment]::NewLine +
            'If any workbook above was refreshed recently, ask before refreshing it again to prevent unnecessary refreshes.') `
        -X 36 -Y 51 -Width 458 `
        -Color ([System.Drawing.Color]::FromArgb(105, 105, 105))

    $gbRecentRefresh.Controls.Add($recentRefreshHint)
    $chkRecentRefresh.Add_CheckedChanged({
        $numRecentRefresh.Enabled = $chkRecentRefresh.Checked
        $cboRecentUnit.Enabled = $chkRecentRefresh.Checked
    }.GetNewClosure())

    $manualHint = New-WizardNote 'This rule runs only when you press Run Now, or choose Run All Manual Rules from the tray icon.' 0 60
    $manualHint.Location = New-Object System.Drawing.Point(0, 4)
    $manualHint.Size = New-Object System.Drawing.Size(520, 54)
    $triggerPanels['Manual'].Controls.Add($manualHint)

    $getSelectedTriggerType = {
        if ($cboTriggerType.SelectedIndex -lt 0) { return 'Scheduled' }
        return (Get-TriggerTypeList)[$cboTriggerType.SelectedIndex]
    }.GetNewClosure()

    $syncStep2 = {
        $selectedType = & $getSelectedTriggerType
        $panelKey = 'Manual'
        if (Test-TriggerUsesFolder $selectedType) { $panelKey = 'Folder' }
        elseif ($selectedType -eq 'FileChangedSpecific') { $panelKey = 'Specific' }
        elseif (Test-TriggerIsScheduled $selectedType) { $panelKey = 'Schedule' }
        elseif (Test-TriggerIsLogon $selectedType) { $panelKey = 'Logon' }
        foreach ($key in @($triggerPanels.Keys)) { $triggerPanels[$key].Visible = ($key -eq $panelKey) }
        $wizardState.RecentApplicable = ($panelKey -ne 'Manual')
        $gbRecentRefresh.Visible = [bool]$wizardState.RecentApplicable
        $lblTriggerHelp.Text = switch ($selectedType) {
            'FileCreated'         { 'Refresh when a new matching file is added to the selected folder.' }
            'FileChangedSpecific' { 'Refresh when one exact source file is updated or replaced.' }
            'FileChangedAny'      { 'Refresh when any matching existing file in the folder is updated.' }
            'Scheduled'           { 'Refresh at a chosen time on the selected days.' }
            'Logon'               { 'Refresh after you sign in to Windows.' }
            default               { 'Do nothing automatically; run it from the dashboard or tray.' }
        }
        if ($null -ne $wizardState.ShowStep) { & $wizardState.ShowStep }
    }.GetNewClosure()
    $cboTriggerType.Add_SelectedIndexChanged($syncStep2)
    & $syncStep2

    # ---------------------------------------------------------------- step 3 --

    $step3.Controls.Add((New-WizardHeading 'Name and review this rule' 18))

    $gbRuleName = New-Object System.Windows.Forms.GroupBox
    $gbRuleName.Text     = 'Rule name'
    $gbRuleName.Location = New-Object System.Drawing.Point(28, 54)
    $gbRuleName.Size     = New-Object System.Drawing.Size(524, 66)
    $step3.Controls.Add($gbRuleName)

    $txtRuleName = New-Object System.Windows.Forms.TextBox
    $txtRuleName.Location = New-Object System.Drawing.Point(14, 25)
    $txtRuleName.Size     = New-Object System.Drawing.Size(494, 24)
    $gbRuleName.Controls.Add($txtRuleName)
    $txtRuleName.Add_TextChanged({
        if (-not [bool]$wizardState.UpdatingName) { $wizardState.NameEdited = $true }
    }.GetNewClosure())

    $gbReviewWorkbook = New-Object System.Windows.Forms.GroupBox
    $gbReviewWorkbook.Text     = 'Workbooks to refresh (in order)'
    $gbReviewWorkbook.Location = New-Object System.Drawing.Point(28, 128)
    $gbReviewWorkbook.Size     = New-Object System.Drawing.Size(524, 158)
    $step3.Controls.Add($gbReviewWorkbook)

    $lvReviewWorkbooks = New-Object System.Windows.Forms.ListView
    $lvReviewWorkbooks.Location      = New-Object System.Drawing.Point(12, 22)
    $lvReviewWorkbooks.Size          = New-Object System.Drawing.Size(500, 124)
    $lvReviewWorkbooks.View          = 'Details'
    $lvReviewWorkbooks.FullRowSelect = $true
    $lvReviewWorkbooks.HeaderStyle   = 'Nonclickable'
    [void]$lvReviewWorkbooks.Columns.Add('Order', 54)
    [void]$lvReviewWorkbooks.Columns.Add('Workbook', 294)
    [void]$lvReviewWorkbooks.Columns.Add('Queries', 128)
    $gbReviewWorkbook.Controls.Add($lvReviewWorkbooks)

    $gbReviewTrigger = New-Object System.Windows.Forms.GroupBox
    $gbReviewTrigger.Text     = 'Trigger settings'
    $gbReviewTrigger.Location = New-Object System.Drawing.Point(28, 294)
    $gbReviewTrigger.Size     = New-Object System.Drawing.Size(524, 228)
    $step3.Controls.Add($gbReviewTrigger)

    $reviewTable = New-Object System.Windows.Forms.TableLayoutPanel
    $reviewTable.Location    = New-Object System.Drawing.Point(12, 22)
    $reviewTable.Size        = New-Object System.Drawing.Size(500, 194)
    $reviewTable.ColumnCount = 2
    $reviewTable.RowCount    = 0
    $reviewTable.AutoScroll  = $true
    $reviewTable.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    $captionColumn = New-Object System.Windows.Forms.ColumnStyle
    $captionColumn.SizeType = [System.Windows.Forms.SizeType]::Absolute
    $captionColumn.Width = 132
    $valueColumn = New-Object System.Windows.Forms.ColumnStyle
    $valueColumn.SizeType = [System.Windows.Forms.SizeType]::Percent
    $valueColumn.Width = 100
    [void]$reviewTable.ColumnStyles.Add($captionColumn)
    [void]$reviewTable.ColumnStyles.Add($valueColumn)
    $gbReviewTrigger.Controls.Add($reviewTable)

    $addReviewRow = {
        param([string]$Caption, [string]$Value)
        $row = [int]$reviewTable.RowCount
        $reviewTable.RowCount = $row + 1
        $rowStyle = New-Object System.Windows.Forms.RowStyle
        $rowStyle.SizeType = [System.Windows.Forms.SizeType]::AutoSize
        [void]$reviewTable.RowStyles.Add($rowStyle)

        $captionLabel = New-Object System.Windows.Forms.Label
        $captionLabel.Text = $Caption
        $captionLabel.AutoSize = $true
        $captionLabel.Font = Get-UiFont 9 'Bold'
        $captionLabel.Margin = New-Object System.Windows.Forms.Padding(8, 7, 6, 7)
        $reviewTable.Controls.Add($captionLabel, 0, $row)

        $valueLabel = New-FormWrappedLabel -Text $Value -X 0 -Y 0 -Width 330 `
            -Color ([System.Drawing.SystemColors]::ControlText)
        $valueLabel.Margin = New-Object System.Windows.Forms.Padding(8, 7, 6, 7)
        $reviewTable.Controls.Add($valueLabel, 1, $row)
    }.GetNewClosure()

    $reviewNote = New-WizardNote ('The refresh uses a separate Excel instance in the background. ' +
        'You can change the name, trigger, workbook order and advanced settings later with Edit.') 532 38
    $step3.Controls.Add($reviewNote)

    $populateReview = {
        if (-not [bool]$wizardState.NameEdited) {
            $suggested = ''
            if (@($workbookState.Actions).Count -gt 0) {
                $suggested = [System.IO.Path]::GetFileNameWithoutExtension([string]$workbookState.Actions[0].path)
                if (@($workbookState.Actions).Count -gt 1) {
                    $suggested = '{0} and {1} more' -f $suggested, (@($workbookState.Actions).Count - 1)
                }
            }
            $wizardState.UpdatingName = $true
            try { $txtRuleName.Text = $suggested }
            finally { $wizardState.UpdatingName = $false }
        }

        $lvReviewWorkbooks.BeginUpdate()
        try {
            $lvReviewWorkbooks.Items.Clear()
            for ($i = 0; $i -lt @($workbookState.Actions).Count; $i++) {
                $action = $workbookState.Actions[$i]
                $item = New-Object System.Windows.Forms.ListViewItem([string]($i + 1))
                [void]$item.SubItems.Add((Split-Path -Leaf ([string]$action.path)))
                $queryText = $(if ([string]$action.refreshMethod -eq 'SelectedQueries') {
                    '{0} selected' -f @($action.selectedQueries).Count
                } else { 'All' })
                [void]$item.SubItems.Add($queryText)
                [void]$lvReviewWorkbooks.Items.Add($item)
            }
        }
        finally { $lvReviewWorkbooks.EndUpdate() }

        $reviewTable.SuspendLayout()
        try {
            foreach ($control in @($reviewTable.Controls)) { $control.Dispose() }
            $reviewTable.Controls.Clear()
            $reviewTable.RowStyles.Clear()
            $reviewTable.RowCount = 0

            $selectedType = & $getSelectedTriggerType
            & $addReviewRow 'Trigger' (Get-TriggerTypeLabel $selectedType)
            if (Test-TriggerUsesFolder $selectedType) {
                & $addReviewRow 'Folder' ([string]$txtFolder.Text)
                & $addReviewRow 'File pattern' ([string]$txtIncomingFilter.Text)
                if (-not [string]::IsNullOrWhiteSpace($txtIncomingContains.Text)) {
                    & $addReviewRow 'Name contains' ([string]$txtIncomingContains.Text)
                }
                if (-not [string]::IsNullOrWhiteSpace($txtIncomingExclude.Text)) {
                    & $addReviewRow 'Ignore names with' ([string]$txtIncomingExclude.Text)
                }
            }
            elseif ($selectedType -eq 'FileChangedSpecific') {
                & $addReviewRow 'File to watch' ([string]$txtWatchedFile.Text)
            }
            elseif (Test-TriggerIsScheduled $selectedType) {
                $days = @($dayNames | Where-Object { $dayChecks[$_].Checked })
                & $addReviewRow 'Run at' ($timePicker.Value.ToString('HH:mm'))
                & $addReviewRow 'Days' ($days -join ', ')
            }
            elseif (Test-TriggerIsLogon $selectedType) {
                $behavior = $(if ($rdoLogonAuto.Checked) { 'Refresh automatically' } else { 'Ask before refreshing' })
                & $addReviewRow 'At logon' $behavior
            }
            else {
                & $addReviewRow 'Starts from' 'Run Now on the dashboard or tray menu'
            }
            if ([bool]$wizardState.RecentApplicable -and $chkRecentRefresh.Checked) {
                $windowValue = [int]$numRecentRefresh.Value
                $windowUnit = [string]$cboRecentUnit.SelectedItem
                & $addReviewRow 'Recent refresh protection' ('Ask before refreshing within {0} {1}' -f $windowValue, $windowUnit)
            }
        }
        finally { $reviewTable.ResumeLayout($true) }

        $reviewNote.Text = $(if (Test-TriggerIsScheduled $selectedType) {
            'Scheduled refreshes require this PC to be turned on and signed in to Windows. Missed times are not replayed.' +
            [Environment]::NewLine + [Environment]::NewLine +
            'You can change the name, trigger, workbook order and advanced settings later with Edit.'
        } else {
            'The refresh uses a separate Excel instance in the background. ' +
            'You can change the name, trigger, workbook order and advanced settings later with Edit.'
        })

        # The summary table should be only as tall as its content. In
        # particular, three short scheduled rows must not leave one large empty
        # Days cell. Folder triggers may use the full available height.
        $tableHeight = [Math]::Min(194, [Math]::Max(76, ([int]$reviewTable.RowCount * 36) + 4))
        $reviewTable.Height = $tableHeight
        $gbReviewTrigger.Height = $tableHeight + 34
        $reviewNote.Location = New-Object System.Drawing.Point(28, ($gbReviewTrigger.Bottom + 10))

        $desiredHeight = [Math]::Max(520, ($reviewNote.Bottom + 60))
        $chromeY = $desiredHeight - 48
        $form.ClientSize = New-Object System.Drawing.Size(580, $desiredHeight)
        foreach ($panel in @($step1, $step2, $step3)) {
            $panel.Size = New-Object System.Drawing.Size(580, ($chromeY - 8))
        }
        $chrome = $wizardState.Chrome
        $chrome['Step'].Location   = New-Object System.Drawing.Point(28, ($chromeY + 4))
        $chrome['Back'].Location   = New-Object System.Drawing.Point(258, $chromeY)
        $chrome['Next'].Location   = New-Object System.Drawing.Point(362, $chromeY)
        $chrome['Cancel'].Location = New-Object System.Drawing.Point(478, $chromeY)
    }.GetNewClosure()

    # ---------------------------------------------------------------- chrome --

    $lblStep = New-Object System.Windows.Forms.Label
    $lblStep.Location  = New-Object System.Drawing.Point(28, 356)
    $lblStep.Size      = New-Object System.Drawing.Size(180, 22)
    $lblStep.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $form.Controls.Add($lblStep)

    $btnBack = New-FormAutoButton -Text 'Back' -MinimumWidth 96
    $form.Controls.Add($btnBack)
    $btnBack.Location = New-Object System.Drawing.Point(258, 352)

    $btnNext = New-FormAutoButton -Text 'Next' -MinimumWidth 110
    $form.Controls.Add($btnNext)
    $btnNext.Location = New-Object System.Drawing.Point(362, 352)

    $btnCancel = New-FormAutoButton -Text 'Cancel' -MinimumWidth 96
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $btnCancel.Location = New-Object System.Drawing.Point(478, 352)
    $form.CancelButton = $btnCancel
    $wizardState.Chrome = @{ Step = $lblStep; Back = $btnBack; Next = $btnNext; Cancel = $btnCancel }

    $showStep = {
        $currentStep = [int]$wizardState.Step
        $clientHeight = 400
        if ($currentStep -eq 1) {
            $selectedType = & $getSelectedTriggerType
            if (Test-TriggerUsesFolder $selectedType) { $clientHeight = 570 }
            elseif (Test-TriggerIsScheduled $selectedType) { $clientHeight = 500 }
            elseif (Test-TriggerIsLogon $selectedType) { $clientHeight = 500 }
            elseif ($selectedType -eq 'FileChangedSpecific') { $clientHeight = 460 }
            else { $clientHeight = 420 }
        }
        elseif ($currentStep -eq 2) {
            $clientHeight = $(if ([bool]$wizardState.RecentApplicable) { 565 } else { 470 })
        }
        elseif ($currentStep -eq 3) { $clientHeight = 640 }
        $chromeY = $clientHeight - 48
        $form.ClientSize = New-Object System.Drawing.Size(580, $clientHeight)
        foreach ($panel in @($step1, $step2, $step3)) {
            $panel.Size = New-Object System.Drawing.Size(580, ($chromeY - 8))
        }
        $lblStep.Location  = New-Object System.Drawing.Point(28, ($chromeY + 4))
        $btnBack.Location  = New-Object System.Drawing.Point(258, $chromeY)
        $btnNext.Location  = New-Object System.Drawing.Point(362, $chromeY)
        $btnCancel.Location = New-Object System.Drawing.Point(478, $chromeY)
        $step1.Visible = ($currentStep -eq 2)
        $step2.Visible = ($currentStep -eq 1)
        $step3.Visible = ($currentStep -eq 3)
        $btnBack.Enabled = ($currentStep -gt 1)
        $btnNext.Enabled = (($currentStep -ne 2) -or (@($workbookState.Actions).Count -gt 0))
        $btnNext.Text = $(if ($currentStep -eq 3) { 'Add it' } else { 'Next' })
        $lblStep.Text = ('Step {0} of 3' -f $currentStep)
        if ($currentStep -eq 3) { & $populateReview }
    }.GetNewClosure()
    $wizardState.ShowStep = $showStep

    $btnBack.Add_Click({
        if ([int]$wizardState.Step -gt 1) { $wizardState.Step = [int]$wizardState.Step - 1 }
        & $showStep
    }.GetNewClosure())

    $btnNext.Add_Click({
        if ([int]$wizardState.Step -eq 1) {
            $selectedType = & $getSelectedTriggerType
            if (Test-TriggerIsScheduled $selectedType) {
                $chosen = @($dayNames | Where-Object { $dayChecks[$_].Checked })
                if ($chosen.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show('Tick at least one day.', 'Add a file', 'OK', 'Information') | Out-Null
                    return
                }
            }
            if ((Test-TriggerUsesFolder $selectedType) -and [string]::IsNullOrWhiteSpace($txtFolder.Text)) {
                [System.Windows.Forms.MessageBox]::Show('Choose the folder to watch.', 'Add a file', 'OK', 'Information') | Out-Null
                return
            }
            if ((Test-TriggerUsesFolder $selectedType) -and [string]::IsNullOrWhiteSpace($txtIncomingFilter.Text)) {
                [System.Windows.Forms.MessageBox]::Show('Enter a file pattern, for example *.csv or *.*.',
                    'Add a file', 'OK', 'Information') | Out-Null
                return
            }
            if ($selectedType -eq 'FileChangedSpecific' -and
                [string]::IsNullOrWhiteSpace($txtWatchedFile.Text)) {
                [System.Windows.Forms.MessageBox]::Show('Choose the exact file to watch.', 'Add a file', 'OK', 'Information') | Out-Null
                return
            }
            $wizardState.Step = 2
            & $showStep
            return
        }

        if ([int]$wizardState.Step -eq 2) {
            if (@($workbookState.Actions).Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show('Add at least one workbook to refresh.', 'Add a file', 'OK', 'Information') | Out-Null
                return
            }
            $wizardState.Step = 3
            & $showStep
            return
        }

        # ---- build the rule -------------------------------------------------
        $config = $wizardState.Config
        if ($config -isnot [hashtable]) {
            throw [System.InvalidOperationException]::new('The live configuration was not available to the Add wizard.')
        }
        $cleanRuleName = $txtRuleName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($cleanRuleName)) {
            [System.Windows.Forms.MessageBox]::Show('Enter a name for this rule.', 'Add a file', 'OK', 'Information') | Out-Null
            $txtRuleName.Focus()
            return
        }
        if (@($workbookState.Actions).Count -eq 0) { return }
        $rule = New-RuleTemplate -AppSettings $config['appSettings']
        $rule.name = $cleanRuleName

        $selectedType = & $getSelectedTriggerType
        $rule.trigger.type = $selectedType
        if (Test-TriggerIsScheduled $selectedType) {
            $rule.trigger.scheduleTime  = $timePicker.Value.ToString('HH:mm')
            $rule.trigger.scheduleDays  = @($dayNames | Where-Object { $dayChecks[$_].Checked })
        }
        elseif (Test-TriggerUsesFolder $selectedType) {
            $rule.trigger.path   = $txtFolder.Text
            $rule.trigger.filter = $txtIncomingFilter.Text.Trim()
            $rule.trigger.contains = $txtIncomingContains.Text.Trim()
            $rule.trigger.exclude  = $txtIncomingExclude.Text.Trim()
        }
        elseif ($selectedType -eq 'FileChangedSpecific') {
            $rule.trigger.path = $txtWatchedFile.Text
            $rule.trigger.filter = Split-Path -Leaf $txtWatchedFile.Text
        }
        elseif (Test-TriggerIsLogon $selectedType) {
            $rule.trigger.logonBehavior = $(if ($rdoLogonAuto.Checked) { 'Automatic' } else { 'Ask' })
        }
        if ([bool]$wizardState.RecentApplicable -and $chkRecentRefresh.Checked) {
            $recentMinutes = [int]$numRecentRefresh.Value
            if ([string]$cboRecentUnit.SelectedItem -eq 'hours') { $recentMinutes *= 60 }
            $rule.trigger.recentRefreshPromptMinutes = $recentMinutes
        }

        $rule.actions = @($workbookState.Actions)

        $validation = Test-RuleConfiguration -Rule $rule
        if ($validation.Errors.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                ('This rule cannot be saved yet:' + [Environment]::NewLine + [Environment]::NewLine +
                    ($validation.Errors -join [Environment]::NewLine)),
                'Check the rule', 'OK', 'Warning') | Out-Null
            return
        }

        $duplicate = Find-DuplicateRuleDefinition -Candidate $rule -ExistingRules @($config['rules'])
        if ($null -ne $duplicate) {
            [System.Windows.Forms.MessageBox]::Show(
                ('An identical rule already exists: "{0}"' -f [string]$duplicate.name) + [Environment]::NewLine + [Environment]::NewLine +
                'Change the trigger or workbook settings instead of creating a duplicate.',
                'Duplicate rule', 'OK', 'Information') | Out-Null
            return
        }

        $rules = New-Object System.Collections.ArrayList
        foreach ($existing in @($config['rules'])) { [void]$rules.Add($existing) }
        [void]$rules.Add($rule)
        $config['rules'] = @($rules.ToArray())

        if (Save-UiConfiguration) {
            Write-AppLog -Level 'INFO' -RuleName ([string]$rule.name) -Message 'Added from the simple view.'
            Update-RuleList
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    }.GetNewClosure())

    & $showStep
    Set-FormWithinWorkingArea -Form $form
    [void]$form.ShowDialog($script:UiControls.Form)
    $form.Dispose()
}
