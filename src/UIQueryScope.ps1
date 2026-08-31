# ==============================================================================
#  UIQueryScope.ps1  (UI runspace only)
#
#  "Which queries in this workbook get refreshed" used to live three levels
#  down: select the rule, press Edit, find the workbook, find the refresh
#  method, then press a button. Nobody was going to find that.
#
#  It is one question, so it gets one dialog, opened straight from the rule
#  list: everything, or the ones you pick. That is the whole screen.
# ==============================================================================

Set-StrictMode -Version 1.0

function Get-RuleQueryScopeText {
    <#  What the Queries column shows for a rule.  #>
    param([hashtable]$Rule)

    $actions = @($Rule.actions)
    if ($actions.Count -eq 0) { return '' }

    $selectedCounts = New-Object System.Collections.ArrayList
    $allCount = 0
    foreach ($action in $actions) {
        if ([string]$action.refreshMethod -eq 'SelectedQueries') {
            [void]$selectedCounts.Add(@($action.selectedQueries).Count)
        }
        else { $allCount++ }
    }

    if ($selectedCounts.Count -eq 0) { return 'All' }
    if ($allCount -eq 0 -and $selectedCounts.Count -eq 1) { return ('{0} chosen' -f $selectedCounts[0]) }
    if ($allCount -eq 0) { return ('{0} chosen' -f (($selectedCounts.ToArray() | Measure-Object -Sum).Sum)) }
    return 'Mixed'
}

function Show-QueryScopeDialog {
    <#
        Everything, or a chosen few, for one workbook.
        Returns the updated action, or $null if nothing was changed.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Action,
        [Parameter(Mandatory = $true)][hashtable]$AppSettings
    )

    $working = Merge-DefaultValues -Default (Get-DefaultAction) -Value $Action
    $path    = [string]$working.path
    # Shared mutable state is required here: WinForms event handlers execute in
    # their own PowerShell scopes, so assigning a plain $chosen variable inside
    # the Choose click handler would not update the value seen by Save.
    $queryState = @{
        Selected = @(@($working.selectedQueries) | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text                = 'Which queries should be refreshed?'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize          = New-Object System.Drawing.Size(560, 320)
    $form.FormBorderStyle     = 'FixedDialog'
    $form.StartPosition       = 'CenterParent'
    $form.MaximizeBox         = $false
    $form.MinimizeBox         = $false
    $form.Font                = (Get-UiFont)
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Location = New-Object System.Drawing.Point(20, 18)
    $lblFile.Size     = New-Object System.Drawing.Size(516, 22)
    $lblFile.Font     = (Get-UiFont 10 'Bold')
    $lblFile.Text     = Split-Path -Leaf $path
    $lblFile.AutoEllipsis = $true
    $form.Controls.Add($lblFile)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Location  = New-Object System.Drawing.Point(20, 40)
    $lblPath.Size      = New-Object System.Drawing.Size(516, 20)
    $lblPath.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
    $lblPath.Text      = $path
    $lblPath.AutoEllipsis = $true
    $form.Controls.Add($lblPath)

    $rdoAll = New-Object System.Windows.Forms.RadioButton
    $rdoAll.Location = New-Object System.Drawing.Point(22, 76)
    $rdoAll.Size     = New-Object System.Drawing.Size(500, 24)
    $rdoAll.Text     = 'Refresh everything in this workbook'
    $form.Controls.Add($rdoAll)

    $noteAll = New-FormWrappedLabel `
        -Text 'The usual choice, and the one to keep unless a refresh is slow.' `
        -X 44 -Y 100 -Width 490
    $form.Controls.Add($noteAll)

    $rdoSome = New-Object System.Windows.Forms.RadioButton
    $rdoSome.Location = New-Object System.Drawing.Point(22, 130)
    $rdoSome.Size     = New-Object System.Drawing.Size(500, 24)
    $rdoSome.Text     = 'Refresh only the queries I choose'
    $form.Controls.Add($rdoSome)

    $noteSome = New-FormWrappedLabel `
        -Text 'Worth doing when one slow query does not need to run every time.' `
        -X 44 -Y 154 -Width 490
    $form.Controls.Add($noteSome)

    $lblChosen = New-Object System.Windows.Forms.Label
    $lblChosen.Location = New-Object System.Drawing.Point(44, 182)
    $lblChosen.Size     = New-Object System.Drawing.Size(360, 40)
    $form.Controls.Add($lblChosen)

    $btnChoose = New-FormAutoButton -Text 'Choose...' -MinimumWidth 110
    $form.Controls.Add($btnChoose)
    $btnChoose.Location = New-Object System.Drawing.Point(414, 180)

    $lblWarn = New-FormWrappedLabel `
        -Text 'Reading the list opens the workbook in a hidden Excel for a moment. It is never saved by this.' `
        -X 20 -Y 228 -Width 516
    $form.Controls.Add($lblWarn)

    $sync = {
        $rdoAll.Checked = -not $rdoSome.Checked
        $btnChoose.Enabled = $rdoSome.Checked
        $lblChosen.Enabled = $rdoSome.Checked
        $lblChosen.Text = $(
            if (@($queryState.Selected).Count -eq 0) { 'Nothing chosen yet.' }
            elseif (@($queryState.Selected).Count -le 3) { @($queryState.Selected) -join ', ' }
            else { '{0} chosen: {1}...' -f @($queryState.Selected).Count, ((@($queryState.Selected)[0..2]) -join ', ') })
    }.GetNewClosure()

    if ([string]$working.refreshMethod -eq 'SelectedQueries') { $rdoSome.Checked = $true } else { $rdoAll.Checked = $true }
    $rdoSome.Add_CheckedChanged($sync)
    & $sync

    $btnChoose.Add_Click({
        $picked = Show-QuerySelectionDialog -WorkbookPath $path -SelectedQueries @($queryState.Selected) `
            -AllowWorkbookMacros:([bool](ConvertTo-BoolValue $working.allowWorkbookMacros $false)) -Owner $form
        if ($null -ne $picked) { $queryState.Selected = @($picked); & $sync }
    }.GetNewClosure())

    $btnSave = New-FormAutoButton -Text 'Save' -MinimumWidth 100
    $form.Controls.Add($btnSave)
    $btnSave.Location = New-Object System.Drawing.Point(330, 274)

    $btnCancel = New-FormAutoButton -Text 'Cancel' -MinimumWidth 100
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $btnCancel.Location = New-Object System.Drawing.Point(438, 274)
    $form.CancelButton = $btnCancel

    $btnSave.Add_Click({
        if ($rdoSome.Checked -and @($queryState.Selected).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Choose at least one query, or switch back to refreshing everything.',
                'Which queries', 'OK', 'Information') | Out-Null
            return
        }
        $working['refreshMethod']   = $(if ($rdoSome.Checked) { 'SelectedQueries' } else { 'RefreshAll' })
        $working['selectedQueries'] = $(if ($rdoSome.Checked) { @($queryState.Selected) } else { @() })
        $form.Tag          = $working
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }.GetNewClosure())

    Set-FormWithinWorkingArea -Form $form
    if ($form.ShowDialog($script:UiControls.Form) -eq [System.Windows.Forms.DialogResult]::OK) { return $form.Tag }
    return $null
}

function Invoke-RuleQueryScope {
    <#
        Opens the scope dialog for the selected rule. A rule with several
        workbooks asks which one first, because "which queries" is a question
        about a workbook, not about a rule.
    #>
    $rule = Get-SelectedRule
    if ($null -eq $rule) { return }

    $actions = @($rule.actions)
    if ($actions.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('This rule has no workbooks yet.', 'Which queries', 'OK', 'Information') | Out-Null
        return
    }

    $index = 0
    if ($actions.Count -gt 1) {
        $index = Show-WorkbookChooser -Rule $rule
        if ($index -lt 0) { return }
    }

    $updated = Show-QueryScopeDialog -Action $actions[$index] -AppSettings $script:UiConfig.appSettings
    if ($null -eq $updated) { return }

    $actions[$index] = $updated
    $rule.actions = @($actions)
    if (Save-UiConfiguration) {
        Write-AppLog -Level 'INFO' -RuleName ([string]$rule.name) -Workbook ([string]$updated.path) `
            -Message ('Query selection changed to: {0}' -f $(if ([string]$updated.refreshMethod -eq 'SelectedQueries') { '{0} chosen' -f @($updated.selectedQueries).Count } else { 'everything' }))
        Update-RuleList
    }
}

function Show-WorkbookChooser {
    <#  Returns the index of the chosen workbook, or -1.  #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    $actions = @($Rule.actions)

    $form = New-Object System.Windows.Forms.Form
    $form.Text                = 'Which workbook?'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize          = New-Object System.Drawing.Size(520, 300)
    $form.FormBorderStyle     = 'FixedDialog'
    $form.StartPosition       = 'CenterParent'
    $form.MaximizeBox         = $false
    $form.MinimizeBox         = $false
    $form.Font                = (Get-UiFont)
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(18, 16)
    $label.Size     = New-Object System.Drawing.Size(480, 20)
    $label.Text     = 'This rule refreshes more than one workbook. Which one?'
    $form.Controls.Add($label)

    $list = New-Object System.Windows.Forms.ListView
    $list.Location      = New-Object System.Drawing.Point(18, 44)
    $list.Size          = New-Object System.Drawing.Size(484, 190)
    $list.View          = 'Details'
    $list.FullRowSelect = $true
    $list.MultiSelect   = $false
    $list.HideSelection = $false
    [void]$list.Columns.Add('Workbook', 330)
    [void]$list.Columns.Add('Queries', 130)
    $form.Controls.Add($list)

    foreach ($action in $actions) {
        $item = New-Object System.Windows.Forms.ListViewItem((Split-Path -Leaf ([string]$action.path)))
        [void]$item.SubItems.Add($(if ([string]$action.refreshMethod -eq 'SelectedQueries') { '{0} chosen' -f @($action.selectedQueries).Count } else { 'All' }))
        $item.ToolTipText = [string]$action.path
        [void]$list.Items.Add($item)
    }
    if ($list.Items.Count -gt 0) { $list.Items[0].Selected = $true }

    $btnOk = New-FormAutoButton -Text 'Continue' -MinimumWidth 110
    $form.Controls.Add($btnOk)
    $btnOk.Location = New-Object System.Drawing.Point(280, 252)
    $btnOk.Add_Click({ $form.Tag = $list.SelectedIndices[0]; $form.DialogResult = 'OK'; $form.Close() }.GetNewClosure())

    $btnCancel = New-FormAutoButton -Text 'Cancel' -MinimumWidth 100
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $btnCancel.Location = New-Object System.Drawing.Point(400, 252)
    $form.CancelButton = $btnCancel

    $list.Add_DoubleClick({ $btnOk.PerformClick() }.GetNewClosure())

    $result = -1
    if ($form.ShowDialog($script:UiControls.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $result = [int]$form.Tag }
    $form.Dispose()
    return $result
}
