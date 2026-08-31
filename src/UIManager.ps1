# ==============================================================================
#  UIManager.ps1  (UI runspace only)
#
#  The dashboard never performs work. It reads the shared state on a WinForms
#  timer and posts commands back to the engine. Because every control is only
#  ever touched from the timer tick (which runs on the UI thread) there is no
#  cross-thread control access anywhere in this file.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:UiShared        = $null
$script:UiPaths         = $null
$script:UiConfig        = $null
$script:UiIcons         = @{}
$script:UiActivitySeq   = 0
$script:UiConfigVersion = -1
$script:UiLastStatus    = ''
$script:UiLastIconKey   = ''
$script:UiExiting       = $false
$script:UiControls      = @{}
$script:UiLastRuleRefresh = [DateTime]::MinValue
$script:UiLogonPromptDone = $true
$script:UiLastManualWorkbook = ''
$script:UiFontSize        = 9
$script:UiFontName        = ''
$script:UiFontPreference  = ''
$script:UiFontFamilies    = $null
$script:UiFonts           = @{}
$script:UiSeenManualHistory = @{}
$script:UiManualHistoryInitialized = $false
$script:UiSeenTriggeredHistory = @{}
$script:UiTriggeredHistoryInitialized = $false
$script:UiUpdatingAskBeforeCheck = $false
$script:UiCancelRequestedAt   = $null
$script:UiCancelDialogsClosed = $false
$script:UiCancelPromptShown   = $false
$script:UiCancelPromptOpen    = $false
$script:UiDialogFirstSeenAt   = $null
$script:UiDialogLastTitle     = ''
$script:UiDialogClosedCount   = 0
$script:UiBottomTips          = $null
$script:UiLastWorkbookScan    = [DateTime]::MinValue
$script:UiActivityBuffer      = New-Object System.Collections.ArrayList
$script:UiActivityRuleChoices = ''
$script:UiPendingJobIds       = @()
$script:UiPendingSignature    = ''
$script:UiRunNowIsQueue       = $null
$script:UiHighlightSignature  = ''
$script:UiRefreshApprovalCheckTask = $null
$script:RuleDragDataFormat    = 'ExcelQueryTrigger.RuleId'
$script:UiRuleDragActive      = $false
$script:UiRuleListRefreshPending = $false
$script:UiRuleSortColumn      = -1
$script:UiRuleSortDirection   = 'Ascending'
$script:UiDisplayedRuleIds    = @()
$script:UiRuleColumnNames     = @('Name', 'Trigger', 'Workbook', 'Where', 'Status', 'Queries',
    'Last Run', 'Next Run', 'Duration', 'Data updated')

function Get-StartupMonitoringLines {
    <#  Human-readable startup outcomes for the foreground progress dialog. #>
    param([hashtable]$Shared)

    $lines = New-Object System.Collections.ArrayList
    foreach ($rule in @($script:UiConfig.rules)) {
        $name = [string]$rule.name
        if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Unnamed rule' }
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) {
            [void]$lines.Add(('[--] {0}: disabled' -f $name))
            continue
        }

        $status = ''
        if ($Shared.RuleState.ContainsKey([string]$rule.id)) {
            $status = [string]$Shared.RuleState[[string]$rule.id].WatcherStatus
        }
        $target = [string]$rule.trigger.path
        switch ($status) {
            'Active'          { [void]$lines.Add(('[OK] {0}: monitoring active' -f $name)) }
            'PathUnavailable' { [void]$lines.Add(('[!!] {0}: path unavailable - {1}' -f $name, $target)) }
            'Misconfigured'   { [void]$lines.Add(('[!!] {0}: monitoring path is not configured' -f $name)) }
            'Error'           { [void]$lines.Add(('[!!] {0}: monitor could not be started' -f $name)) }
            'Waiting for time' { [void]$lines.Add(('[OK] {0}: schedule ready' -f $name)) }
            'Waiting for logon' { [void]$lines.Add(('[OK] {0}: logon rule ready' -f $name)) }
            'Manual only'     { [void]$lines.Add(('[OK] {0}: manual rule ready' -f $name)) }
            default           { [void]$lines.Add(('[..] {0}: {1}' -f $name, $(if ([string]::IsNullOrWhiteSpace($status)) { 'preparing' } else { $status }))) }
        }
    }
    if ($lines.Count -eq 0) { [void]$lines.Add('[OK] No trigger rules are configured') }
    return @($lines.ToArray())
}

function New-FormButtonRow {
    <#  A left-to-right strip that lays auto-sized buttons out for us.  #>
    param([int]$X, [int]$Y, [int]$Width, [int]$Height = 36)

    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Location      = New-Object System.Drawing.Point($X, $Y)
    $panel.Size          = New-Object System.Drawing.Size($Width, $Height)
    $panel.FlowDirection = 'LeftToRight'
    $panel.WrapContents  = $false
    $panel.AutoSize      = $false
    $panel.Padding       = New-Object System.Windows.Forms.Padding(0)
    $panel.Margin        = New-Object System.Windows.Forms.Padding(0)
    return $panel
}

function Set-FormWithinWorkingArea {
    <#
        Keep every form inside the current Windows working area. The application
        uses Windows compatibility scaling so the complete fixed-coordinate UI is
        scaled as one unit on high-DPI/Retina displays. This function remains a
        second guard for small screens, taskbars, remote sessions and unusual work
        areas: clamp only when needed and enable scrolling instead of hiding controls.
    #>
    param([System.Windows.Forms.Form]$Form)

    try {
        # Last-resort protection for compact captions throughout every dialog.
        # Paragraphs use New-FormWrappedLabel and grow to show all text; labels
        # that are intentionally one line show an ellipsis instead of losing
        # half a word when a substituted font is wider than expected.
        $pending = New-Object System.Collections.Stack
        $pending.Push($Form)
        while ($pending.Count -gt 0) {
            $parent = $pending.Pop()
            foreach ($control in @($parent.Controls)) {
                $pending.Push($control)
                if (($control -is [System.Windows.Forms.Label]) -and (-not $control.AutoSize) -and ($control.Height -le 30)) {
                    $control.AutoEllipsis = $true
                }
                elseif (($control -is [System.Windows.Forms.ButtonBase]) -and (-not $control.AutoSize)) {
                    $control.AutoEllipsis = $true
                }
            }
        }

        $work      = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $maxWidth  = $work.Width  - 16
        $maxHeight = $work.Height - 16
        if ($Form.Width -le $maxWidth -and $Form.Height -le $maxHeight) { return }

        $Form.AutoScroll = $true
        $Form.Size = New-Object System.Drawing.Size(
            [Math]::Min($Form.Width, $maxWidth),
            [Math]::Min($Form.Height, $maxHeight))
    }
    catch { }
}

function New-FormAutoButton {
    <#
        A button that is exactly as wide as its caption needs, whatever font or
        display scale is in use.
    #>
    param([string]$Text, [int]$MinimumWidth = 84)

    $button = New-Object System.Windows.Forms.Button
    $button.Text         = $Text
    $button.AutoSize     = $true
    $button.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $button.MinimumSize  = New-Object System.Drawing.Size($MinimumWidth, 30)
    $button.Padding      = New-Object System.Windows.Forms.Padding(10, 3, 10, 3)
    $button.Margin       = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    return $button
}

function Get-InstalledFontFamilies {
    <#
        Asking GDI+ for a font that is not installed silently substitutes
        another one, so the only reliable test is to look at what is actually
        installed. Enumerated once and cached.
    #>
    if ($null -eq $script:UiFontFamilies) {
        $script:UiFontFamilies = @{}
        try {
            $collection = New-Object System.Drawing.Text.InstalledFontCollection
            foreach ($family in $collection.Families) { $script:UiFontFamilies[$family.Name] = $true }
            $collection.Dispose()
        }
        catch { }
    }
    return $script:UiFontFamilies
}

function Get-UiFontName {
    <#
        Noto Sans JP first: it covers Latin and Japanese in one family, so the
        two do not end up drawn by different fonts on the same line. Everything
        after it is a fallback for machines where it is not installed.
    #>
    if (-not [string]::IsNullOrWhiteSpace($script:UiFontName)) { return $script:UiFontName }

    $installed  = Get-InstalledFontFamilies
    $candidates = @('Noto Sans JP', 'Noto Sans CJK JP', 'Meiryo UI', 'Yu Gothic UI', 'Segoe UI', 'Microsoft Sans Serif')
    if (-not [string]::IsNullOrWhiteSpace($script:UiFontPreference)) {
        $candidates = @($script:UiFontPreference) + $candidates
    }
    foreach ($name in $candidates) {
        if ($installed.ContainsKey($name)) {
            $script:UiFontName = $name
            return $name
        }
    }
    $script:UiFontName = 'Microsoft Sans Serif'
    return $script:UiFontName
}

function Get-UiFont {
    <#
        Size 0 (the default) means "the size the user chose". Passing a number
        is a relative request: 11 asks for two points larger than the base.
    #>
    param([single]$Size = 0, [string]$Style = 'Regular')

    $base = [single]$script:UiFontSize
    if ($base -lt 7) { $base = 9 }
    if ($Size -le 0) { $Size = $base }
    else             { $Size = $base + ($Size - 9) }

    $fontStyle = [System.Drawing.FontStyle]::Regular
    if ($Style -eq 'Bold') { $fontStyle = [System.Drawing.FontStyle]::Bold }

    try   { return (New-Object System.Drawing.Font((Get-UiFontName), $Size, $fontStyle)) }
    catch { return (New-Object System.Drawing.Font('Microsoft Sans Serif', $Size, $fontStyle)) }
}

function New-StatusIcon {
    <#
        Builds the tray icons at runtime so the application needs no .ico assets
        Created once and reused; the handles live as long as
        the process.
    #>
    param([System.Drawing.Color]$Color)

    $bitmap = New-Object System.Drawing.Bitmap 16, 16
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $brush = New-Object System.Drawing.SolidBrush($Color)
        $pen   = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 60, 60), 1)
        $graphics.FillEllipse($brush, 1, 1, 13, 13)
        $graphics.DrawEllipse($pen, 1, 1, 13, 13)
        $brush.Dispose()
        $pen.Dispose()
    }
    finally {
        $graphics.Dispose()
    }

    $handle = $bitmap.GetHicon()
    return [System.Drawing.Icon]::FromHandle($handle)
}

function Initialize-UiIcons {
    $script:UiIcons = @{
        Running    = (New-StatusIcon ([System.Drawing.Color]::FromArgb(34, 150, 60)))
        Refreshing = (New-StatusIcon ([System.Drawing.Color]::FromArgb(220, 170, 20)))
        Error      = (New-StatusIcon ([System.Drawing.Color]::FromArgb(200, 50, 45)))
        Paused     = (New-StatusIcon ([System.Drawing.Color]::FromArgb(130, 130, 130)))
    }
}

function Get-StatusIconKey {
    param([string]$Status)
    switch ($Status) {
        'Refreshing' { return 'Refreshing' }
        'Paused'     { return 'Paused' }
        'Stopping'   { return 'Paused' }
        'Stopped'    { return 'Paused' }
        'Degraded'   { return 'Error' }
        'Error'      { return 'Error' }
        default      { return 'Running' }
    }
}

function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        'Refreshing' { return [System.Drawing.Color]::FromArgb(176, 132, 10) }
        'Paused'     { return [System.Drawing.Color]::FromArgb(90, 90, 90) }
        'Degraded'   { return [System.Drawing.Color]::FromArgb(190, 90, 20) }
        'Error'      { return [System.Drawing.Color]::FromArgb(180, 40, 35) }
        'Stopped'    { return [System.Drawing.Color]::FromArgb(90, 90, 90) }
        default      { return [System.Drawing.Color]::FromArgb(30, 120, 55) }
    }
}

# ------------------------------------------------------------------------------
# Region: dashboard construction
# ------------------------------------------------------------------------------

function Show-Dashboard {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Shared,
        [Parameter(Mandatory = $true)][hashtable]$Paths,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [switch]$StartedFromLogon,
        [AllowNull()]$Splash = $null
    )

    $script:UiShared = $Shared
    $script:UiPaths  = $Paths
    $script:UiConfig = $Config

    Initialize-UiIcons

    $script:UiFontSize       = ConvertTo-IntValue $script:UiConfig.appSettings.uiFontSize 9 8
    $script:UiFontPreference = [string]$script:UiConfig.appSettings.uiFontName
    $script:UiFontName       = ''
    $font     = Get-UiFont
    $fontBold = Get-UiFont -Style 'Bold'
    $script:UiFonts = @{ Regular = $font; Bold = $fontBold }

    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'Excel Query Trigger Manager'
    # Every coordinate in this dialog was written for a 96 DPI screen. Font
    # scaling (the WinForms default) would re-scale them by whatever runtime UI
    # font happens to be available and make labels overlap the fields next to
    # them; DPI scaling multiplies them by the real display scale instead, which
    # is exactly what is wanted. On a 100% display the factor is 1 and nothing
    # moves, so this is safe either way.
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::None
    # 972 wide, not 884: Duration earns its column back and the ten columns
    # need 904 px. The height is unchanged - the space came out of the header
    # and the rule list, which were both showing more emptiness than content.
    $form.ClientSize    = New-Object System.Drawing.Size(972, 730)
    $form.MinimumSize   = New-Object System.Drawing.Size(988, 660)
    $form.StartPosition = 'CenterScreen'
    $form.Font          = $font
    # The shipped icon when it is there, the drawn one otherwise, so the window,
    # the taskbar and the Start menu shortcut all look like the same program.
    $form.Icon = $script:UiIcons.Running
    $shippedIcon = Join-Path $script:UiPaths.AppRoot 'assets\ExcelQueryTrigger.ico'
    if (Test-Path -LiteralPath $shippedIcon) {
        try { $form.Icon = New-Object System.Drawing.Icon($shippedIcon) } catch { }
    }
    $script:UiControls.Form = $form

    # ---- status header ----------------------------------------------------
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(14, 11)
    # AutoSize avoids clipping descenders such as the final 'g' in 'Running'
    # when a different installed font or DPI produces a slightly taller glyph.
    $lblStatus.AutoSize = $true
    $lblStatus.MaximumSize = New-Object System.Drawing.Size(600, 0)
    $lblStatus.Font     = (Get-UiFont 11 'Bold')
    $lblStatus.Text     = 'Status: Starting'
    $form.Controls.Add($lblStatus)
    $script:UiControls.Status = $lblStatus

    # Global one-off refresh action. This is intentionally outside Trigger Rules:
    # it can refresh any Excel file, even one that is not part of a rule.
    $btnRunFile = New-Object System.Windows.Forms.Button
    $btnRunFile.Size     = New-Object System.Drawing.Size(190, 30)
    $btnRunFile.Location = New-Object System.Drawing.Point(728, 8)
    $btnRunFile.Anchor   = 'Top,Right'
    $btnRunFile.Text     = 'Refresh Any Excel File...'
    $btnRunFile.TextAlign = 'MiddleCenter'
    $form.Controls.Add($btnRunFile)
    $script:UiControls.BtnRunFile = $btnRunFile

    # Compact information button. Use a plain text glyph rather than a bitmap
    # icon: Windows can scale SystemIcons inconsistently at different DPI/font
    # settings, which made the previous icon look crushed on some PCs. Keep the
    # internal control key stable so future caption changes cannot break startup.
    $btnRefreshRules = New-Object System.Windows.Forms.Button
    $btnRefreshRules.Size     = New-Object System.Drawing.Size(34, 30)
    # 926 + 34 = 960, which is the right edge of the group boxes below. These
    # were still sitting at the old 884 px positions, which left a gap.
    $btnRefreshRules.Location = New-Object System.Drawing.Point(926, 8)
    $btnRefreshRules.Anchor   = 'Top,Right'
    $btnRefreshRules.Text     = 'i'
    $btnRefreshRules.Font     = (Get-UiFont 12 'Bold')
    $btnRefreshRules.TextAlign = 'MiddleCenter'
    # WinForms draws a lowercase i slightly below the visual centre. A tiny
    # bottom padding correction keeps the glyph optically centred without
    # changing the control size or its internal key.
    $btnRefreshRules.Padding  = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
    $btnRefreshRules.TabStop  = $true
    $form.Controls.Add($btnRefreshRules)
    $script:UiControls.BtnRefreshRules = $btnRefreshRules
    $tipInfo = New-Object System.Windows.Forms.ToolTip
    $tipInfo.SetToolTip($btnRefreshRules, 'Information & Help: monitoring, triggers, refresh behavior, terms, and Q&A')
    $script:UiControls.InfoToolTip = $tipInfo

    $lblStatusDetail = New-Object System.Windows.Forms.Label
    $lblStatusDetail.Location  = New-Object System.Drawing.Point(16, 33)
    $lblStatusDetail.Size      = New-Object System.Drawing.Size(942, 18)
    $lblStatusDetail.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblStatusDetail.Anchor    = 'Top,Left,Right'
    $form.Controls.Add($lblStatusDetail)
    $script:UiControls.StatusDetail = $lblStatusDetail
    $statusTip = New-Object System.Windows.Forms.ToolTip
    $statusTip.AutoPopDelay = 15000
    $script:UiControls.StatusToolTip = $statusTip

    # ---- rules ------------------------------------------------------------
    $gbRules = New-Object System.Windows.Forms.GroupBox
    $gbRules.Text     = 'Trigger Rules'
    $gbRules.Location = New-Object System.Drawing.Point(12, 52)
    $gbRules.Size     = New-Object System.Drawing.Size(948, 194)
    $gbRules.Anchor   = 'Top,Left,Right'
    $form.Controls.Add($gbRules)
    $script:UiControls.RulesGroup = $gbRules

    # A compact title-bar action replaces the large draggable divider. The
    # expanded height is calculated from the number of rules and available room.
    $lnkRulesView = New-Object System.Windows.Forms.LinkLabel
    $lnkRulesView.Text       = 'Show all'
    $lnkRulesView.TextAlign  = 'MiddleRight'
    $lnkRulesView.Location   = New-Object System.Drawing.Point(844, 1)
    $lnkRulesView.Size       = New-Object System.Drawing.Size(92, 22)
    $lnkRulesView.Anchor     = 'Top,Right'
    $lnkRulesView.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
    $lnkRulesView.BackColor  = [System.Drawing.SystemColors]::Control
    $gbRules.Controls.Add($lnkRulesView)
    $lnkRulesView.BringToFront()
    $script:UiControls.RulesViewToggle = $lnkRulesView

    $lvRules = New-Object System.Windows.Forms.ListView
    $lvRules.Location      = New-Object System.Drawing.Point(10, 27)
    # Leave a dedicated action strip inside the Trigger Rules group so the
    # Add/Edit/Delete/Run controls read visually as operations on these rules.
    $lvRules.Size          = New-Object System.Drawing.Size(928, 116)
    $lvRules.Anchor        = 'Top,Left,Right,Bottom'
    $lvRules.View          = 'Details'
    $lvRules.FullRowSelect = $true
    $lvRules.GridLines     = $true
    $lvRules.MultiSelect   = $false
    $lvRules.HideSelection = $false
    $lvRules.AllowDrop     = $true
    $lvRules.InsertionMark.Color = [System.Drawing.Color]::FromArgb(35, 100, 180)
    $lvRules.AccessibleName = 'Trigger rules'
    $lvRules.AccessibleDescription = 'Rules, their trigger, workbook, status, schedule and last refresh information. Click a column heading to sort, or drag a row to change its saved order.'
    # Native ListView item tooltips are inconsistent for subitems on some
    # Windows/DPI combinations. A dedicated hover tooltip is registered below.
    $lvRules.ShowItemToolTips = $false
    # Ten columns, 904 px, inside a 928 px list. Where and Data updated were
    # added without taking Duration's place - the window widened instead.
    [void]$lvRules.Columns.Add('Name', 126)
    [void]$lvRules.Columns.Add('Trigger', 72)
    [void]$lvRules.Columns.Add('Workbook', 124)
    [void]$lvRules.Columns.Add('Where', 74)
    [void]$lvRules.Columns.Add('Status', 100)
    [void]$lvRules.Columns.Add('Queries', 62)
    [void]$lvRules.Columns.Add('Last Run', 90)
    [void]$lvRules.Columns.Add('Next Run', 88)
    [void]$lvRules.Columns.Add('Duration', 72)
    [void]$lvRules.Columns.Add('Data updated', 96)
    $gbRules.Controls.Add($lvRules)
    $script:UiControls.Rules = $lvRules

    $rulesHoverTip = New-Object System.Windows.Forms.ToolTip
    $rulesHoverTip.InitialDelay = 250
    $rulesHoverTip.ReshowDelay  = 100
    $rulesHoverTip.AutoPopDelay = 20000
    $rulesHoverTip.ShowAlways   = $true
    $script:UiControls.RulesHoverToolTip = $rulesHoverTip
    $script:UiControls.RulesHoverState = @{ Key = '' }

    # Sits on top of the rule list while there are no rules. A first-time user
    # otherwise sees an empty grid and five buttons with nothing to act on.
    $emptyPanel = New-Object System.Windows.Forms.Panel
    $emptyPanel.Location  = $lvRules.Location
    $emptyPanel.Size      = $lvRules.Size
    $emptyPanel.Anchor    = 'Top,Left,Right,Bottom'
    $emptyPanel.BackColor = [System.Drawing.Color]::White
    $emptyPanel.Visible   = $false
    $gbRules.Controls.Add($emptyPanel)
    $emptyPanel.BringToFront()
    $script:UiControls.RulesEmpty = $emptyPanel

    $lblEmptyTitle = New-Object System.Windows.Forms.Label
    $lblEmptyTitle.Text      = 'No rules yet'
    $lblEmptyTitle.Font      = New-Object System.Drawing.Font($lvRules.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $lblEmptyTitle.TextAlign = 'MiddleCenter'
    $lblEmptyTitle.Location  = New-Object System.Drawing.Point(0, 10)
    $lblEmptyTitle.Size      = New-Object System.Drawing.Size(928, 26)
    $lblEmptyTitle.Anchor    = 'Top,Left,Right'
    $emptyPanel.Controls.Add($lblEmptyTitle)

    $lblEmptyBody = New-Object System.Windows.Forms.Label
    $lblEmptyBody.Text      = 'A rule says when to refresh and which workbooks to refresh.' + [Environment]::NewLine +
                              'For example: every weekday at 07:30, or whenever a new file lands in a folder.'
    $lblEmptyBody.TextAlign = 'MiddleCenter'
    $lblEmptyBody.Location  = New-Object System.Drawing.Point(0, 36)
    $lblEmptyBody.Size      = New-Object System.Drawing.Size(928, 40)
    $lblEmptyBody.Anchor    = 'Top,Left,Right'
    $emptyPanel.Controls.Add($lblEmptyBody)

    $btnEmptyAdd = New-FormAutoButton -Text 'Create your first rule' -MinimumWidth 190
    $emptyPanel.Controls.Add($btnEmptyAdd)
    # Point takes integers; a division would hand it a double and there is no
    # matching constructor.
    $emptyAddX = [int](($emptyPanel.Width - $btnEmptyAdd.Width) / 2)
    $btnEmptyAdd.Location = New-Object System.Drawing.Point($emptyAddX, 78)
    $btnEmptyAdd.Anchor   = 'Top'
    $script:UiControls.BtnEmptyAdd = $btnEmptyAdd

    $btnEmptyHelp = New-Object System.Windows.Forms.LinkLabel
    $btnEmptyHelp.Text      = 'or read how it works first'
    $btnEmptyHelp.TextAlign = 'MiddleCenter'
    $btnEmptyHelp.Location  = New-Object System.Drawing.Point(0, 114)
    $btnEmptyHelp.Size      = New-Object System.Drawing.Size(928, 22)
    $btnEmptyHelp.Anchor    = 'Top,Left,Right'
    $emptyPanel.Controls.Add($btnEmptyHelp)
    $script:UiControls.LinkEmptyHelp = $btnEmptyHelp

    # ---- right-click menu on the rule list ---------------------------------
    # One row now represents the whole rule, so workbook-related commands are
    # submenus populated from that rule's Excel Actions when the menu opens.
    $ruleMenu = New-Object System.Windows.Forms.ContextMenuStrip

    $menuRefreshWorkbook = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuRefreshWorkbook.Text = 'Refresh Workbook Now'
    [void]$ruleMenu.Items.Add($menuRefreshWorkbook)
    $script:UiControls.MenuRefreshWorkbook = $menuRefreshWorkbook

    $menuRunWholeRule = $ruleMenu.Items.Add('Run the Whole Rule Now')
    $script:UiControls.MenuRunWholeRule = $menuRunWholeRule

    $menuQueries = $ruleMenu.Items.Add('Choose which queries to refresh...')
    $script:UiControls.MenuQueries = $menuQueries

    $menuTestDataSources = $ruleMenu.Items.Add('Test data sources...')
    $script:UiControls.MenuTestDataSources = $menuTestDataSources

    # Add now opens the guided three-question flow. The full editor is still
    # one click away for a rule that needs several workbooks or a debounce.
    $menuAddAdvanced = $ruleMenu.Items.Add('New rule with all options...')
    $script:UiControls.MenuAddAdvanced = $menuAddAdvanced
    [void]$ruleMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $menuOpenMonitoredFile = $ruleMenu.Items.Add('Open Monitored File')
    $script:UiControls.MenuOpenMonitoredFile = $menuOpenMonitoredFile
    $menuOpenMonitoredFolder = $ruleMenu.Items.Add('Open Monitored Folder')
    $script:UiControls.MenuOpenMonitoredFolder = $menuOpenMonitoredFolder
    [void]$ruleMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $menuOpenWorkbook = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuOpenWorkbook.Text = 'Open Workbook'
    [void]$ruleMenu.Items.Add($menuOpenWorkbook)
    $script:UiControls.MenuOpenWorkbook = $menuOpenWorkbook

    $menuOpenWorkbookFolder = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuOpenWorkbookFolder.Text = 'Open Workbook Folder'
    [void]$ruleMenu.Items.Add($menuOpenWorkbookFolder)
    $script:UiControls.MenuOpenWorkbookFolder = $menuOpenWorkbookFolder

    [void]$ruleMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $script:UiControls.MenuEditRule = $ruleMenu.Items.Add('Edit Rule')
    $script:UiControls.MenuEnableDisableRule = $ruleMenu.Items.Add('Enable / Disable Rule')
    $script:UiControls.MenuDeleteRule = $ruleMenu.Items.Add('Delete Rule')

    $lvRules.ContextMenuStrip = $ruleMenu
    $script:UiControls.RuleMenu = $ruleMenu

    # ---- rule buttons -----------------------------------------------------
    # Keep rule-specific actions inside the Trigger Rules group. The one-off
    # 'Refresh Any Excel File' command is a global action and lives in the
    # header instead.
    $ruleButtonRow = New-FormButtonRow -X 10 -Y 149 -Width 928 -Height 36
    $ruleButtonRow.Anchor = 'Left,Right,Bottom'
    $gbRules.Controls.Add($ruleButtonRow)
    $script:UiControls.RuleButtonRow = $ruleButtonRow

    # Ordered by how often it is used, with the destructive one last and away
    # from the rest. "Queries..." is here because deciding what gets refreshed
    # is a thing you do to a rule, not a setting buried three levels inside it.
    $buttonSpecs = @(
        @{ Key = 'Add';     Text = 'Add';              Width = 82 },
        @{ Key = 'Edit';    Text = 'Edit';             Width = 82 },
        @{ Key = 'Queries'; Text = 'Queries...';       Width = 108 },
        @{ Key = 'Toggle';  Text = 'Enable / Disable'; Width = 142 },
        # Wide enough for "Add to Queue" as well: the caption changes while a
        # refresh is running, and a resizing button row reads as a glitch.
        @{ Key = 'RunNow';  Text = 'Run Now';          Width = 132 },
        @{ Key = 'Delete';  Text = 'Delete';           Width = 88 }
    )
    foreach ($spec in $buttonSpecs) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text       = [string]$spec.Text
        $button.Size       = New-Object System.Drawing.Size([int]$spec.Width, 32)
        $button.FlatStyle  = [System.Windows.Forms.FlatStyle]::System
        $button.UseVisualStyleBackColor = $true
        $button.Margin     = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
        $button.TextAlign  = 'MiddleCenter'
        [void]$ruleButtonRow.Controls.Add($button)
        $script:UiControls[('Btn' + $spec.Key)] = $button
    }

    $chkAskBefore = New-Object System.Windows.Forms.CheckBox
    $chkAskBefore.Text       = 'Ask before refresh'
    $chkAskBefore.AutoSize   = $true
    $chkAskBefore.Enabled    = $false
    $chkAskBefore.Margin     = New-Object System.Windows.Forms.Padding(14, 7, 0, 0)
    $chkAskBefore.Tag        = 'AskBeforeRefresh'
    [void]$ruleButtonRow.Controls.Add($chkAskBefore)
    $script:UiControls.ChkAskBeforeRefresh = $chkAskBefore

    # ---- current job ------------------------------------------------------
    $gbCurrent = New-Object System.Windows.Forms.GroupBox
    $gbCurrent.Text     = 'Current Job'
    $gbCurrent.Location = New-Object System.Drawing.Point(12, 252)
    # 150 high, not 128: the field block plus a 30px button row did not fit in
    # 128, so the Cancel Job button was drawn over the bottom border.
    $gbCurrent.Size     = New-Object System.Drawing.Size(578, 170)
    $form.Controls.Add($gbCurrent)
    $script:UiControls.CurrentGroup = $gbCurrent

    $currentFields = @(
        @{ Key = 'JobRule';     Caption = 'Rule:' },
        @{ Key = 'JobWorkbook'; Caption = 'Workbook:' },
        @{ Key = 'JobStage';    Caption = 'Stage:' },
        @{ Key = 'JobQuery';    Caption = 'Query:' },
        @{ Key = 'JobElapsed';  Caption = 'Elapsed:' }
    )
    $y = 20
    foreach ($field in $currentFields) {
        $caption = New-Object System.Windows.Forms.Label
        $caption.Text     = $field.Caption
        $caption.Location = New-Object System.Drawing.Point(12, $y)
        $caption.Size     = New-Object System.Drawing.Size(72, 20)
        $gbCurrent.Controls.Add($caption)

        $value = New-Object System.Windows.Forms.Label
        $value.Text      = '-'
        $value.Location  = New-Object System.Drawing.Point(88, $y)
        $value.Size      = New-Object System.Drawing.Size(478, 20)
        $value.Font      = $fontBold
        $value.AutoEllipsis = $true
        $gbCurrent.Controls.Add($value)
        $script:UiControls[$field.Key] = $value

        $y += 20
    }

    $queryProgressToolTip = New-Object System.Windows.Forms.ToolTip
    $queryProgressToolTip.AutoPopDelay = 30000
    $queryProgressToolTip.InitialDelay = 250
    $queryProgressToolTip.ReshowDelay = 100
    $script:UiControls.QueryProgressToolTip = $queryProgressToolTip

    # All Current/Pending job actions share one baseline and one height. This
    # keeps Cancel Job aligned with Move Up / Move Down / Remove at every DPI.
    $jobActionY = 132
    $jobActionHeight = 30

    $progress = New-Object System.Windows.Forms.ProgressBar
    # Keep the indefinite bar aligned with the action buttons below the fields.
    $progress.Location = New-Object System.Drawing.Point(12, 139)
    $progress.Size     = New-Object System.Drawing.Size(400, 16)
    $progress.Style    = 'Marquee'
    $progress.MarqueeAnimationSpeed = 28
    $progress.Visible  = $false
    $gbCurrent.Controls.Add($progress)
    $script:UiControls.Progress = $progress

    # A refresh has no upper time limit - the configured time is only a warning
    # threshold - so there has to be a way out of one that never finishes.
    $btnCancelJob = New-FormAutoButton -Text 'Cancel Job' -MinimumWidth 90
    $btnCancelJob.AutoSize = $false
    $btnCancelJob.Height   = $jobActionHeight
    $btnCancelJob.Enabled  = $false
    $btnCancelJob.AccessibleDescription = 'Cancels only while Excel is before the save step.'
    $gbCurrent.Controls.Add($btnCancelJob)

    # The button auto-sizes to its caption, so anchor it to the right edge from
    # its measured width and give the bar whatever is left of the row. A fixed
    # X had it overhanging the group border once the button grew.
    $cancelWidth = [Math]::Max($btnCancelJob.Width, $btnCancelJob.PreferredSize.Width)
    $btnCancelJob.Width = $cancelWidth
    $btnCancelJob.Location = New-Object System.Drawing.Point(($gbCurrent.Width - 15 - $cancelWidth), $jobActionY)
    $progress.Width = $btnCancelJob.Left - 10 - $progress.Left
    $script:UiControls.BtnCancelJob = $btnCancelJob

    # ---- pending queue ----------------------------------------------------
    $gbPending = New-Object System.Windows.Forms.GroupBox
    $gbPending.Text     = 'Pending Jobs'
    $gbPending.Location = New-Object System.Drawing.Point(602, 252)
    $gbPending.Size     = New-Object System.Drawing.Size(358, 170)
    $gbPending.Anchor   = 'Top,Left,Right'
    $form.Controls.Add($gbPending)
    $script:UiControls.PendingGroup = $gbPending

    $lstPending = New-Object System.Windows.Forms.ListBox
    $lstPending.Location = New-Object System.Drawing.Point(10, 20)
    $lstPending.Size     = New-Object System.Drawing.Size(338, 108)
    $lstPending.Anchor   = 'Top,Left,Right,Bottom'
    $lstPending.IntegralHeight = $false
    $lstPending.AccessibleName = 'Pending jobs'
    $lstPending.AccessibleDescription = 'Refresh jobs waiting to run. Use the adjacent buttons to reorder or remove them.'
    $gbPending.Controls.Add($lstPending)
    $script:UiControls.Pending = $lstPending

    # Laid out left to right from each button's measured width, so a wider font
    # cannot push the last one through the group border.
    $pendingX = 10
    foreach ($spec in @(
        @{ Key = 'PendingUp';     Text = 'Move Up' },
        @{ Key = 'PendingDown';   Text = 'Move Down' },
        @{ Key = 'PendingRemove'; Text = 'Remove' })) {
        $button = New-FormAutoButton -Text ([string]$spec.Text) -MinimumWidth 84
        $button.AutoSize = $false
        $button.Height = $jobActionHeight
        $button.Width = [Math]::Max($button.Width, $button.PreferredSize.Width)
        $button.Enabled = $false
        $gbPending.Controls.Add($button)
        $button.Location = New-Object System.Drawing.Point($pendingX, $jobActionY)
        $pendingX = $button.Right + 8
        $script:UiControls[('Btn' + $spec.Key)] = $button
    }

    # ---- activity ---------------------------------------------------------
    $gbActivity = New-Object System.Windows.Forms.GroupBox
    $gbActivity.Text     = 'Recent Activity  (select a row to see the details)'
    $gbActivity.Location = New-Object System.Drawing.Point(12, 430)
    $gbActivity.Size     = New-Object System.Drawing.Size(948, 256)
    $gbActivity.Anchor   = 'Top,Left,Right,Bottom'
    $form.Controls.Add($gbActivity)
    $script:UiControls.ActivityGroup = $gbActivity

    $lblShow = New-Object System.Windows.Forms.Label
    $lblShow.Text     = 'Show'
    $lblShow.Location = New-Object System.Drawing.Point(10, 22)
    $lblShow.AutoSize = $true
    $gbActivity.Controls.Add($lblShow)

    $cmbActivityLevel = New-Object System.Windows.Forms.ComboBox
    $cmbActivityLevel.AccessibleName = 'Activity level filter'
    $cmbActivityLevel.Location      = New-Object System.Drawing.Point(62, 18)
    $cmbActivityLevel.Size          = New-Object System.Drawing.Size(158, 22)
    $cmbActivityLevel.DropDownStyle = 'DropDownList'
    [void]$cmbActivityLevel.Items.AddRange(@('Everything', 'Hide routine detail', 'Problems only', 'Errors only'))
    $cmbActivityLevel.SelectedIndex = 0
    $gbActivity.Controls.Add($cmbActivityLevel)
    $script:UiControls.ActivityLevel = $cmbActivityLevel

    $lblForRule = New-Object System.Windows.Forms.Label
    $lblForRule.Text     = 'for rule'
    $lblForRule.Location = New-Object System.Drawing.Point(232, 22)
    $lblForRule.AutoSize = $true
    $gbActivity.Controls.Add($lblForRule)

    $cmbActivityRule = New-Object System.Windows.Forms.ComboBox
    $cmbActivityRule.AccessibleName = 'Activity rule filter'
    $cmbActivityRule.Location      = New-Object System.Drawing.Point(300, 18)
    $cmbActivityRule.Size          = New-Object System.Drawing.Size(196, 22)
    $cmbActivityRule.DropDownStyle = 'DropDownList'
    [void]$cmbActivityRule.Items.Add('Any rule')
    $cmbActivityRule.SelectedIndex = 0
    $gbActivity.Controls.Add($cmbActivityRule)
    $script:UiControls.ActivityRule = $cmbActivityRule

    $lblContaining = New-Object System.Windows.Forms.Label
    $lblContaining.Text     = 'containing'
    $lblContaining.Location = New-Object System.Drawing.Point(574, 22)
    $lblContaining.AutoSize = $true
    $lblContaining.Anchor   = 'Top,Right'
    $gbActivity.Controls.Add($lblContaining)

    $txtActivityFind = New-Object System.Windows.Forms.TextBox
    $txtActivityFind.AccessibleName = 'Search recent activity'
    $txtActivityFind.Location = New-Object System.Drawing.Point(662, 18)
    $txtActivityFind.Size     = New-Object System.Drawing.Size(186, 22)
    $txtActivityFind.Anchor   = 'Top,Right'
    $gbActivity.Controls.Add($txtActivityFind)
    $script:UiControls.ActivityFind = $txtActivityFind

    $btnActivityClear = New-Object System.Windows.Forms.Button
    $btnActivityClear.Text      = 'Reset'
    $btnActivityClear.Location  = New-Object System.Drawing.Point(858, 17)
    $btnActivityClear.Size      = New-Object System.Drawing.Size(80, 24)
    $btnActivityClear.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $btnActivityClear.Anchor    = 'Top,Right'
    $gbActivity.Controls.Add($btnActivityClear)
    $script:UiControls.ActivityClear = $btnActivityClear

    $lvActivity = New-Object System.Windows.Forms.ListView
    $lvActivity.AccessibleName = 'Recent activity'
    $lvActivity.AccessibleDescription = 'Recent trigger, refresh, warning and error entries.'
    $lvActivity.Location      = New-Object System.Drawing.Point(10, 48)
    $lvActivity.Size          = New-Object System.Drawing.Size(928, 140)
    $lvActivity.Anchor        = 'Top,Left,Right,Bottom'
    $lvActivity.View          = 'Details'
    $lvActivity.FullRowSelect = $true
    $lvActivity.MultiSelect   = $false
    $lvActivity.HideSelection = $false
    [void]$lvActivity.Columns.Add('Time', 70)
    [void]$lvActivity.Columns.Add('Level', 70)
    [void]$lvActivity.Columns.Add('Rule', 140)
    [void]$lvActivity.Columns.Add('Message', 550)
    $gbActivity.Controls.Add($lvActivity)
    $script:UiControls.Activity = $lvActivity

    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.AccessibleName = 'Activity details'
    $txtDetail.Location   = New-Object System.Drawing.Point(10, 194)
    $txtDetail.Size       = New-Object System.Drawing.Size(928, 54)
    $txtDetail.Anchor     = 'Left,Right,Bottom'
    $txtDetail.Multiline  = $true
    $txtDetail.ReadOnly   = $true
    $txtDetail.ScrollBars = 'Vertical'
    $txtDetail.BackColor  = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $gbActivity.Controls.Add($txtDetail)
    $script:UiControls.Detail = $txtDetail

    $savedRulesMode = [string]$script:UiConfig.appSettings.rulesPaneMode
    Set-RulesPaneMode -Mode $savedRulesMode

    # ---- bottom status / version ------------------------------------------
    # Keep version information visible on the dashboard, but out of the primary
    # status header. A small muted label at the lower-left follows the visual
    # convention used by many Windows desktop utilities for build/version info.
    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.AutoSize = $true
    $lblVersion.Location = New-Object System.Drawing.Point(16, 699)
    $lblVersion.Anchor = 'Left,Bottom'
    $lblVersion.Text = ('Version {0}' -f (Get-AppVersion))
    $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(115, 115, 115)
    $form.Controls.Add($lblVersion)
    $lblVersion.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:UiControls.Version = $lblVersion

    $versionTip = New-Object System.Windows.Forms.ToolTip
    $versionTip.SetToolTip($lblVersion, 'Which version this is, and whether a newer one is available.')

    # ---- bottom buttons ---------------------------------------------------
    # App-wide utility commands follow the normal Windows-dialog convention:
    # keep them together at the lower-right, separate from rule operations.
    $bottomButtonRow = New-FormButtonRow -X 100 -Y 690 -Width 860
    $bottomButtonRow.FlowDirection = 'RightToLeft'
    $bottomButtonRow.Anchor = 'Right,Bottom'
    $form.Controls.Add($bottomButtonRow)

    # Added right-to-left so the group hugs the right edge. Add the controls in
    # reverse order to keep the visible order Settings ... Exit, with Exit at
    # the far right.
    # "Exit" and "Hide to Tray" look alike but do opposite things, and getting
    # them the wrong way round silently stops every rule. A gap between the two
    # groups, and a caption that says what Exit costs, make that harder to do.
    $bottomSpecs = @(
        @{ Key = 'Exit';     Text = 'Exit (stop refreshing)'; Gap = $false },
        @{ Key = 'Tray';     Text = 'Hide to Tray';           Gap = $true  },
        @{ Key = 'Pause';    Text = 'Pause Monitoring';       Gap = $false },
        @{ Key = 'OpenLog';  Text = 'Open Log';               Gap = $false },
        @{ Key = 'Settings'; Text = 'Settings';               Gap = $false }
    )
    foreach ($spec in $bottomSpecs) {
        $button = New-FormAutoButton -Text $spec.Text
        if ([bool]$spec.Gap) {
            # Each value is worked out first. Writing the arithmetic inline
            # would not do what it looks like: the comma binds tighter than the
            # plus, so "a + 24, b, c" parses as "a + (24, b, c)" - adding an
            # array to a number.
            $current   = $button.Margin
            $gapLeft   = [int]$current.Left + 24
            $gapTop    = [int]$current.Top
            $gapRight  = [int]$current.Right
            $gapBottom = [int]$current.Bottom
            $button.Margin = New-Object System.Windows.Forms.Padding($gapLeft, $gapTop, $gapRight, $gapBottom)
        }
        [void]$bottomButtonRow.Controls.Add($button)
        $script:UiControls[('Btn' + $spec.Key)] = $button
    }

    $bottomTips = New-Object System.Windows.Forms.ToolTip
    $bottomTips.AutoPopDelay = 12000
    $script:UiBottomTips = $bottomTips
    $bottomTips.SetToolTip($script:UiControls.BtnSettings, 'How it starts, what it tells you, and the advanced timing numbers.')
    $bottomTips.SetToolTip($script:UiControls.BtnOpenLog,  'Opens today''s log file. Everything the application did, in plain text.')
    $bottomTips.SetToolTip($script:UiControls.BtnPause,    'Stops rules setting themselves off, but leaves the application running. Nothing is lost - it simply stops watching until you resume.')
    $bottomTips.SetToolTip($script:UiControls.BtnTray,     'Closes the window and keeps the application running by the clock. Rules carry on as normal.')
    $bottomTips.SetToolTip($script:UiControls.BtnExit,     'Closes the application completely. No rule runs again until you start it, and it does not catch up on what it missed.')

    # ---- tray -------------------------------------------------------------
    $trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
    # Keep stable internal keys separate from visible captions. This prevents a
    # harmless wording change in the UI from breaking event-handler registration.
    $traySpecs = @(
        @{ Key = 'OpenDashboard';      Text = 'Open Dashboard' },
        @{ Key = 'RefreshWorkbook';    Text = 'Refresh Any Excel File...' },
        @{ Key = 'PauseMonitoring';    Text = 'Pause Monitoring' },
        @{ Key = 'ResumeMonitoring';   Text = 'Resume Monitoring' },
        @{ Key = 'RunAllManualRules';  Text = 'Run All Manual Rules' },
        @{ Key = 'ViewLog';            Text = 'View Log' },
        @{ Key = 'Settings';           Text = 'Settings' },
        @{ Key = 'Exit';               Text = 'Exit' }
    )
    foreach ($spec in $traySpecs) {
        $item = $trayMenu.Items.Add([string]$spec.Text)
        $script:UiControls[('Tray' + [string]$spec.Key)] = $item
    }

    $tray = New-Object System.Windows.Forms.NotifyIcon
    $tray.Icon             = $script:UiIcons.Running
    $tray.Text             = 'Excel Query Trigger Manager'
    # The splash is the only startup surface. Publish the tray icon only after
    # the first workbook-information scan has finished as well.
    $tray.Visible          = $false
    $tray.ContextMenuStrip = $trayMenu
    $script:UiControls.Tray = $tray

    Register-DashboardHandlers
    Update-RuleList

    # ---- the one timer that drives every visual update --------------------
    # 500 ms by default. This only paints the dashboard; nothing about file
    # monitoring depends on it, so a slower tick costs nothing but a slightly
    # less smooth elapsed counter.
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = ConvertTo-IntValue $script:UiConfig.appSettings.uiRefreshMilliseconds 500 100
    $timer.Add_Tick({ Update-DashboardTick })
    $script:UiControls.Timer = $timer

    # A second launch signals this named event instead of opening another engine.
    # The activation timer is started only after the startup readiness gates have
    # completed, so a second launch can never bypass the splash or expose a
    # Degraded Dashboard. The AutoReset event remains signalled until then.
    try {
        $activateEvent = New-Object System.Threading.EventWaitHandle(
            $false, [System.Threading.EventResetMode]::AutoReset, 'Local\ExcelQueryTriggerActivateExisting')
        $activateAckEvent = New-Object System.Threading.EventWaitHandle(
            $false, [System.Threading.EventResetMode]::AutoReset, 'Local\ExcelQueryTriggerActivateAcknowledged')
        $activateTimer = New-Object System.Windows.Forms.Timer
        $activateTimer.Interval = 200
        $activateTimer.Add_Tick({
            try {
                if ($script:UiControls.ActivateEvent.WaitOne(0, $false)) {
                    Show-DashboardWindow
                    [void]$script:UiControls.ActivateAckEvent.Set()
                }
            }
            catch {
                try { [void]$script:UiControls.ActivateAckEvent.Set() } catch { }
            }
        })
        $script:UiControls.ActivateEvent = $activateEvent
        $script:UiControls.ActivateAckEvent = $activateAckEvent
        $script:UiControls.ActivateTimer = $activateTimer
    }
    catch {
        Write-AppLog -Level 'WARN' -Message ('Single-instance Dashboard activation could not be initialized: {0}' -f $_.Exception.Message)
    }

    $startMinimized = ConvertTo-BoolValue $script:UiConfig.appSettings.startMinimized $true
    $minimizeToTray = ConvertTo-BoolValue $script:UiConfig.appSettings.minimizeToTray $true
    $firstRunWelcomeDue = Test-FirstRunWelcomeDue

    # The engine is on its own runspace. Show the real watcher-registration
    # progress. The Dashboard stays hidden until both the engine and the first
    # workbook metadata scan are complete, so its first visible frame is ready
    # for interaction and already contains Data updated information.
    # Startup readiness is enforced on every launch, including Windows logon.
    # When there is no splash the same gates run silently before the tray icon
    # or Dashboard is published.
    if ($null -ne $Shared) {
        # The splash is the only startup surface. Step 3 waits specifically for
        # the monitoring gate, not for the final startup gate; the engine keeps
        # validating those watchers while Step 4 loads workbook metadata.
        $workbookScanStatus = $null
        while ((-not (ConvertTo-BoolValue $Shared.StartupMonitoringReady $false)) -and
               [string]::IsNullOrWhiteSpace([string]$Shared.FatalError)) {
            $message = [string]$Shared.StartupMessage
            if ([string]::IsNullOrWhiteSpace($message)) { $message = 'Starting file and folder monitors...' }
            $total = [Math]::Max(0, [int]$Shared.StartupTotal)
            $current = [Math]::Min([int]$Shared.StartupCurrent, $total)
            $ratio = $(if ($total -gt 0) { $current / [double]$total } else { 1.0 })
            $percent = 36 + [int](28 * $ratio)
            $detail = $(if ($total -gt 0) {
                '{0} of {1} file/folder monitors ready. Waiting is normal while network locations connect.' -f $current, $total
            } else {
                'No file/folder monitors need to be started.'
            })
            $startupLines = @(
                '[OK] Step 1 of 4 - Application components loaded',
                '[OK] Step 2 of 4 - Settings and rules loaded',
                ('[..] Step 3 of 4 - File and folder monitors ({0}/{1} ready)' -f $current, $total),
                '[  ] Step 4 of 4 - Workbook information',
                ''
            ) + @(Get-StartupMonitoringLines -Shared $Shared)
            Update-StartupSplash -Splash $Splash `
                -Message ('Step 3 of 4 - {0}' -f $message) `
                -Detail $detail -Percent $percent `
                -ActivityText ($startupLines -join [Environment]::NewLine)
            Start-Sleep -Milliseconds 60
        }

        # Begin workbook metadata only after monitoring has really been armed.
        # This makes the progress sequence truthful and prevents network-heavy
        # workbook reads from competing with watcher startup.
        if ([string]::IsNullOrWhiteSpace([string]$Shared.FatalError)) {
            [void](Start-WorkbookInfoBackgroundScan -Force)
        }

        # Keep the splash foreground until the initial rule-grid metadata is
        # complete. The engine continues revalidating its watchers in parallel.
        if ([string]::IsNullOrWhiteSpace([string]$Shared.FatalError)) {
            while ($true) {
                $workbookScanStatus = Receive-WorkbookInfoBackgroundScan
                if ($workbookScanStatus.Updated) { Update-RuleList }
                if (-not $workbookScanStatus.Running) { break }

                $fileName = [string]$workbookScanStatus.CurrentPath
                try { $fileName = Split-Path -Leaf $fileName } catch { }
                $total = [Math]::Max(1, [int]$workbookScanStatus.Total)
                $displayCurrent = [Math]::Min($total, ([int]$workbookScanStatus.Current + 1))
                $percent = 66 + [int](30 * ([int]$workbookScanStatus.Current / [double]$total))
                $workbookLine = '[..] Workbook {0} of {1}: {2}' -f $displayCurrent, $total, $fileName
                if (-not [string]::IsNullOrWhiteSpace([string]$workbookScanStatus.LastResult)) {
                    $workbookLine = '[OK] {0}{1}[..] Workbook {2} of {3}: {4}' -f `
                        [string]$workbookScanStatus.LastResult, [Environment]::NewLine,
                        $displayCurrent, $total, $fileName
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$workbookScanStatus.Error)) {
                    $workbookLine = '[!!] Workbook information: {0}' -f [string]$workbookScanStatus.Error
                }

                $monitoringNow = $(if (ConvertTo-BoolValue $Shared.StartupMonitoringReady $false) {
                    '[OK] Step 3 of 4 - File and folder monitors ready'
                } else {
                    '[..] Step 3 of 4 - Reconnecting file and folder monitors'
                })
                $phaseLines = @(
                    '[OK] Step 1 of 4 - Application components loaded',
                    '[OK] Step 2 of 4 - Settings and rules loaded',
                    $monitoringNow,
                    ('[..] Step 4 of 4 - Workbook information ({0}/{1})' -f [int]$workbookScanStatus.Current, $total),
                    ''
                )
                Update-StartupSplash -Splash $Splash `
                    -Message 'Step 4 of 4 - Reading workbook information...' `
                    -Detail ('{0} read, {1} unavailable. Monitoring is still being checked in the background.' -f [int]$workbookScanStatus.Succeeded, [int]$workbookScanStatus.Failed) `
                    -ActivityText (($phaseLines + @(Get-StartupMonitoringLines -Shared $Shared) + @($workbookLine)) -join [Environment]::NewLine) `
                    -Percent $percent
                Start-Sleep -Milliseconds 60
            }

            # The last worker status already has the final counters. Publish the
            # UI-side gate and let the engine perform one final watcher check.
            Update-RuleList
            $Shared.StartupUiReady = $true

            while (((-not (ConvertTo-BoolValue $Shared.StartupReady $false)) -or
                    [string]$Shared.Status -ne 'Running') -and
                   [string]::IsNullOrWhiteSpace([string]$Shared.FatalError)) {
                $monitorTotal = [Math]::Max(0, [int]$Shared.StartupTotal)
                $monitorCurrent = [Math]::Min([int]$Shared.StartupCurrent, $monitorTotal)
                $monitorLine = $(if ($monitorCurrent -ge $monitorTotal) {
                    '[..] Final check - confirming monitoring is still active'
                } else {
                    '[..] Final check - restoring file/folder monitors ({0}/{1} ready)' -f $monitorCurrent, $monitorTotal
                })
                $workbookOutcome = $(if ($null -ne $workbookScanStatus -and [int]$workbookScanStatus.Failed -gt 0) {
                    '[!!] Step 4 of 4 - Workbook information loaded with some unavailable files'
                } else {
                    '[OK] Step 4 of 4 - Workbook information loaded'
                })
                Update-StartupSplash -Splash $Splash `
                    -Message 'Final check - preparing the Dashboard...' `
                    -Detail ([string]$Shared.StatusDetail) `
                    -ActivityText ((@(
                        '[OK] Step 1 of 4 - Application components loaded',
                        '[OK] Step 2 of 4 - Settings and rules loaded',
                        $monitorLine,
                        $workbookOutcome,
                        ''
                    ) + @(Get-StartupMonitoringLines -Shared $Shared)) -join [Environment]::NewLine) `
                    -Percent 99
                Start-Sleep -Milliseconds 60
            }
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$Shared.FatalError)) {
            Update-StartupSplash -Splash $Splash -Message 'The monitoring engine could not start.' -Detail ([string]$Shared.FatalError) -Percent 100
        }
        else {
            # This is the sole normal path to closing the splash.
            # The explicit status assertion makes a Degraded first frame
            # impossible even if a future startup refactor changes one gate.
            if (-not (ConvertTo-BoolValue $Shared.StartupReady $false) -or [string]$Shared.Status -ne 'Running') {
                throw ('Dashboard startup gate failed. Status={0}, Ready={1}' -f [string]$Shared.Status, [string]$Shared.StartupReady)
            }

            $finalDetail = $(if ($null -ne $workbookScanStatus -and [int]$workbookScanStatus.Total -gt 0) {
                'Workbook information loaded: {0} read, {1} unavailable. Monitoring status: Running.' -f `
                    [int]$workbookScanStatus.Succeeded, [int]$workbookScanStatus.Failed
            } else {
                'Monitoring status: Running. No workbook information needed to be loaded.'
            })
            $workbookOutcome = $(if ($null -ne $workbookScanStatus -and [int]$workbookScanStatus.Failed -gt 0) {
                '[!!] Some workbook information could not be read'
            } else {
                '[OK] Workbook information loaded'
            })
            $finalLines = @((Get-StartupMonitoringLines -Shared $Shared)) + @($workbookOutcome)
            Update-StartupSplash -Splash $Splash `
                -Message 'Ready - opening Dashboard...' `
                -Detail $finalDetail `
                -ActivityText ((@(
                    '[OK] Step 1 of 4 - Application components loaded',
                    '[OK] Step 2 of 4 - Settings and rules loaded',
                    '[OK] Step 3 of 4 - File and folder monitors ready',
                    '[OK] Step 4 of 4 - Workbook information loaded',
                    '',
                    '[OK] Status - Running'
                ) + $finalLines) -join [Environment]::NewLine) `
                -Percent 100
        }

        # A fast machine can otherwise flash the splash too briefly to notice.
        try {
            $remaining = 900 - [int](((Get-Date) - [DateTime]$Splash.Tag.ShownAt).TotalMilliseconds)
            if ($remaining -gt 0) { Start-Sleep -Milliseconds $remaining }
        }
        catch { Start-Sleep -Milliseconds 180 }
        Close-StartupSplash -Splash $Splash
    }

    # Only now may the Dashboard become visible. This ordering also prevents
    # welcome/logon timers from opening another window while the splash is
    # still reporting workbook progress.
    if ($startMinimized -and $minimizeToTray) {
        $form.ShowInTaskbar = $false
    }
    else {
        Set-FormWithinWorkingArea -Form $form
        $form.Show()
        if ($startMinimized) { $form.WindowState = 'Minimized' }
    }
    $script:UiControls.Tray.Visible = $true

    # A first run has nothing configured, so show the prepared Dashboard even
    # when start-minimized is enabled, then explain the first action.
    if ($firstRunWelcomeDue) {
        $welcomeTimer = New-Object System.Windows.Forms.Timer
        $welcomeTimer.Interval = 400
        $welcomeTimer.add_Tick({
            $this.Stop(); $this.Dispose()
            try {
                Show-DashboardWindow
                Show-FirstRunWelcome
            }
            catch {
                Write-AppLog -Level 'ERROR' -Message ('Welcome screen failed: {0}' -f $_.Exception.Message)
            }
        })
        $welcomeTimer.Start()
    }

    # Start the logon countdown only after startup preparation is finished.
    $script:UiLogonPromptDone = $true
    if ($StartedFromLogon -and (Get-LogonRuleList -Config $script:UiConfig).Count -gt 0) {
        $script:UiLogonPromptDone = $false
        $delaySeconds = ConvertTo-IntValue $script:UiConfig.appSettings.startupPromptDelaySeconds 30 0
        $logonTimer = New-Object System.Windows.Forms.Timer
        $logonTimer.Interval = [Math]::Max(1000, $delaySeconds * 1000)
        $logonTimer.Add_Tick({
            $script:UiControls.LogonTimer.Stop()
            if ($script:UiLogonPromptDone) { return }
            $script:UiLogonPromptDone = $true
            try {
                Invoke-LogonRefreshDecision -Shared $script:UiShared -Config $script:UiConfig
            }
            catch {
                Write-AppLog -Level 'ERROR' -ErrorType 'UnexpectedError' `
                    -Message ('Logon refresh prompt failed: {0}' -f $_.Exception.Message)
            }
        })
        $logonTimer.Start()
        $script:UiControls.LogonTimer = $logonTimer
        Write-AppLog -Level 'INFO' `
            -Message ('Logon rules found. The refresh prompt will appear in {0} seconds.' -f $delaySeconds)
    }

    # Visual polling begins only after startup preparation, preventing timer
    # dialogs or scan polling from competing with the splash loop.
    $timer.Start()
    try { if ($script:UiControls.ContainsKey('ActivateTimer')) { $script:UiControls.ActivateTimer.Start() } } catch { }

    # The tray icon owns the application lifetime; no form is passed here.
    [System.Windows.Forms.Application]::Run()
}

function Register-DashboardHandlers {
    $form = $script:UiControls.Form

    $script:UiControls.BtnAdd.Add_Click({ Invoke-AddFileWizard })
    $script:UiControls.BtnEmptyAdd.Add_Click({ Invoke-RuleAdd })
    $script:UiControls.LinkEmptyHelp.Add_Click({ Show-ExcelRefreshRulesDialog })
    $script:UiControls.BtnEdit.Add_Click({ Invoke-RuleEdit })
    $script:UiControls.BtnDelete.Add_Click({ Invoke-RuleDelete })
    $script:UiControls.BtnToggle.Add_Click({ Invoke-RuleToggle })
    $script:UiControls.BtnRunNow.Add_Click({ Invoke-RuleRunNow })
    $script:UiControls.BtnRunFile.Add_Click({ Invoke-WorkbookRefreshNow })
    $script:UiControls.BtnCancelJob.Add_Click({ Invoke-CurrentJobCancel })
    $script:UiControls.BtnRefreshRules.Add_Click({ Show-ExcelRefreshRulesDialog })
    $script:UiControls.MenuTestDataSources.Add_Click({ Invoke-RuleDataSourceTest })
    $script:UiControls.MenuAddAdvanced.Add_Click({ Invoke-RuleAdd })
    $script:UiControls.MenuQueries.Add_Click({ Invoke-RuleQueryScope })
    $script:UiControls.BtnQueries.Add_Click({ Invoke-RuleQueryScope })
    $script:UiControls.Activity.Add_DoubleClick({ Invoke-ActivityRowActivate })
    $script:UiControls.ActivityLevel.Add_SelectedIndexChanged({ Update-ActivityFilter })
    $script:UiControls.ActivityRule.Add_SelectedIndexChanged({ Update-ActivityFilter })
    $script:UiControls.ActivityFind.Add_TextChanged({ Update-ActivityFilter })
    $script:UiControls.ActivityClear.Add_Click({ Reset-ActivityFilter })
    $script:UiControls.Pending.Add_SelectedIndexChanged({ Update-PendingButtons })
    $script:UiControls.BtnPendingUp.Add_Click({ Invoke-PendingJobMove -Delta -1 })
    $script:UiControls.BtnPendingDown.Add_Click({ Invoke-PendingJobMove -Delta 1 })
    $script:UiControls.BtnPendingRemove.Add_Click({ Invoke-PendingJobRemove })
    $script:UiControls.Rules.Add_DoubleClick({ Invoke-RuleListDoubleClick })
    $script:UiControls.Rules.Add_SelectedIndexChanged({ Update-SelectedRuleAskBeforeCheck })
    $script:UiControls.ChkAskBeforeRefresh.Add_CheckedChanged({ Invoke-SelectedRuleAskBeforeToggle })
    $script:UiControls.Rules.Add_ColumnClick({
        param($sender, $columnArgs)
        if ($script:UiRuleDragActive) { return }
        $column = [int]$columnArgs.Column
        if ($column -lt 0 -or $column -ge $script:UiRuleColumnNames.Count) { return }

        if ($script:UiRuleSortColumn -eq $column) {
            $script:UiRuleSortDirection = $(if ($script:UiRuleSortDirection -eq 'Ascending') { 'Descending' } else { 'Ascending' })
        }
        else {
            $script:UiRuleSortColumn = $column
            $script:UiRuleSortDirection = 'Ascending'
        }
        Update-RuleColumnHeaders
        Update-RuleList
    })
    $script:UiControls.Rules.Add_ItemDrag({
        param($sender, $dragArgs)
        if ($dragArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $item = $dragArgs.Item
        if ($null -eq $item -or $item.Tag -isnot [hashtable]) { return }
        $ruleId = [string]$item.Tag.RuleId
        if ([string]::IsNullOrWhiteSpace($ruleId)) { return }

        $item.Selected = $true
        $data = New-Object System.Windows.Forms.DataObject
        $data.SetData($script:RuleDragDataFormat, $ruleId)
        $script:UiControls.RulesHoverToolTip.Hide($sender)
        $script:UiControls.RulesHoverState.Key = ''
        $script:UiRuleDragActive = $true
        try {
            [void]$sender.DoDragDrop($data, [System.Windows.Forms.DragDropEffects]::Move)
        }
        finally {
            $script:UiRuleDragActive = $false
            Clear-RuleDragInsertionMark
            if ($script:UiRuleListRefreshPending) { Update-RuleList }
        }
    })
    $script:UiControls.Rules.Add_DragEnter({
        param($sender, $dragArgs)
        if ($dragArgs.Data.GetDataPresent($script:RuleDragDataFormat)) {
            $dragArgs.Effect = [System.Windows.Forms.DragDropEffects]::Move
        }
        else {
            $dragArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    })
    $script:UiControls.Rules.Add_DragOver({
        param($sender, $dragArgs)
        if (-not $dragArgs.Data.GetDataPresent($script:RuleDragDataFormat)) {
            $dragArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
            Clear-RuleDragInsertionMark
            return
        }
        $dragArgs.Effect = [System.Windows.Forms.DragDropEffects]::Move
        Update-RuleDragInsertionMark -ListView $sender -DragEventArgs $dragArgs
    })
    $script:UiControls.Rules.Add_DragLeave({ Clear-RuleDragInsertionMark })
    $script:UiControls.Rules.Add_DragDrop({
        param($sender, $dragArgs)
        if (-not $dragArgs.Data.GetDataPresent($script:RuleDragDataFormat)) {
            Clear-RuleDragInsertionMark
            return
        }

        # Refresh the marker once at the exact release point. This also handles
        # a fast drag whose final mouse position did not raise another DragOver.
        Update-RuleDragInsertionMark -ListView $sender -DragEventArgs $dragArgs
        $insertIndex = [int]$sender.InsertionMark.Index
        if ($insertIndex -ge 0 -and [bool]$sender.InsertionMark.AppearsAfterItem) { $insertIndex++ }
        $ruleId = [string]$dragArgs.Data.GetData($script:RuleDragDataFormat)
        Clear-RuleDragInsertionMark
        if ($insertIndex -ge 0 -and -not [string]::IsNullOrWhiteSpace($ruleId)) {
            # DoDragDrop is synchronous. Mark the drag complete before the move
            # rebuilds the list so Update-RuleList is not deferred.
            $script:UiRuleDragActive = $false
            Move-TriggerRule -RuleId $ruleId -InsertionIndex $insertIndex
        }
    })

    # A right-click selects the row under the cursor first, otherwise the menu
    # would act on whatever happened to be selected before.
    $script:UiControls.Rules.Add_MouseDown({
        param($sender, $mouseArgs)
        if ($mouseArgs.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
        $hit = $script:UiControls.Rules.HitTest($mouseArgs.X, $mouseArgs.Y)
        if ($null -ne $hit.Item) { $hit.Item.Selected = $true }
    })
    $script:UiControls.Rules.Add_MouseMove({
        param($sender, $mouseArgs)
        $list  = $script:UiControls.Rules
        $tip   = $script:UiControls.RulesHoverToolTip
        $state = $script:UiControls.RulesHoverState
        $hit   = $list.HitTest($mouseArgs.X, $mouseArgs.Y)
        if ($null -eq $hit.Item) {
            if (-not [string]::IsNullOrWhiteSpace([string]$state.Key)) {
                $tip.Hide($list)
                $state.Key = ''
            }
            return
        }

        # ListView.HitTest is reliable for the row but not for subitems at every
        # DPI setting, so derive the hovered field from the actual column widths.
        $columnIndex = -1
        $left = 0
        for ($i = 0; $i -lt $list.Columns.Count; $i++) {
            $right = $left + [int]$list.Columns[$i].Width
            if ($mouseArgs.X -ge $left -and $mouseArgs.X -lt $right) {
                $columnIndex = $i
                break
            }
            $left = $right
        }

        if ($columnIndex -lt 0) {
            $tip.Hide($list)
            $state.Key = ''
            return
        }

        $text = ''
        if ($hit.Item.Tag -is [hashtable] -and $hit.Item.Tag.ContainsKey('FieldTooltips')) {
            $fieldTips = $hit.Item.Tag.FieldTooltips
            if ($fieldTips -is [hashtable] -and $fieldTips.ContainsKey($columnIndex)) {
                $text = [string]$fieldTips[$columnIndex]
            }
        }

        $key = '{0}|{1}' -f $hit.Item.Index, $columnIndex
        if ($state.Key -eq $key) { return }

        $tip.Hide($list)
        $state.Key = $key
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $tip.Show($text, $list, $mouseArgs.X + 14, $mouseArgs.Y + 18, 20000)
        }
    })
    $script:UiControls.Rules.Add_MouseLeave({
        $script:UiControls.RulesHoverToolTip.Hide($script:UiControls.Rules)
        $script:UiControls.RulesHoverState.Key = ''
    })
    $script:UiControls.RulesViewToggle.Add_LinkClicked({ Toggle-RulesPaneMode })
    $script:UiControls.RuleMenu.Add_Opening({
        param($sender, $cancelArgs)
        $rule = Get-SelectedRule
        if ($null -eq $rule) {
            $cancelArgs.Cancel = $true
            return
        }

        Update-RuleContextMenu -Rule $rule
    })
    $script:UiControls.MenuRunWholeRule.Add_Click({ Invoke-RuleRunNow -WholeRule })
    $script:UiControls.MenuOpenMonitoredFile.Add_Click({ Invoke-OpenMonitoredTarget -OpenFile })
    $script:UiControls.MenuOpenMonitoredFolder.Add_Click({ Invoke-OpenMonitoredTarget -OpenFolder })
    $script:UiControls.MenuEditRule.Add_Click({ Invoke-RuleEdit })
    $script:UiControls.MenuEnableDisableRule.Add_Click({ Invoke-RuleToggle })
    $script:UiControls.MenuDeleteRule.Add_Click({ Invoke-RuleDelete })

    $script:UiControls.BtnSettings.Add_Click({ Invoke-SettingsDialog })
    $script:UiControls.Version.Add_Click({ Show-AboutDialog })
    $script:UiControls.BtnOpenLog.Add_Click({ Invoke-OpenLog })
    $script:UiControls.BtnPause.Add_Click({ Invoke-PauseToggle })
    $script:UiControls.BtnTray.Add_Click({ Hide-Dashboard })
    $script:UiControls.BtnExit.Add_Click({ Stop-Application })

    $script:UiControls.Activity.Add_SelectedIndexChanged({ Update-ActivityDetail })

    $script:UiControls.TrayOpenDashboard.Add_Click({ Show-DashboardWindow })
    $script:UiControls.TrayPauseMonitoring.Add_Click({ Set-PauseState $true })
    $script:UiControls.TrayResumeMonitoring.Add_Click({ Set-PauseState $false })
    $script:UiControls.TrayRefreshWorkbook.Add_Click({ Show-DashboardWindow; Invoke-WorkbookRefreshNow })
    $script:UiControls.TrayRunAllManualRules.Add_Click({ Send-EngineCommand -Shared $script:UiShared -Type 'RunAllManual' })
    $script:UiControls.TrayViewLog.Add_Click({ Invoke-OpenLog })
    $script:UiControls.TraySettings.Add_Click({ Show-DashboardWindow; Invoke-SettingsDialog })
    $script:UiControls.TrayExit.Add_Click({ Stop-Application })

    $script:UiControls.Tray.Add_MouseClick({
        param($sender, $mouseArgs)
        if ($mouseArgs.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Show-DashboardWindow
        }
    })

    # Keep double-click support as a harmless fallback for users accustomed to it.
    $script:UiControls.Tray.Add_DoubleClick({ Show-DashboardWindow })

    $form.Add_Resize({
        if ($script:UiControls.ContainsKey('RulesGroup')) {
            Set-RulesPaneMode -Mode ([string]$script:UiConfig.appSettings.rulesPaneMode)
        }
        if ($script:UiControls.Form.WindowState -eq 'Minimized' -and
            (ConvertTo-BoolValue $script:UiConfig.appSettings.minimizeToTray $true)) {
            $script:UiControls.Form.ShowInTaskbar = $false
            $script:UiControls.Form.Hide()
        }
    })

    $form.Add_FormClosing({
        param($sender, $eventArgs)
        if ($script:UiExiting) { return }

        # A running refresh is never interrupted by closing the dashboard.
        if ($script:UiShared.CurrentJob.Active) {
            $eventArgs.Cancel = $true
            if ($eventArgs.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and
                (ConvertTo-BoolValue $script:UiConfig.appSettings.minimizeToTray $true)) {
                Hide-Dashboard
                return
            }
            Stop-Application
            return
        }

        if ($eventArgs.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and
            (ConvertTo-BoolValue $script:UiConfig.appSettings.minimizeToTray $true)) {
            $eventArgs.Cancel = $true
            Hide-Dashboard
            return
        }
        Stop-Application
    })
}

# ------------------------------------------------------------------------------
# Region: window helpers
# ------------------------------------------------------------------------------

function Show-DashboardWindow {
    $form = $script:UiControls.Form
    Set-FormWithinWorkingArea -Form $form
    $form.Show()
    $form.ShowInTaskbar = $true
    $form.WindowState   = 'Normal'
    $form.BringToFront()
    [void]$form.Activate()
}

function Set-RulesPaneMode {
    param([string]$Mode, [switch]$Save)

    $normalized = $(if ($Mode -eq 'ShowAll') { 'ShowAll' } else { 'Compact' })
    $script:UiConfig.appSettings.rulesPaneMode = $normalized

    $height = 194
    if ($normalized -eq 'ShowAll') {
        # Header + every visible rule row + the action strip. The layout helper
        # clamps this when Recent Activity needs the remaining window space.
        $rowCount = [Math]::Max(1, @($script:UiConfig.rules).Count)
        $height = [Math]::Max(194, (110 + (22 * $rowCount)))
    }
    Set-RulesPaneHeight -Height $height

    if ($script:UiControls.ContainsKey('RulesViewToggle')) {
        $script:UiControls.RulesViewToggle.Text = $(if ($normalized -eq 'ShowAll') { 'Compact' } else { 'Show all' })
    }

    if ($Save) {
        try { Save-AppConfiguration -Path $script:UiPaths.ConfigPath -Config $script:UiConfig }
        catch { Write-AppLog -Level 'WARN' -Message ('The Trigger Rules view could not be saved: {0}' -f $_.Exception.Message) }
    }
}

function Toggle-RulesPaneMode {
    $next = $(if ([string]$script:UiConfig.appSettings.rulesPaneMode -eq 'ShowAll') { 'Compact' } else { 'ShowAll' })
    Set-RulesPaneMode -Mode $next -Save
}

function Set-RulesPaneHeight {
    <#  Internal layout helper used by the Compact / Show all modes. #>
    param([int]$Height)

    if (-not $script:UiControls.ContainsKey('RulesGroup') -or
        -not $script:UiControls.ContainsKey('ActivityGroup')) { return }

    $form       = $script:UiControls.Form
    $rules      = $script:UiControls.RulesGroup
    $current    = $script:UiControls.CurrentGroup
    $pending    = $script:UiControls.PendingGroup
    $activity   = $script:UiControls.ActivityGroup
    $minimum    = 154
    $footerTop  = $form.ClientSize.Height - 44
    $minimumActivity = 170
    $maximum = $footerTop - $rules.Top - 6 - $current.Height - 8 - $minimumActivity
    if ($maximum -lt $minimum) { $maximum = $minimum }
    $newHeight = [Math]::Max($minimum, [Math]::Min($Height, $maximum))

    $rules.Height = $newHeight
    $script:UiControls.Rules.Height      = [Math]::Max(76, $newHeight - 78)
    $script:UiControls.RulesEmpty.Height = $script:UiControls.Rules.Height
    $script:UiControls.RulesEmpty.Width  = $script:UiControls.Rules.Width
    # The action row always stays at the bottom of the group.
    $script:UiControls.RuleButtonRow.Top = $newHeight - 45

    # The empty-state help link is hidden only at the very smallest size; the
    # title, explanation and primary Create button remain available.
    foreach ($control in @($script:UiControls.RulesEmpty.Controls)) {
        if ($control -is [System.Windows.Forms.LinkLabel]) {
            $control.Visible = ($script:UiControls.RulesEmpty.Height -ge 108)
        }
    }

    $current.Top = $rules.Bottom + 6
    $pending.Top = $current.Top
    $activity.Top = $current.Bottom + 8
    $activity.Height = [Math]::Max($minimumActivity, $footerTop - $activity.Top)
}

function Hide-Dashboard {
    $form = $script:UiControls.Form
    if (ConvertTo-BoolValue $script:UiConfig.appSettings.minimizeToTray $true) {
        $form.WindowState   = 'Minimized'
        $form.ShowInTaskbar = $false
        $form.Hide()
        $script:UiControls.Tray.Text = 'Excel Query Trigger Manager'
    }
    else {
        $form.WindowState = 'Minimized'
    }
}

function Get-SelectedRowTag {
    $listView = $script:UiControls.Rules
    if ($listView.SelectedItems.Count -eq 0) { return $null }
    return $listView.SelectedItems[0].Tag
}

function Get-SelectedRule {
    <#  Selecting a workbook row counts as selecting the rule it belongs to.  #>
    $tag = Get-SelectedRowTag
    if ($null -eq $tag) { return $null }
    foreach ($rule in @($script:UiConfig.rules)) {
        if ([string]$rule.id -eq [string]$tag.RuleId) { return $rule }
    }
    return $null
}

function Open-PathWithShell {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            [System.Windows.Forms.MessageBox]::Show(
                ('The path is currently unavailable:' + [Environment]::NewLine + $Path),
                'Open path', 'OK', 'Warning') | Out-Null
            return
        }
        Start-Process -FilePath $Path | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            ('The path could not be opened:' + [Environment]::NewLine + $Path + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message),
            'Open path', 'OK', 'Warning') | Out-Null
    }
}

function Open-FolderInExplorer {
    param([string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Folder)) { return }
    try {
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show(
                ('The folder is currently unavailable:' + [Environment]::NewLine + $Folder),
                'Open folder', 'OK', 'Warning') | Out-Null
            return
        }
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Folder) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            ('The folder could not be opened:' + [Environment]::NewLine + $Folder),
            'Open folder', 'OK', 'Warning') | Out-Null
    }
}

function Invoke-OpenMonitoredTarget {
    param([switch]$OpenFile, [switch]$OpenFolder)

    $rule = Get-SelectedRule
    if ($null -eq $rule) { return }
    $type = [string]$rule.trigger.type
    $path = [string]$rule.trigger.path
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    if ($OpenFile) {
        if ($type -ne 'FileChangedSpecific') { return }
        Open-PathWithShell -Path $path
        return
    }

    if ($OpenFolder) {
        if (Test-TriggerUsesFolder $type) { Open-FolderInExplorer -Folder $path; return }
        if ($type -eq 'FileChangedSpecific') { Open-FolderInExplorer -Folder (Split-Path -Parent $path); return }
    }
}

function Update-RuleContextMenu {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    $type = [string]$Rule.trigger.type
    $hasFolder = (Test-TriggerUsesFolder $type)
    $hasSpecificFile = ($type -eq 'FileChangedSpecific')
    $script:UiControls.MenuOpenMonitoredFile.Enabled = $hasSpecificFile
    $script:UiControls.MenuOpenMonitoredFolder.Enabled = ($hasFolder -or $hasSpecificFile)

    foreach ($menu in @($script:UiControls.MenuRefreshWorkbook, $script:UiControls.MenuOpenWorkbook, $script:UiControls.MenuOpenWorkbookFolder)) {
        $menu.DropDownItems.Clear()
    }

    $actions = @($Rule.actions | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.path) })
    $hasActions = ($actions.Count -gt 0)
    $script:UiControls.MenuRefreshWorkbook.Enabled = $hasActions
    $script:UiControls.MenuOpenWorkbook.Enabled = ($hasActions -and -not [bool]$script:UiShared.CurrentJob.Active)
    $script:UiControls.MenuOpenWorkbookFolder.Enabled = $hasActions

    foreach ($action in $actions) {
        $path = [string]$action.path
        $name = Split-Path -Leaf $path

        $refreshItem = $script:UiControls.MenuRefreshWorkbook.DropDownItems.Add($name)
        $refreshPath = $path
        $refreshAction = ConvertTo-HashtableDeep $action
        $refreshItem.Add_Click(({
            Write-AppLog -Level 'INFO' -Message ('Refresh requested for {0}' -f (Split-Path -Leaf $refreshPath))
            Add-QueuedJobFromUi -Rule (New-ManualWorkbookRule -Path $refreshPath -Action $refreshAction) -TriggerSource 'Manual' | Out-Null
        }.GetNewClosure()))

        $openItem = $script:UiControls.MenuOpenWorkbook.DropDownItems.Add($name)
        $openPath = $path
        $openItem.Add_Click(({ Open-PathWithShell -Path $openPath }.GetNewClosure()))

        # The submenu is a list of folders, so it should read as folders. The
        # full path stays on the tooltip for the ones that share a leaf name.
        $folderPath = Split-Path -Parent $path
        $folderLabel = $(if ([string]::IsNullOrWhiteSpace($folderPath)) { $path } else { Split-Path -Leaf $folderPath })
        if ([string]::IsNullOrWhiteSpace($folderLabel)) { $folderLabel = $folderPath }
        $folderItem = $script:UiControls.MenuOpenWorkbookFolder.DropDownItems.Add($folderLabel)
        $folderItem.ToolTipText = $folderPath
        $folderItem.Add_Click(({ Open-FolderInExplorer -Folder $folderPath }.GetNewClosure()))
    }
}

function Show-ExcelRefreshRulesDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Information & Help'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::None
    # Keep the help window useful on 125-150% displays. Content-heavy tabs
    # scroll inside the tab; the window itself should not occupy the screen.
    $form.ClientSize      = New-Object System.Drawing.Size(900, 700)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterParent'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = Get-UiFont
    $appIcon = Join-Path $script:UiPaths.AppRoot 'assets\ExcelQueryTrigger.ico'
    if (Test-Path -LiteralPath $appIcon) {
        try { $form.Icon = New-Object System.Drawing.Icon($appIcon) } catch { $form.Icon = [System.Drawing.SystemIcons]::Information }
    }
    else { $form.Icon = [System.Drawing.SystemIcons]::Information }

    $intro = New-Object System.Windows.Forms.Label
    $intro.Location = New-Object System.Drawing.Point(22, 16)
    $intro.AutoSize = $true
    $intro.MaximumSize = New-Object System.Drawing.Size(840, 0)
    $intro.Font     = Get-UiFont 11 'Bold'
    $intro.Text     = 'How monitoring, triggers, Excel refreshes, and safety controls work'
    $form.Controls.Add($intro)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(18, 52)
    $tabs.Size     = New-Object System.Drawing.Size(864, 570)
    $form.Controls.Add($tabs)

    $normalFont = Get-UiFont
    $boldFont   = Get-UiFont 9 'Bold'

    $newFlowNode = {
        param([System.Windows.Forms.Control]$Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
        # Use a dark outer panel plus a one-pixel inset white label. This avoids
        # native border/DPI artefacts and does not rely on custom Paint events.
        $node = New-Object System.Windows.Forms.Panel
        $node.Location  = New-Object System.Drawing.Point($X, $Y)
        $node.Size      = New-Object System.Drawing.Size($W, $H)
        $node.Margin    = New-Object System.Windows.Forms.Padding(0)
        $node.BackColor = [System.Drawing.SystemColors]::ControlDark

        $caption = New-Object System.Windows.Forms.Label
        $caption.Location  = New-Object System.Drawing.Point(1, 1)
        $caption.Size      = New-Object System.Drawing.Size(($W - 2), ($H - 2))
        $caption.Margin    = New-Object System.Windows.Forms.Padding(0)
        $caption.Padding   = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
        $caption.Text      = $Text
        $caption.TextAlign = 'MiddleCenter'
        $caption.BackColor = [System.Drawing.SystemColors]::Window
        $node.Controls.Add($caption)

        $Parent.Controls.Add($node)
        return $node
    }
    $newArrow = {
        param([System.Windows.Forms.Control]$Parent, [int]$X, [int]$Y, [int]$W = 23)
        $arrow = New-Object System.Windows.Forms.Label
        $arrow.Location  = New-Object System.Drawing.Point($X, $Y)
        # Keep the arrow entirely inside the gap between nodes. The previous
        # 28px label overlapped the following node by several pixels and could
        # visually erase its upper-left border at some DPI settings.
        $arrow.Size      = New-Object System.Drawing.Size($W, 42)
        $arrow.Text      = [char]0x2192
        $arrow.TextAlign = 'MiddleCenter'
        $arrow.Font      = Get-UiFont 13 'Bold'
        $arrow.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($arrow)
    }
    $appendHeading = {
        param([System.Windows.Forms.RichTextBox]$Box, [string]$Text)
        $Box.SelectionStart  = $Box.TextLength
        $Box.SelectionFont   = $boldFont
        $Box.SelectionIndent = 0
        $Box.SelectionRightIndent = 8
        $Box.AppendText($Text + [Environment]::NewLine)
    }
    $appendBody = {
        param([System.Windows.Forms.RichTextBox]$Box, [string]$Text)
        $Box.SelectionStart  = $Box.TextLength
        $Box.SelectionFont   = $normalFont
        # Visually nest explanations under their headings instead of making
        # headings and body copy start at the same column.
        $Box.SelectionIndent = 24
        $Box.SelectionRightIndent = 12
        $Box.AppendText($Text + [Environment]::NewLine + [Environment]::NewLine)
        $Box.SelectionIndent = 0
        $Box.SelectionRightIndent = 0
    }

    # ------------------------------------------------------------------
    # Tab 1: discoverable version and update entry point
    # ------------------------------------------------------------------
    $tabUpdates = New-Object System.Windows.Forms.TabPage
    $tabUpdates.Text = 'About & updates' + [char]0x00A0
    $tabUpdates.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabUpdates.AutoScroll = $true
    $tabUpdates.BackColor = [System.Drawing.SystemColors]::Window
    [void]$tabs.TabPages.Add($tabUpdates)

    $updateGroup = New-Object System.Windows.Forms.GroupBox
    $updateGroup.Text = 'Version and updates'
    $updateGroup.Location = New-Object System.Drawing.Point(18, 18)
    $updateGroup.Size = New-Object System.Drawing.Size(806, 178)
    $tabUpdates.Controls.Add($updateGroup)

    $updateCaption = New-Object System.Windows.Forms.Label
    $updateCaption.Text = 'Installed version'
    $updateCaption.Location = New-Object System.Drawing.Point(20, 28)
    $updateCaption.Size = New-Object System.Drawing.Size(190, 20)
    $updateCaption.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
    $updateGroup.Controls.Add($updateCaption)

    $updateVersion = New-Object System.Windows.Forms.Label
    $updateVersion.Text = (Get-AppVersion)
    $updateVersion.Font = Get-UiFont 17 'Bold'
    $updateVersion.Location = New-Object System.Drawing.Point(20, 50)
    $updateVersion.Size = New-Object System.Drawing.Size(220, 38)
    $updateGroup.Controls.Add($updateVersion)

    $btnOpenUpdates = New-FormAutoButton -Text 'Check for updates...' -MinimumWidth 210
    $btnOpenUpdates.Location = New-Object System.Drawing.Point(560, 46)
    $btnOpenUpdates.Add_Click({ Show-AboutDialog -Owner $form }.GetNewClosure())
    $updateGroup.Controls.Add($btnOpenUpdates)

    $updateExplanation = New-FormWrappedLabel `
        -Text ('Checks the newest public GitHub Release without signing in.' + [Environment]::NewLine +
            'The application does not send reports or create GitHub Issues.') `
        -X 20 -Y 100 -Width 760 -Color ([System.Drawing.Color]::FromArgb(90, 90, 90))
    $updateGroup.Controls.Add($updateExplanation)

    $updateSteps = New-Object System.Windows.Forms.GroupBox
    $updateSteps.Text = 'Updating the application'
    $updateSteps.Location = New-Object System.Drawing.Point(18, 214)
    $updateSteps.Size = New-Object System.Drawing.Size(806, 246)
    $tabUpdates.Controls.Add($updateSteps)

    $updateStep1 = New-FormWrappedLabel `
        -Text ('1.  Check for updates' + [Environment]::NewLine +
            '     View the latest version and its release notes.') `
        -X 20 -Y 30 -Width 760 -Color ([System.Drawing.SystemColors]::ControlText)
    $updateSteps.Controls.Add($updateStep1)

    $updateStep2 = New-FormWrappedLabel `
        -Text ('2.  Choose Install it' + [Environment]::NewLine +
            '     This appears in the update window only when a newer version is available.') `
        -X 20 -Y 86 -Width 760 -Color ([System.Drawing.SystemColors]::ControlText)
    $updateSteps.Controls.Add($updateStep2)

    $updateStep3 = New-FormWrappedLabel `
        -Text ('3.  Let Setup finish' + [Environment]::NewLine +
            '     The package is verified before Setup starts. The old window closes automatically.') `
        -X 20 -Y 142 -Width 760 -Color ([System.Drawing.SystemColors]::ControlText)
    $updateSteps.Controls.Add($updateStep3)

    $updateKeepsData = New-FormWrappedLabel `
        -Text 'Your rules, history, logs, and settings are retained during an update.' `
        -X 20 -Y 204 -Width 760 -Color ([System.Drawing.Color]::FromArgb(35, 105, 60))
    $updateSteps.Controls.Add($updateKeepsData)

    # ------------------------------------------------------------------
    # Tab 2: refresh flow and long-refresh behavior
    # ------------------------------------------------------------------
    $tabRefresh = New-Object System.Windows.Forms.TabPage
    $tabRefresh.Text = 'Excel refresh' + [char]0x00A0
    $tabRefresh.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabRefresh.AutoScroll = $true
    $tabRefresh.BackColor = [System.Drawing.SystemColors]::Window
    [void]$tabs.TabPages.Add($tabRefresh)

    $gbNormal = New-Object System.Windows.Forms.GroupBox
    $gbNormal.Text = 'Normal refresh flow'
    $gbNormal.Location = New-Object System.Drawing.Point(10, 10)
    $gbNormal.Size = New-Object System.Drawing.Size(822, 184)
    $tabRefresh.Controls.Add($gbNormal)

    $flowY = 34
    $nodeW = 88
    $gap = 23
    $x = 16
    foreach ($text in @('Trigger / Manual','Job Queue','Open Excel','Refresh queries','Save workbook','Close Excel','Complete')) {
        & $newFlowNode $gbNormal $text $x $flowY $nodeW 52 | Out-Null
        $x += $nodeW
        if ($text -ne 'Complete') { & $newArrow $gbNormal $x ($flowY + 5) $gap; $x += $gap }
    }
    $lblNormal = New-FormWrappedLabel `
        -Text ('Jobs run one at a time in a dedicated Excel instance.' + [Environment]::NewLine +
            'The workbook is saved and closed only after refresh activity is consistently quiet.' + [Environment]::NewLine +
            'The app does not intentionally refresh the same workbook in two app-owned Excel instances at once.') `
        -X 18 -Y 102 -Width 782 -Color ([System.Drawing.SystemColors]::ControlText)
    $gbNormal.Controls.Add($lblNormal)

    $gbLong = New-Object System.Windows.Forms.GroupBox
    $gbLong.Text = 'Long refresh / warning threshold behavior'
    $gbLong.Location = New-Object System.Drawing.Point(10, 206)
    $gbLong.Size = New-Object System.Drawing.Size(822, 226)
    $tabRefresh.Controls.Add($gbLong)

    $x = 16
    foreach ($text in @('Refresh active','Warn after reached','Warning only','Keep waiting','Refresh finishes','Save / Close')) {
        & $newFlowNode $gbLong $text $x 34 104 52 | Out-Null
        $x += 104
        if ($text -ne 'Save / Close') { & $newArrow $gbLong $x 39 25; $x += 25 }
    }
    $lblLong = New-FormWrappedLabel `
        -Text ('"Warn after" is a warning threshold, not a forced timeout.' + [Environment]::NewLine +
            [char]0x2022 + ' A slow refresh stays open and keeps running until Excel reports completion.' + [Environment]::NewLine +
            [char]0x2022 + ' The app does not save a partly refreshed workbook or kill Excel because a time limit was reached.' + [Environment]::NewLine +
            [char]0x2022 + ' If it is genuinely stuck, use Cancel Job. Exit asks before cancelling an active refresh.') `
        -X 18 -Y 104 -Width 782 -Color ([System.Drawing.SystemColors]::ControlText)
    $gbLong.Controls.Add($lblLong)

    $gbSync = New-Object System.Windows.Forms.GroupBox
    $gbSync.Text = 'Wait for supported connections before saving (recommended)'
    $gbSync.Location = New-Object System.Drawing.Point(10, 446)
    $gbSync.Size = New-Object System.Drawing.Size(822, 118)
    $tabRefresh.Controls.Add($gbSync)
    $lblSync = New-FormWrappedLabel `
        -Text ('Compatible OLE DB and ODBC connections are made to finish before the workbook is saved.' + [Environment]::NewLine +
            'This prevents an incomplete refresh result from being saved. The original connection setting is restored afterwards.' + [Environment]::NewLine +
            'It does not show Excel or change the connection permanently.') `
        -X 18 -Y 28 -Width 782 -Color ([System.Drawing.SystemColors]::ControlText)
    $gbSync.Controls.Add($lblSync)

    # ------------------------------------------------------------------
    # Tab 2: monitoring / triggers
    # ------------------------------------------------------------------
    $tabMonitor = New-Object System.Windows.Forms.TabPage
    $tabMonitor.Text = 'Monitoring & triggers' + [char]0x00A0
    $tabMonitor.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabMonitor.AutoScroll = $true
    $tabMonitor.BackColor = [System.Drawing.SystemColors]::Window
    [void]$tabs.TabPages.Add($tabMonitor)

    $gbMonitorFlow = New-Object System.Windows.Forms.GroupBox
    $gbMonitorFlow.Text = 'File / folder monitoring flow'
    $gbMonitorFlow.Location = New-Object System.Drawing.Point(10, 10)
    $gbMonitorFlow.Size = New-Object System.Drawing.Size(822, 190)
    $tabMonitor.Controls.Add($gbMonitorFlow)

    $x = 18
    foreach ($text in @('Windows file event','Type / name filter','Debounce','File-ready check','Cooldown','Job Queue')) {
        & $newFlowNode $gbMonitorFlow $text $x 34 104 52 | Out-Null
        $x += 104
        if ($text -ne 'Job Queue') { & $newArrow $gbMonitorFlow $x 39 25; $x += 25 }
    }
    $lblMon = New-FormWrappedLabel `
        -Text ('Windows notifies the application when a matching file changes; the application waits idle between events.' + [Environment]::NewLine +
            'It does not repeatedly scan every file, so a large monitored folder does not create a per-second scan.') `
        -X 18 -Y 104 -Width 782 -Color ([System.Drawing.SystemColors]::ControlText)
    $gbMonitorFlow.Controls.Add($lblMon)

    $monitorText = New-Object System.Windows.Forms.RichTextBox
    $monitorText.Location = New-Object System.Drawing.Point(10, 214)
    $monitorText.Size = New-Object System.Drawing.Size(822, 300)
    $monitorText.ReadOnly = $true
    $monitorText.BorderStyle = 'FixedSingle'
    $monitorText.BackColor = [System.Drawing.SystemColors]::Window
    $monitorText.ScrollBars = 'Vertical'
    $monitorText.WordWrap = $true
    $monitorText.DetectUrls = $false
    $tabMonitor.Controls.Add($monitorText)

    & $appendHeading $monitorText 'NEW FILE ADDED TO FOLDER'
    & $appendBody $monitorText 'Use this when the arrival of a new file should start the rule. The File type and optional name filters decide which Windows events are accepted.'
    & $appendHeading $monitorText 'SPECIFIC FILE UPDATED'
    & $appendBody $monitorText 'Use this when one known file is the source. The rule fires when Windows reports a change to that specific file.'
    & $appendHeading $monitorText 'MATCHING FILE UPDATED'
    & $appendBody $monitorText 'Use this when several files in a folder may be updated and only matching extensions / names should count.'
    & $appendHeading $monitorText 'FILE TYPE PRESETS'
    & $appendBody $monitorText 'CSV, Text, Excel files, Any file, and Custom are provided for convenience. Excel files covers .xlsx, .xlsm, .xlsb, and .xls. Custom keeps support for patterns such as ABC*.csv.'
    & $appendHeading $monitorText 'SCHEDULED / LOGON / MANUAL'
    & $appendBody $monitorText 'Scheduled rules enter the same job queue at their configured time. Logon rules can ask before refreshing after the startup delay. Manual rules and Refresh Any Excel File use the same refresh engine without requiring a file event.'
    & $appendHeading $monitorText 'ASK BEFORE REFRESH'
    & $appendBody $monitorText 'For file and scheduled rules, Ask before refresh pauses after the trigger is accepted but before Excel is opened. The popup explains what triggered the rule and which workbook(s) would run. Yes queues the normal job; No skips only that trigger and leaves the rule enabled. Run Now remains immediate because the click itself is already an explicit approval.'
    & $appendHeading $monitorText 'NO STARTUP CATCH-UP / NO PAST-DIFFERENCE SCAN'
    & $appendBody $monitorText ('File monitoring starts from the moment the watcher is created. It does not compare the folder with yesterday or scan for changes that happened while the application was closed.' + [Environment]::NewLine +
        'Example: you leave at 20:00 and the app stops; a CSV arrives at 22:00; you log in at 08:30 the next morning. That 22:00 arrival does NOT fire a trigger at 08:30. Only a new matching Windows file event after monitoring starts can fire the rule.')

    & $appendHeading $monitorText 'IF A NETWORK PATH TEMPORARILY DISAPPEARS'
    & $appendBody $monitorText 'The application stays running. The watcher health check can detect that monitoring needs to be recreated after the path becomes reachable again.'
    $monitorText.SelectionStart = 0
    $monitorText.SelectionLength = 0

    # ------------------------------------------------------------------
    # Tab 3: technical terms + configuration/storage
    # ------------------------------------------------------------------
    $tabTerms = New-Object System.Windows.Forms.TabPage
    $tabTerms.Text = 'Terms & storage' + [char]0x00A0
    $tabTerms.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabTerms.AutoScroll = $true
    $tabTerms.BackColor = [System.Drawing.SystemColors]::Window
    [void]$tabs.TabPages.Add($tabTerms)

    $termsText = New-Object System.Windows.Forms.RichTextBox
    $termsText.Location = New-Object System.Drawing.Point(10, 10)
    $termsText.Size = New-Object System.Drawing.Size(822, 510)
    $termsText.ReadOnly = $true
    $termsText.BorderStyle = 'FixedSingle'
    $termsText.BackColor = [System.Drawing.SystemColors]::Window
    $termsText.ScrollBars = 'Vertical'
    $termsText.WordWrap = $true
    $termsText.DetectUrls = $false
    $tabTerms.Controls.Add($termsText)

    $debounceDefault = 5
    $cooldownDefault = 30
    $warningDefault = 300
    $healthDefault = 60
    try {
        $debounceDefault = ConvertTo-IntValue $script:UiConfig.appSettings.defaultDebounceSeconds 5 0
        $cooldownDefault = ConvertTo-IntValue $script:UiConfig.appSettings.defaultCooldownSeconds 30 0
        $warningDefault = ConvertTo-IntValue $script:UiConfig.appSettings.defaultRefreshWarningSeconds 300 5
        $healthDefault = ConvertTo-IntValue $script:UiConfig.appSettings.watcherHealthCheckSeconds 60 10
    }
    catch { }

    & $appendHeading $termsText 'CURRENT APPLICATION DEFAULTS'
    & $appendBody $termsText (('Combine repeated changes: {0} sec' + [Environment]::NewLine +
        'Minimum repeat interval: {1} sec' + [Environment]::NewLine +
        'Long-refresh warning: {2} sec' + [Environment]::NewLine +
        'Monitoring health check: {3} sec' + [Environment]::NewLine +
        'Workbook macros by default: {4}') -f $debounceDefault, $cooldownDefault, $warningDefault, $healthDefault, $(if (ConvertTo-BoolValue $script:UiConfig.appSettings.allowWorkbookMacrosByDefault $false) { 'Allowed (Trust Center applies)' } else { 'Blocked' }))

    & $appendHeading $termsText 'DEBOUNCE'
    & $appendBody $termsText 'A single file copy or save can generate several Windows notifications (for example Created followed by multiple Changed events). Debounce waits briefly and combines those related notifications into one trigger instead of creating several Excel refresh jobs.'

    & $appendHeading $termsText 'COOLDOWN'
    & $appendBody $termsText 'After a rule fires, the same rule is prevented from firing again for a short minimum interval. This protects against duplicate or closely repeated file events causing unnecessary back-to-back refreshes.'

    & $appendHeading $termsText 'FILE-READY CHECK'
    & $appendBody $termsText 'Before starting Excel, the app can wait until a newly copied or changed source file appears stable. This reduces the chance that Power Query reads a CSV, text file, or workbook while another process is still writing it.'

    & $appendHeading $termsText 'WATCHER HEALTH CHECK'
    & $appendBody $termsText 'A lightweight periodic check confirms that the FileSystemWatcher is still enabled and the configured path is reachable. This is not a scan of every file in the folder.'

    & $appendHeading $termsText 'WARN AFTER / TIMEOUT WARNING'
    & $appendBody $termsText 'The value is a warning threshold only. It tells you that refresh is taking longer than expected; it does not force a cancellation. Each Excel Action may override the application default.'

    & $appendHeading $termsText 'ALL QUERIES / SELECTED QUERIES'
    & $appendBody $termsText 'Each Excel Action can use RefreshAll or refresh selected Power Query workbook connections. Query names are loaded from the workbook when you choose Selected queries. Upstream dependencies may still be evaluated when Power Query requires them. Completion logs include the query names requested/refreshed.'

    & $appendHeading $termsText 'SYNCHRONOUS / FOREGROUND QUERY MODE'
    & $appendBody $termsText 'The Excel Action option "Wait for supported connections to finish before saving" temporarily sets BackgroundQuery = False on compatible OLE DB/ODBC connections. The original values are restored before save. It does not make Excel visible or change the connection permanently.'
    & $appendBody $termsText 'The advanced "safer temporary-file saving" option is per workbook and is off by default. It writes and verifies a temporary copy in the same folder, closes the original without saving, then replaces the original only after the copy is complete. Use it only for workbooks known to hang during normal Save because temporary free space is required.'

    & $appendHeading $termsText 'WORKBOOK MACROS'
    & $appendBody $termsText 'Macros are blocked by default during automated refresh. For .xlsm and .xlsb files you can opt in per Excel Action. Opt-in uses Excel Trust Center policy; the application does not lower macro security. This may be required when Workbook_Open VBA prepares Power Query connections, parameters, or source paths.'

    & $appendHeading $termsText 'JOB QUEUE'
    & $appendBody $termsText 'All triggers and manual refreshes enter one sequential queue. Excel COM automation is intentionally not run in parallel, which reduces instability and avoids simultaneous app-owned writes to the same workbook.'

    & $appendHeading $termsText 'DISPLAY SCALING / HIGH-DPI COMPATIBILITY'
    & $appendBody $termsText 'The UI intentionally uses Windows compatibility scaling so the whole WinForms window scales together on 125%, 150%, 200%, Retina/Parallels, remote-desktop, and mixed-DPI setups. This prioritizes stable geometry over native-pixel sharpness. Dialogs are also clamped to the current Windows working area and become scrollable if necessary.'

    & $appendHeading $termsText 'WHERE RULES AND SETTINGS ARE SAVED'
    $configPath = ''
    $logPath = ''
    try { $configPath = [string]$script:UiPaths.ConfigPath } catch { }
    try { $logPath = [string]$script:UiPaths.LogDir } catch { }
    & $appendBody $termsText (('Excel Query Trigger Manager v{0}' -f (Get-AppVersion)) + [Environment]::NewLine + [Environment]::NewLine + 'Rules and application settings are stored in:' + [Environment]::NewLine + $configPath + [Environment]::NewLine + [Environment]::NewLine + 'Logs are stored in:' + [Environment]::NewLine + $logPath)
    $termsText.SelectionStart = 0
    $termsText.SelectionLength = 0

    # ------------------------------------------------------------------
    # Tab 4: troubleshooting / questions raised during design
    # ------------------------------------------------------------------
    $tabTrouble = New-Object System.Windows.Forms.TabPage
    $tabTrouble.Text = 'Q&A' + [char]0x00A0
    $tabTrouble.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabTrouble.AutoScroll = $true
    $tabTrouble.BackColor = [System.Drawing.SystemColors]::Window
    [void]$tabs.TabPages.Add($tabTrouble)

    $trouble = New-Object System.Windows.Forms.RichTextBox
    $trouble.Location = New-Object System.Drawing.Point(10, 10)
    $trouble.Size = New-Object System.Drawing.Size(822, 510)
    $trouble.ReadOnly = $true
    $trouble.BorderStyle = 'FixedSingle'
    $trouble.BackColor = [System.Drawing.SystemColors]::Window
    $trouble.ScrollBars = 'Vertical'
    $trouble.WordWrap = $true
    $trouble.DetectUrls = $false
    $tabTrouble.Controls.Add($trouble)

    & $appendHeading $trouble 'WHY CAN A POWER QUERY REFRESH TAKE A LONG TIME?'
    & $appendBody $trouble ('A long refresh does not automatically mean failure.' + [Environment]::NewLine +
        'Large source files, a busy network share, slow external data sources, many Power Query transformations, Data Model work, and Excel calculation can all extend the refresh.' + [Environment]::NewLine +
        'The app therefore treats the configured threshold as a warning rather than an automatic stop.')

    & $appendHeading $trouble 'WHAT HAPPENS AFTER THE WARNING THRESHOLD?'
    & $appendBody $trouble ('The Current Job stage changes to "taking longer than expected" and a warning is logged.' + [Environment]::NewLine +
        'Excel stays open. The query keeps running. The workbook is not saved early and is not force-closed.' + [Environment]::NewLine +
        'Use Cancel Job if the refresh is genuinely stuck. Exit asks for confirmation before cancelling an active refresh.')

    & $appendHeading $trouble 'HOW DOES THE APP DECIDE THAT REFRESH IS FINISHED?'
    & $appendBody $trouble ('The app checks workbook connections, QueryTables/ListObjects, and Excel calculation state at a human-paced interval (about 750 ms).' + [Environment]::NewLine +
        'It also calls CalculateUntilAsyncQueriesDone as a final confirmation and requires several consecutive quiet checks before considering the workbook complete.' + [Environment]::NewLine +
        'This conservative behavior is intended to avoid saving the workbook too early.')

    & $appendHeading $trouble 'WHY CAN IT STILL SAY REFRESHING AFTER A TABLE LOOKS UPDATED?'
    & $appendBody $trouble ('One visible table may already have new values while another connection, Power Query operation, Data Model load, or Excel calculation is still active.' + [Environment]::NewLine +
        'The app waits for the whole workbook to become consistently quiet rather than relying on one visible result.')

    & $appendHeading $trouble 'WHAT DOES SELECTED QUERIES MEAN?'
    & $appendBody $trouble ('Selected queries starts refresh only on the Power Query workbook connections you chose.' + [Environment]::NewLine +
        'Power Query can still evaluate upstream dependency queries when they are required to produce the selected result.' + [Environment]::NewLine +
        'Queries without their own refreshable workbook connection are not offered for individual refresh; use All queries when needed.')

    & $appendHeading $trouble 'WHAT HAPPENS AFTER REFRESH ANY EXCEL FILE FINISHES?'
    & $appendBody $trouble ('A result window appears with the workbook, status, elapsed time, whether it was saved, the query names, and the final log message.' + [Environment]::NewLine +
        'The result is also written to the normal application history/log, so a one-off refresh is auditable just like a rule-based refresh.')

    & $appendHeading $trouble 'CANCEL DID NOT SEEM TO DO ANYTHING - WHY?'
    & $appendBody $trouble ('Cancellation is checked between calls into Excel, so it cannot interrupt Excel while it is inside one. Cancel Job is a request to stop, not an immediate stop.' + [Environment]::NewLine +
        'With synchronous queries the RefreshAll call itself blocks until Excel returns from it. While that is happening the request is accepted but cannot be acted on, and the elapsed time keeps rising.' + [Environment]::NewLine +
        'The usual reason Excel stops answering altogether is a data source provider raising its own credential or connection dialog. Two things now handle that, and both only start after you have confirmed Cancel Job.' + [Environment]::NewLine +
        'Outside cancellation an answerable dialog is closed after the configured 15-second grace period; after cancellation that grace is one second.' + [Environment]::NewLine +
        'After about 12 seconds you are asked whether to end the dedicated process. This option exists only before saving starts, and the process id and start time must match the Excel instance identified from its own window handle.' + [Environment]::NewLine +
        'During Saving, Closing, and Releasing Excel, Cancel and Exit wait for that safety-critical step to finish.' + [Environment]::NewLine +
        'Use Test data sources, the Current Job stage, and Open Log to identify what the hidden instance was waiting for.')

    & $appendHeading $trouble 'WHAT IF THE QUERY REALLY IS STUCK?'
    & $appendBody $trouble ('Possible causes include an unavailable source, network interruption, credentials/authentication, a source application that is not responding, or Excel/Power Query itself being hung.' + [Environment]::NewLine +
        'Because Excel is normally hidden, an unexpected login or external prompt may not be obvious. Check Recent Activity and Open Log for the current stage or a real Excel/COM error.' + [Environment]::NewLine +
        'The app intentionally does not kill Excel merely because elapsed time is long.')

    & $appendHeading $trouble 'WHY WAIT FOR SUPPORTED CONNECTIONS BEFORE SAVING?'
    & $appendBody $trouble ('Some Excel connections refresh in the background, allowing RefreshAll to return while work continues.' + [Environment]::NewLine +
        'For supported OLEDB/ODBC connections, synchronous mode disables that background behavior so the application can observe completion more reliably.' + [Environment]::NewLine +
        'Keep it enabled unless a particular workbook is known to require background-query behavior. The original BackgroundQuery setting is restored before save.')

    & $appendHeading $trouble 'WHAT ABOUT .XLSM OR WORKBOOK_OPEN MACROS?'
    & $appendBody $trouble ('Workbook macros are blocked by default during automated refresh.' + [Environment]::NewLine +
        'If an .xlsm/.xlsb workbook uses Workbook_Open to build Power Query connections, parameters, credentials, or source paths, choose Allow macros for that Excel Action.' + [Environment]::NewLine +
        'Allowing macros does not bypass Excel Trust Center policy; Excel still decides whether the macro is trusted and permitted.')

    & $appendHeading $trouble 'CAN I KEEP USING EXCEL WHILE THIS APP IS REFRESHING?'
    & $appendBody $trouble ('Yes. Each refresh job uses a dedicated hidden Excel instance, separate from the Excel window you normally work in. You can continue editing other workbooks normally.' + [Environment]::NewLine +
        'Do not manually open or edit the same workbook that the app is currently refreshing. A heavy Power Query job can also use CPU, memory, disk, or network bandwidth, so other Excel work may temporarily feel slower.')

    & $appendHeading $trouble 'WILL A FILE THAT CHANGED OVERNIGHT TRIGGER WHEN I LOG IN?'
    & $appendBody $trouble ('No. FileSystemWatcher is event-driven and does not look backwards.' + [Environment]::NewLine +
        'Example: app stops at 20:00; a matching CSV arrives at 22:00; monitoring starts again at 08:30. The 22:00 file does not trigger at 08:30. Only a matching event that occurs after monitoring has started can trigger the rule.' + [Environment]::NewLine +
        'If you need an overnight catch-up workflow in the future, that should be a separate explicit feature rather than an implicit scan.')

    & $appendHeading $trouble 'WHAT IF THE WORKBOOK IS ALREADY OPEN?'
    & $appendBody $trouble 'The app checks for an Excel owner/lock condition before refresh. If it cannot safely update and save the workbook, the job fails with a clear error rather than overwriting or closing the user''s workbook.'

    & $appendHeading $trouble 'WHAT IF THE UI CLOSES OR FAILS DURING A REFRESH?'
    & $appendBody $trouble ('Cancel Job provides an explicit escape path if a refresh never returns. Exit asks whether the active refresh should be cancelled; choosing No leaves it running.' + [Environment]::NewLine +
        'If the UI ends unexpectedly but the engine is still alive, the host waits for the current workbook job to finish rather than cancelling it.' + [Environment]::NewLine +
        'The application does not force-kill Excel during global shutdown simply because a fixed number of seconds has passed.')

    & $appendHeading $trouble 'WHEN CAN THE APP FORCE-TERMINATE ITS OWN EXCEL PROCESS?'
    & $appendBody $trouble ('Only the dedicated Excel process created by this application is eligible, never unrelated Excel windows.' + [Environment]::NewLine +
        'After the workbook has been confirmed closed, the app gives Excel several seconds to exit normally. If that dedicated process still remains, it may be terminated.' + [Environment]::NewLine +
        'If workbook closure cannot be confirmed, the Excel process is deliberately left running for safety.')

    & $appendHeading $trouble 'WHAT IF EXCEL RETURNS A REAL ERROR?'
    & $appendBody $trouble ('A genuine Excel/COM error is different from a long refresh. The job is marked Failed and the details are written to the log.' + [Environment]::NewLine +
        'Cleanup is limited to the dedicated Excel instance created for the job; unrelated Excel windows are not intentionally terminated.')

    & $appendHeading $trouble 'WHEN IS THE WORKBOOK SAVED?'
    & $appendBody $trouble 'Only after the refresh-completion checks pass. Reaching the long-refresh warning threshold never causes a partial save.'

    & $appendHeading $trouble 'CAN I CLOSE THE PROGRAM DURING REFRESH?'
    & $appendBody $trouble ('Yes, but only deliberately. Cancel Job asks for confirmation and cancels at the next refresh poll; the workbook is then closed without saving. Exit also asks whether to cancel the active refresh before closing the application.' + [Environment]::NewLine +
        'Choosing No keeps the refresh running. Windows shutdown, forced process termination, or a machine crash are outside the application''s control and should be avoided during refresh.')

    $trouble.SelectionStart = 0
    $trouble.SelectionLength = 0

    $btnClose = New-FormAutoButton -Text 'Close' -MinimumWidth 90
    $btnClose.Location = New-Object System.Drawing.Point(790, 652)
    $btnClose.Anchor = 'Bottom,Right'
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    Set-FormWithinWorkingArea -Form $form
    [void]$form.ShowDialog($script:UiControls.Form)
    $form.Dispose()
}

function Save-UiConfiguration {
    <#  Persists rules.json and asks the engine to re-arm its watchers.  #>
    try {
        Save-AppConfiguration -Path $script:UiPaths.ConfigPath -Config $script:UiConfig
        Send-EngineCommand -Shared $script:UiShared -Type 'Reload'
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            ('The configuration could not be saved:' + [Environment]::NewLine + $_.Exception.Message),
            'Excel Query Trigger', 'OK', 'Error') | Out-Null
        return $false
    }
}

# ------------------------------------------------------------------------------
# Region: rule commands
# ------------------------------------------------------------------------------

function Get-RuleAskBeforeRefreshValue {
    param([hashtable]$Rule)
    if ($null -eq $Rule) { return $false }
    $type = [string]$Rule.trigger.type
    if (Test-TriggerIsLogon $type) { return ([string]$Rule.trigger.logonBehavior -ne 'Automatic') }
    if ($type -eq 'Manual') { return $false }
    return (ConvertTo-BoolValue $Rule.askBeforeRefresh $false)
}

function Update-SelectedRuleAskBeforeCheck {
    if ($null -eq $script:UiControls.ChkAskBeforeRefresh) { return }
    $rule = Get-SelectedRule
    $script:UiUpdatingAskBeforeCheck = $true
    try {
        if ($null -eq $rule) {
            $script:UiControls.ChkAskBeforeRefresh.Checked = $false
            $script:UiControls.ChkAskBeforeRefresh.Enabled = $false
            return
        }
        $type = [string]$rule.trigger.type
        $script:UiControls.ChkAskBeforeRefresh.Checked = Get-RuleAskBeforeRefreshValue -Rule $rule
        # Manual already requires a click. Logon has its own Ask/Automatic setting
        # in the editor, but the dashboard checkbox can still expose/toggle it.
        $script:UiControls.ChkAskBeforeRefresh.Enabled = ($type -ne 'Manual')
    }
    finally { $script:UiUpdatingAskBeforeCheck = $false }
}

function Invoke-SelectedRuleAskBeforeToggle {
    if ($script:UiUpdatingAskBeforeCheck) { return }
    $rule = Get-SelectedRule
    if ($null -eq $rule) { return }
    $type = [string]$rule.trigger.type
    if ($type -eq 'Manual') { Update-SelectedRuleAskBeforeCheck; return }

    $value = [bool]$script:UiControls.ChkAskBeforeRefresh.Checked
    if (Test-TriggerIsLogon $type) {
        $rule.trigger.logonBehavior = $(if ($value) { 'Ask' } else { 'Automatic' })
    }
    else {
        $rule.askBeforeRefresh = $value
    }
    if (Save-UiConfiguration) {
        Update-RuleList
    }
    else {
        Update-SelectedRuleAskBeforeCheck
    }
}


function Show-FirstRunWelcome {
    <#
        Shown once, the first time the dashboard opens with no rules. Three
        sentences and a way straight into making the first rule - a new user
        should not have to find the manual to get started.
    #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Welcome'
    # Tall enough for the third point and both of its lines. The label had
    # 200 px for fourteen lines of text, so the last one fell off the bottom.
    $form.ClientSize      = New-Object System.Drawing.Size(540, 452)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition   = 'CenterParent'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = $script:UiFonts.Regular
    try { $form.Icon = $script:UiControls.Form.Icon } catch { }

    $title = New-Object System.Windows.Forms.Label
    $title.Text     = 'Excel Query Trigger Manager'
    $title.Font     = New-Object System.Drawing.Font($script:UiFonts.Regular.FontFamily, 13, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(22, 20)
    $title.Size     = New-Object System.Drawing.Size(470, 28)
    $form.Controls.Add($title)

    $body = New-Object System.Windows.Forms.Label
    $body.Location = New-Object System.Drawing.Point(22, 56)
    $body.Size     = New-Object System.Drawing.Size(496, 300)
    $body.Text = @'
It refreshes your Excel queries for you, in a separate Excel running in the
background, so you can keep working while it does.

Three things to know:

   1.  A rule is "when to refresh" plus "which workbooks".
        A time of day, or a file arriving in a folder.

   2.  It watches only while it is running.
        Anything that happened while it was closed is not picked up.

   3.  Closing the window does not stop it.
        It keeps running in the tray by the clock. Use Exit to stop it.
'@
    $form.Controls.Add($body)

    $chkAgain = New-Object System.Windows.Forms.CheckBox
    $chkAgain.Text     = 'Show this again next time'
    $chkAgain.Location = New-Object System.Drawing.Point(22, 366)
    $chkAgain.Size     = New-Object System.Drawing.Size(280, 22)
    $form.Controls.Add($chkAgain)

    $btnCreate = New-FormAutoButton -Text 'Create my first rule' -MinimumWidth 170
    $form.Controls.Add($btnCreate)
    $btnCreate.Location = New-Object System.Drawing.Point(180, 402)

    $btnLater = New-FormAutoButton -Text 'Later' -MinimumWidth 100
    $form.Controls.Add($btnLater)
    $btnLater.Location = New-Object System.Drawing.Point(410, 402)
    $btnLater.DialogResult = 'Cancel'
    $form.CancelButton = $btnLater

    $btnCreate.Add_Click({ $form.Tag = 'create'; $form.DialogResult = 'OK'; $form.Close() }.GetNewClosure())

    [void]$form.ShowDialog($script:UiControls.Form)
    $again = $chkAgain.Checked
    $create = ([string]$form.Tag -eq 'create')
    $form.Dispose()

    if (-not $again) {
        try {
            $marker = Join-Path $script:UiPaths.ConfigDir 'welcome-seen.txt'
            Set-Content -LiteralPath $marker -Value (Get-Date).ToString('yyyy-MM-dd HH:mm') -Encoding UTF8
        }
        catch { }
    }
    if ($create) { Invoke-RuleAdd }
}

function Test-FirstRunWelcomeDue {
    if (@($script:UiConfig.rules).Count -gt 0) { return $false }
    try { return (-not (Test-Path -LiteralPath (Join-Path $script:UiPaths.ConfigDir 'welcome-seen.txt'))) }
    catch { return $false }
}

function Invoke-RuleAdd {
    $rule = Show-RuleEditor -Rule (New-RuleTemplate -AppSettings $script:UiConfig.appSettings) -AppSettings $script:UiConfig.appSettings
    if ($null -eq $rule) { return }

    $duplicate = Find-DuplicateRuleDefinition -Candidate $rule -ExistingRules @($script:UiConfig.rules)
    if ($null -ne $duplicate) {
        [System.Windows.Forms.MessageBox]::Show(
            ('An identical rule already exists: "{0}"' -f [string]$duplicate.name) + [Environment]::NewLine + [Environment]::NewLine +
            'Change the trigger or workbook settings instead of creating a duplicate.',
            'Duplicate rule', 'OK', 'Information') | Out-Null
        return
    }

    $rules = New-Object System.Collections.ArrayList
    foreach ($existing in @($script:UiConfig.rules)) { [void]$rules.Add($existing) }
    [void]$rules.Add($rule)
    $script:UiConfig.rules = @($rules.ToArray())

    if (Save-UiConfiguration) { Update-RuleList }
}

function Invoke-RuleEdit {
    $selected = Get-SelectedRule
    if ($null -eq $selected) { return }

    $edited = Show-RuleEditor -Rule $selected -AppSettings $script:UiConfig.appSettings
    if ($null -eq $edited) { return }

    $duplicate = Find-DuplicateRuleDefinition -Candidate $edited -ExistingRules @($script:UiConfig.rules) `
        -ExcludeRuleId ([string]$edited.id)
    if ($null -ne $duplicate) {
        [System.Windows.Forms.MessageBox]::Show(
            ('These settings are already used by rule "{0}".' -f [string]$duplicate.name) + [Environment]::NewLine + [Environment]::NewLine +
            'Change the trigger or workbook settings before saving.',
            'Duplicate rule', 'OK', 'Information') | Out-Null
        return
    }

    $rules = New-Object System.Collections.ArrayList
    foreach ($existing in @($script:UiConfig.rules)) {
        if ([string]$existing.id -eq [string]$edited.id) { [void]$rules.Add($edited) }
        else { [void]$rules.Add($existing) }
    }
    $script:UiConfig.rules = @($rules.ToArray())

    if (Save-UiConfiguration) { Update-RuleList }
}

function Invoke-RuleDelete {
    $selected = Get-SelectedRule
    if ($null -eq $selected) { return }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        ('Delete the rule "{0}"?' -f $selected.name), 'Confirm delete', 'YesNo', 'Warning')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $rules = New-Object System.Collections.ArrayList
    foreach ($existing in @($script:UiConfig.rules)) {
        if ([string]$existing.id -ne [string]$selected.id) { [void]$rules.Add($existing) }
    }
    $script:UiConfig.rules = @($rules.ToArray())

    if (Save-UiConfiguration) { Update-RuleList }
}

function Invoke-RuleToggle {
    $selected = Get-SelectedRule
    if ($null -eq $selected) { return }
    $selected.enabled = -not (ConvertTo-BoolValue $selected.enabled $true)
    if (Save-UiConfiguration) { Update-RuleList }
}

function Invoke-RuleListDoubleClick {
    <# One dashboard row represents one rule, so double-click edits that rule. #>
    Invoke-RuleEdit
}

function Update-RuleColumnHeaders {
    $list = $script:UiControls.Rules
    if ($null -eq $list) { return }
    for ($i = 0; $i -lt $script:UiRuleColumnNames.Count -and $i -lt $list.Columns.Count; $i++) {
        $caption = [string]$script:UiRuleColumnNames[$i]
        if ($i -eq $script:UiRuleSortColumn) {
            $caption += $(if ($script:UiRuleSortDirection -eq 'Descending') { ' ▼' } else { ' ▲' })
        }
        $list.Columns[$i].Text = $caption
    }
}

function Clear-RuleDragInsertionMark {
    $list = $script:UiControls.Rules
    if ($null -ne $list) { $list.InsertionMark.Index = -1 }
}

function Update-RuleDragInsertionMark {
    <# Shows the exact slot that will receive the dragged rule. #>
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.ListView]$ListView,
        [Parameter(Mandatory = $true)]$DragEventArgs
    )

    if ($ListView.Items.Count -eq 0) {
        $ListView.InsertionMark.Index = -1
        return
    }

    $screenPoint = New-Object System.Drawing.Point([int]$DragEventArgs.X, [int]$DragEventArgs.Y)
    $clientPoint = $ListView.PointToClient($screenPoint)
    $target = $ListView.HitTest($clientPoint.X, $clientPoint.Y).Item

    if ($null -eq $target) {
        if ($clientPoint.Y -lt $ListView.Items[0].Bounds.Top) {
            # Above the first row (including the header) means move to the top.
            $ListView.InsertionMark.Index = 0
            $ListView.InsertionMark.AppearsAfterItem = $false
        }
        else {
            # Empty space below the final visible row means append to the end.
            $ListView.InsertionMark.Index = $ListView.Items.Count - 1
            $ListView.InsertionMark.AppearsAfterItem = $true
        }
    }
    else {
        $ListView.InsertionMark.Index = $target.Index
        $ListView.InsertionMark.AppearsAfterItem = ($clientPoint.Y -ge ($target.Bounds.Top + [int]($target.Bounds.Height / 2)))
    }

    # Dragging near an edge scrolls long rule lists one row at a time.
    $markerIndex = [int]$ListView.InsertionMark.Index
    if ($clientPoint.Y -lt 22 -and $markerIndex -gt 0) {
        $ListView.Items[$markerIndex - 1].EnsureVisible()
    }
    elseif ($clientPoint.Y -gt ($ListView.ClientSize.Height - 22) -and $markerIndex -lt ($ListView.Items.Count - 1)) {
        $ListView.Items[$markerIndex + 1].EnsureVisible()
    }
}

function Move-TriggerRule {
    <# Reorders the in-memory rule array, persists it, and rolls back on failure. #>
    param(
        [Parameter(Mandatory = $true)][string]$RuleId,
        [Parameter(Mandatory = $true)][int]$InsertionIndex
    )

    $storedRules = @($script:UiConfig.rules)
    if ($storedRules.Count -lt 2) { return }

    $previousSortColumn = [int]$script:UiRuleSortColumn
    $previousSortDirection = [string]$script:UiRuleSortDirection
    $convertedSortedView = $false
    $originalRules = $storedRules

    # A drag is always a manual-order operation. If a column sort is active,
    # first make the visible order the new manual base, then apply the drop.
    if ($previousSortColumn -ge 0 -and @($script:UiDisplayedRuleIds).Count -eq $storedRules.Count) {
        $byId = @{}
        foreach ($rule in $storedRules) { $byId[[string]$rule.id] = $rule }
        $visibleRules = New-Object System.Collections.ArrayList
        foreach ($visibleId in @($script:UiDisplayedRuleIds)) {
            if ($byId.ContainsKey([string]$visibleId)) { [void]$visibleRules.Add($byId[[string]$visibleId]) }
        }
        if ($visibleRules.Count -eq $storedRules.Count) {
            $originalRules = @($visibleRules.ToArray())
            $convertedSortedView = $true
            $script:UiRuleSortColumn = -1
            $script:UiRuleSortDirection = 'Ascending'
        }
    }

    $sourceIndex = -1
    for ($i = 0; $i -lt $originalRules.Count; $i++) {
        if ([string]$originalRules[$i].id -eq $RuleId) {
            $sourceIndex = $i
            break
        }
    }
    if ($sourceIndex -lt 0) {
        $script:UiRuleSortColumn = $previousSortColumn
        $script:UiRuleSortDirection = $previousSortDirection
        Update-RuleColumnHeaders
        return
    }

    $targetIndex = [Math]::Max(0, [Math]::Min($InsertionIndex, $originalRules.Count))
    # The insertion slot was measured before removing the source row.
    if ($sourceIndex -lt $targetIndex) { $targetIndex-- }
    $finalRules = $originalRules
    if ($targetIndex -ne $sourceIndex) {
        $reordered = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $originalRules.Count; $i++) {
            if ($i -ne $sourceIndex) { [void]$reordered.Add($originalRules[$i]) }
        }
        $reordered.Insert($targetIndex, $originalRules[$sourceIndex])
        $finalRules = @($reordered.ToArray())
    }
    elseif (-not $convertedSortedView) {
        return
    }
    $script:UiConfig.rules = $finalRules

    if (Save-UiConfiguration) {
        Update-RuleColumnHeaders
        Write-AppLog -Level 'INFO' -RuleName ([string]$originalRules[$sourceIndex].name) `
            -Message ('Trigger Rules order changed: position {0} -> {1}.' -f ($sourceIndex + 1), ($targetIndex + 1))
        Update-RuleList
    }
    else {
        $script:UiConfig.rules = $storedRules
        $script:UiRuleSortColumn = $previousSortColumn
        $script:UiRuleSortDirection = $previousSortDirection
        Update-RuleColumnHeaders
        Update-RuleList
    }
}

function Initialize-ExcelWindowHelper {
    <#
        Lets the dashboard close a dialog the dedicated Excel is showing. Only
        top-level windows of one known process id are ever touched.
    #>
    if ('ExcelWindowHelper' -as [type]) { return $true }

    $source = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelWindowHelper
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder buffer, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder buffer, int max);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);

    private const uint WM_CLOSE = 0x0010;

    /// <summary>
    /// A window we may close has something to click and no progress bar.
    /// Excel shows a progress window with the same shape as a dialog while it
    /// downloads a workbook from a network share, and closing that aborts a
    /// perfectly healthy refresh - so a progress bar means hands off.
    /// </summary>
    private static bool LooksAnswerable(IntPtr hWnd)
    {
        bool hasButton = false;
        bool hasProgress = false;
        EnumChildWindows(hWnd, delegate(IntPtr child, IntPtr lParam)
        {
            StringBuilder cls = new StringBuilder(256);
            GetClassName(child, cls, cls.Capacity);
            string name = cls.ToString();
            if (string.Equals(name, "Button", StringComparison.OrdinalIgnoreCase)) { hasButton = true; }
            if (name.IndexOf("progress", StringComparison.OrdinalIgnoreCase) >= 0) { hasProgress = true; }
            return true;
        }, IntPtr.Zero);
        return (hasButton && !hasProgress);
    }

    private static List<IntPtr> FindDialogHandles(int processId)
    {
        List<IntPtr> found = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
        {
            uint owner;
            GetWindowThreadProcessId(hWnd, out owner);
            if (owner != (uint)processId) { return true; }
            if (!IsWindowVisible(hWnd)) { return true; }

            StringBuilder className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            string cls = className.ToString();

            // XLMAIN is the Excel frame itself. Anything else visible and
            // top-level in this process is a dialog sitting on top of it.
            if (cls == "XLMAIN" || cls == "XLDESK" || cls == "EXCEL7") { return true; }

            found.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static string DescribeWindow(IntPtr hWnd)
    {
        StringBuilder caption = new StringBuilder(512);
        GetWindowText(hWnd, caption, caption.Capacity);
        string title = caption.ToString();
        if (title.Length == 0)
        {
            StringBuilder className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            title = className.ToString();
        }
        return title;
    }

    /// <summary>
    /// Titles of the windows this process is showing. Nothing is changed.
    /// With answerableOnly, progress windows are left out.
    /// </summary>
    public static string[] FindDialogs(int processId, bool answerableOnly)
    {
        List<string> titles = new List<string>();
        foreach (IntPtr hWnd in FindDialogHandles(processId))
        {
            if (answerableOnly && !LooksAnswerable(hWnd)) { continue; }
            titles.Add(DescribeWindow(hWnd));
        }
        return titles.ToArray();
    }

    /// <summary>Closes only the windows that are waiting for an answer.</summary>
    public static string[] CloseDialogs(int processId)
    {
        List<string> closed = new List<string>();
        foreach (IntPtr hWnd in FindDialogHandles(processId))
        {
            if (!LooksAnswerable(hWnd)) { continue; }
            string title = DescribeWindow(hWnd);
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            closed.Add(title);
        }
        return closed.ToArray();
    }
}
'@

    try {
        Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
        return $true
    }
    catch {
        Write-AppLog -Level 'WARN' -Message ('The Excel dialog helper is unavailable: {0}' -f $_.Exception.Message)
        return $false
    }
}

function Get-WatchdogSetting {
    param([string]$Name, $Fallback)
    try { return $script:UiConfig.appSettings[$Name] } catch { return $Fallback }
}

$script:UiProgressCaptions = @(
    'downloading', 'uploading', 'opening', 'saving', 'copying', 'moving',
    'loading', 'preparing', 'contacting', 'connecting', 'processing',
    'refreshing', 'calculating', 'publishing', 'printing', 'converting',
    'please wait', 'working', 'checking out', 'checking in', 'synchronizing',
    'syncing', 'retrieving', 'transferring'
)

function Test-DialogIsProgress {
    <#
        Excel shows a progress window while it fetches a workbook from a network
        share or OneDrive. It looks like a dialog, but Excel is working, not
        waiting - closing it aborts a healthy refresh.
    #>
    param([string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return $true }
    $text = $Title.Trim().ToLowerInvariant()
    foreach ($word in $script:UiProgressCaptions) {
        if ($text.StartsWith($word)) { return $true }
    }
    return $false
}

function Reset-ExcelWatchdogState {
    $script:UiDialogFirstSeenAt   = $null
    $script:UiDialogLastTitle     = ''
    $script:UiDialogClosedCount   = 0
    $script:UiCancelRequestedAt   = $null
    $script:UiCancelDialogsClosed = $false
    $script:UiCancelPromptShown   = $false
}

function Update-ExcelWatchdog {
    <#
        A dialog raised inside a hidden, automation-driven Excel is always a
        dead end: nobody can see it, nobody can answer it, and while it stands
        Excel refuses every COM call - which is why cancellation cannot reach it.

        So the dashboard watches for one the whole time a job is running, not
        only after Cancel Job. Closing it turns a silent hang into an ordinary
        error that the existing failure path reports. Nothing is closed when the
        action asked for the Excel window to be shown, because then the dialog
        is there to be read.

        Cancel Job on top of that shortens the timings and, if Excel still will
        not answer, offers to end that one process.
    #>
    if ($script:UiCancelPromptOpen) { return }

    $shared  = $script:UiShared
    $current = $shared.CurrentJob

    if (-not [bool]$current.Active) {
        if ($null -ne $script:UiCancelRequestedAt -or $null -ne $script:UiDialogFirstSeenAt -or $script:UiDialogClosedCount -gt 0) {
            Reset-ExcelWatchdogState
        }
        return
    }

    $excelPid = [int]$shared.OwnedExcelPid
    if ($excelPid -le 0) { return }
    if ([bool]$current.ExcelVisible) { return }
    $autoDismissDialogs = ConvertTo-BoolValue (Get-WatchdogSetting 'autoDismissExcelDialogs' $true) $true
    $canInspectDialogs = $false
    if ($autoDismissDialogs -and (ConvertTo-BoolValue $current.CanInspectDialogs $false)) {
        $canInspectDialogs = Initialize-ExcelWindowHelper
    }

    $cancelWaited = $(if ($null -eq $script:UiCancelRequestedAt) { -1 } else { ((Get-Date) - $script:UiCancelRequestedAt).TotalSeconds })
    $grace = ConvertTo-IntValue (Get-WatchdogSetting 'excelDialogGraceSeconds' 15) 15 1
    # Once cancellation has been asked for there is nothing left to wait for.
    if ($cancelWaited -ge 0) { $grace = 1 }

    # answerableOnly: something to click, and no progress bar.
    $titles = @()
    if ($canInspectDialogs) {
        try { $titles = @([ExcelWindowHelper]::FindDialogs($excelPid, $true)) } catch { $titles = @() }
    }
    $titles = @($titles | Where-Object { -not (Test-DialogIsProgress -Title ([string]$_)) })

    if ($titles.Count -gt 0) {
        $title = [string]$titles[0]
        if ($null -eq $script:UiDialogFirstSeenAt -or $title -ne $script:UiDialogLastTitle) {
            $script:UiDialogFirstSeenAt = Get-Date
            $script:UiDialogLastTitle   = $title
            Write-AppLog -Level 'WARN' -RuleName ([string]$current.RuleName) -Workbook ([string]$current.Workbook) `
                -Message ('The hidden Excel is showing a dialog and cannot continue: "{0}". It will be closed shortly.' -f $title)
        }
        elseif (((Get-Date) - $script:UiDialogFirstSeenAt).TotalSeconds -ge $grace) {
            $closed = @()
            try { $closed = @([ExcelWindowHelper]::CloseDialogs($excelPid)) } catch { }
            $script:UiDialogFirstSeenAt = $null
            $script:UiDialogLastTitle   = ''
            if ($closed.Count -gt 0) {
                $script:UiDialogClosedCount++
                Write-AppLog -Level 'WARN' -RuleName ([string]$current.RuleName) -Workbook ([string]$current.Workbook) `
                    -Message ('Closed the dialog "{0}". The refresh will now fail with an error instead of waiting for an answer.' -f ($closed -join '; '))
            }

            # A provider that keeps re-asking will never succeed unattended.
            if ($script:UiDialogClosedCount -ge 3 -and $null -eq $script:UiCancelRequestedAt) {
                Write-AppLog -Level 'ERROR' -RuleName ([string]$current.RuleName) -Workbook ([string]$current.Workbook) `
                    -ErrorType 'ExcelDialogLoop' `
                    -Message 'Excel asked the same question three times, so the job is being stopped. The data source most likely needs its connection or credentials updated.'
                Request-CurrentJobCancel -Reason 'Watchdog'
            }
        }
        return
    }

    $script:UiDialogFirstSeenAt = $null
    $script:UiDialogLastTitle   = ''

    # No dialog, but cancellation still has not landed: offer the last resort.
    if ($cancelWaited -lt 0 -or $script:UiCancelPromptShown) { return }
    $limit = ConvertTo-IntValue (Get-WatchdogSetting 'cancelEscalationSeconds' 12) 12 3
    if ($cancelWaited -lt $limit -and $script:UiDialogClosedCount -eq 0) { return }
    if ($cancelWaited -lt 4) { return }
    if (-not (ConvertTo-BoolValue $current.CanForceTerminate $false)) { return }

    $excelProcess = $null
    try { $excelProcess = Get-Process -Id $excelPid -ErrorAction SilentlyContinue } catch { $excelProcess = $null }
    # A recycled process id must never be enough to end something else. Match
    # the process start time captured from Excel.Application.Hwnd as well.
    if ($null -eq $excelProcess -or $excelProcess.ProcessName -ne 'EXCEL') { return }
    $expectedStartUtc = $shared.OwnedExcelStartedAtUtc
    if ($null -eq $expectedStartUtc) { return }
    $actualStartUtc = $null
    try { $actualStartUtc = $excelProcess.StartTime.ToUniversalTime() } catch { return }
    if ([DateTime]$actualStartUtc -ne [DateTime]$expectedStartUtc) {
        Write-AppLog -Level 'ERROR' -Message ('Refused to end Excel process {0}: its identity no longer matches the instance started by this application.' -f $excelPid)
        return
    }
    $script:UiCancelPromptShown = $true

    $workbookName = [string]$current.Workbook
    if ([string]::IsNullOrWhiteSpace($workbookName)) { $workbookName = 'the current workbook' }

    $safetyText = 'The application has confirmed that this instance is still before the save step. The refresh result is discarded and the file on disk is unchanged.'
    $message = 'Excel is not answering the cancellation.' + [Environment]::NewLine + [Environment]::NewLine +
               ('Workbook: {0}' -f $workbookName) + [Environment]::NewLine +
               ('Excel process started by this application: {0}' -f $excelPid) + [Environment]::NewLine + [Environment]::NewLine +
               'End that one process now?' + [Environment]::NewLine + [Environment]::NewLine + $safetyText

    $script:UiCancelPromptOpen = $true
    try {
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Excel is not responding', 'YesNo', 'Warning')
    }
    finally {
        $script:UiCancelPromptOpen = $false
    }

    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog -Level 'WARN' -Message ('Excel process {0} was left running at your request. It can be ended from Task Manager.' -f $excelPid)
        return
    }

    $shared.ForcedExcelTermination = $true
    try {
        Stop-Process -Id $excelPid -Force -ErrorAction Stop
        Write-AppLog -Level 'WARN' -Message ('The unresponsive Excel instance (pid {0}) was ended from the dashboard. The job will be reported as cancelled.' -f $excelPid)
    }
    catch {
        $shared.ForcedExcelTermination = $false
        Write-AppLog -Level 'ERROR' -Message ('Could not end Excel process {0}: {1}' -f $excelPid, $_.Exception.Message)
    }
}

function Request-CurrentJobCancel {
    <#  Starts the cancellation clock. Shared by the button and the watchdog.  #>
    param([string]$Reason = 'Dashboard')
    if ($null -ne $script:UiCancelRequestedAt) { return $false }
    if (-not [bool]$script:UiShared.CurrentJob.Active -or
        -not (ConvertTo-BoolValue $script:UiShared.CurrentJob.CanCancel $false)) { return $false }
    # The engine cannot dequeue commands while it is blocked in Excel, so set
    # the cooperative token directly. The log below is the audit trail.
    $script:UiShared.CancelCurrentJob = $true
    $script:UiShared.CancelRequestedAt = Get-Date
    $script:UiCancelRequestedAt   = $script:UiShared.CancelRequestedAt
    $script:UiCancelDialogsClosed = $false
    $script:UiCancelPromptShown   = $false
    try { $script:UiControls.BtnCancelJob.Enabled = $false } catch { }
    Write-AppLog -Level 'WARN' -Message $(if ($Reason -eq 'Watchdog') { 'Cancellation requested automatically after repeated Excel dialogs.' } else { 'Cancellation requested from the dashboard.' })
    return $true
}

function Invoke-CurrentJobCancel {
    <#
        Asks the engine to abandon the running job. The refresh itself is inside
        Excel, so cancellation takes effect at the next poll: the wait stops, the
        workbook is closed without saving and the Excel instance is released.
        This sends the CancelJob command the engine has always understood but
        which nothing in the interface could previously reach.
    #>
    $current = $script:UiShared.CurrentJob
    if (-not $current.Active) { return }

    if (-not (ConvertTo-BoolValue $current.CanCancel $false)) {
        [System.Windows.Forms.MessageBox]::Show(
            ('Excel is currently {0}. This step cannot be interrupted safely. Please wait for it to finish.' -f (Get-StageDisplayText ([string]$current.Stage)).ToLowerInvariant()),
            'Cannot cancel safely', 'OK', 'Information') | Out-Null
        return
    }

    $workbook = [string]$current.Workbook
    if ([string]::IsNullOrWhiteSpace($workbook)) { $workbook = 'the current workbook' }

    $message = 'Stop the refresh that is running?' + [Environment]::NewLine + [Environment]::NewLine +
               ('Workbook: {0}' -f $workbook) + [Environment]::NewLine + [Environment]::NewLine +
               'Cancellation is accepted only before saving starts. Excel may take a few seconds to respond.' + [Environment]::NewLine +
               'If it does not respond, any dialog it is waiting on is closed at once, and you are then asked whether to end that Excel process.'
    $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Cancel refresh', 'YesNo', 'Question')
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    Request-CurrentJobCancel -Reason 'Dashboard'
}

function Invoke-WorkbookRefreshNow {
    <#
        One-off refresh with an explicit review step. The user can choose a
        different file or cancel after seeing exactly what will happen.
    #>
    while ($true) {
        $picked = Show-FilePicker -InitialPath $script:UiLastManualWorkbook -Owner $script:UiControls.Form `
            -Title 'Refresh any Excel file' `
            -Filter 'Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*'
        if ([string]::IsNullOrWhiteSpace($picked)) { return }

        $allowMacros = ConvertTo-BoolValue $script:UiConfig.appSettings.allowWorkbookMacrosByDefault $false
        $extension = [System.IO.Path]::GetExtension($picked).ToLowerInvariant()
        if ($extension -eq '.xlsm' -or $extension -eq '.xlsb') {
            $message = 'This workbook can contain macros.' + [Environment]::NewLine + [Environment]::NewLine +
                'Macros are blocked by default during automated refresh.' + [Environment]::NewLine +
                'Some workbooks use Workbook_Open VBA to create Power Query connections, parameters, or source paths. Those workbooks may not refresh correctly when macros are blocked.' + [Environment]::NewLine + [Environment]::NewLine +
                'Yes  - allow macros for this refresh (Excel Trust Center policy still applies)' + [Environment]::NewLine +
                'No   - keep macros blocked (recommended)' + [Environment]::NewLine +
                'Cancel - return without opening Excel'
            $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Macro-enabled workbook', 'YesNoCancel', 'Warning')
            if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            $allowMacros = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
        }

        $warningSeconds = ConvertTo-IntValue $script:UiConfig.appSettings.defaultRefreshWarningSeconds 300 5
        $review = Show-AdhocRefreshConfirmation -WorkbookPath $picked -AllowWorkbookMacros:$allowMacros `
            -WarningSeconds $warningSeconds -Owner $script:UiControls.Form
        if ($null -eq $review -or [string]$review.Decision -eq 'Cancel') { return }
        if ([string]$review.Decision -eq 'ChooseAnother') {
            $script:UiLastManualWorkbook = $picked
            continue
        }

        $script:UiLastManualWorkbook = $picked
        $refreshDescription = $(if ([string]$review.Action.refreshMethod -eq 'SelectedQueries') { 'Selected queries: ' + (@($review.Action.selectedQueries) -join ', ') } else { 'All queries' })
        Write-AppLog -Level 'INFO' -Message ('Manual refresh requested for {0}. {1}' -f (Split-Path -Leaf $picked), $refreshDescription)
        $adhoc = New-ManualWorkbookRule -Path $picked -Action $review.Action
        Add-QueuedJobFromUi -Rule $adhoc -TriggerSource 'Manual' | Out-Null
        return
    }
}

function Invoke-RuleRunNow {
    <#
        One dashboard row represents one rule, so Run Now always queues the
        whole rule. Per-workbook refresh remains available in the context menu.
    #>
    param([switch]$WholeRule)

    $selected = Get-SelectedRule
    if ($null -eq $selected) { return }
    if (@($selected.actions).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('This rule has no Excel actions yet.', 'Run Now', 'OK', 'Information') | Out-Null
        return
    }

    # Queued from here rather than sent as a command: while Excel is busy the
    # engine is blocked inside a COM call and will not read its command queue,
    # so a command would sit there invisibly until the current job ended.
    Add-QueuedJobFromUi -Rule $selected -TriggerSource 'Manual'
}

function Add-QueuedJobFromUi {
    <#  Puts a job on the shared queue and says so in the activity log.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [string]$TriggerSource = 'Manual'
    )
    $busy = [bool]$script:UiShared.CurrentJob.Active
    Write-AppLog -Level 'INFO' -RuleName ([string]$Rule.name) `
        -Message $(if ($busy) { 'Added to the queue from the dashboard; a refresh is already running.' } else { 'Manual run requested.' })

    $job = New-PendingJob -Shared $script:UiShared -Rule $Rule -TriggerSource $TriggerSource
    $queued = Add-PendingJob -Shared $script:UiShared -Job $job `
        -CoalesceDuplicateWorkbooks (ConvertTo-BoolValue $script:UiConfig.appSettings.coalesceDuplicateWorkbooks $true)
    Update-PendingList
    return $queued
}

function Invoke-RuleDataSourceTest {
    $selected = Get-SelectedRule
    if ($null -eq $selected) { return }
    Show-DataSourceDiagnostics -Rule $selected
}

function Invoke-PauseToggle {
    Set-PauseState (-not [bool]$script:UiShared.Paused)
}

function Set-PauseState {
    param([bool]$Paused)
    if ($Paused) { Send-EngineCommand -Shared $script:UiShared -Type 'Pause' }
    else { Send-EngineCommand -Shared $script:UiShared -Type 'Resume' }
}

function Invoke-OpenLog {
    try {
        $path = Join-Path $script:UiPaths.LogDir ('app-{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
        if (Test-Path -LiteralPath $path) { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"{0}"' -f $path) }
        else { Start-Process -FilePath $script:UiPaths.LogDir }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(('The log could not be opened: {0}' -f $_.Exception.Message),
            'Excel Query Trigger', 'OK', 'Warning') | Out-Null
    }
}

function Invoke-SettingsDialog {
    $updated = Show-SettingsDialog -AppSettings $script:UiConfig.appSettings -Paths $script:UiPaths
    if ($null -eq $updated) { return }
    $script:UiConfig.appSettings = $updated
    if (Save-UiConfiguration) { Update-RuleList }
}

function Stop-Application {
    <#
        Ordered shutdown. The engine may be inside a refresh,
        so we ask, wait, and only then force the runspace down. The engine
        publishes the pid of any Excel instance it owns, which lets us clean up
        even in the forced case.
    #>
    if ($script:UiExiting) { return }
    $script:UiExiting = $true

    $shared = $script:UiShared
    if ($shared.CurrentJob.Active) {
        $workbookName = [string]$shared.CurrentJob.Workbook
        if ([string]::IsNullOrWhiteSpace($workbookName)) { $workbookName = 'the current workbook' }
        if (-not (ConvertTo-BoolValue $shared.CurrentJob.CanCancel $false)) {
            [System.Windows.Forms.MessageBox]::Show(
                ('Excel is currently {0}. The application cannot exit until this write/close step finishes safely.' -f (Get-StageDisplayText ([string]$shared.CurrentJob.Stage)).ToLowerInvariant()),
                'Please wait', 'OK', 'Information') | Out-Null
            $script:UiExiting = $false
            return
        }
        $message = 'A workbook refresh is still running.' + [Environment]::NewLine + [Environment]::NewLine +
                   ('Workbook: {0}' -f $workbookName) + [Environment]::NewLine + [Environment]::NewLine +
                   'Exit anyway? The refresh will be stopped before saving and the workbook closed without saving.' + [Environment]::NewLine +
                   'Choose No to leave it running and close the application later.'
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Refresh in progress', 'YesNo', 'Warning')
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            $script:UiExiting = $false
            return
        }
        Write-AppLog -Level 'WARN' -Message 'Exit requested while a refresh was running; the job will be cancelled.'
        [void](Request-CurrentJobCancel -Reason 'Exit')
    }

    $script:UiControls.Status.Text = 'Status: Stopping...'
    Send-EngineCommand -Shared $shared -Type 'Exit'
    $shared.ShouldExit       = $true
    if ($shared.CurrentJob.Active) { $shared.CancelCurrentJob = $true }

    $deadline = (Get-Date).AddSeconds(20)
    while ($shared.EngineAlive -and (Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
    }

    if ($shared.EngineAlive) {
        $script:UiControls.Status.Text = 'Status: Waiting for Excel to stop safely...'
        [System.Windows.Forms.MessageBox]::Show(
            'Excel has not reached a safe stopping point. The Dashboard will remain open rather than abandoning a hidden Excel process. You can wait for Excel to respond or use Task Manager if you accept the workbook risk.',
            'Exit is still waiting', 'OK', 'Warning') | Out-Null
        $script:UiExiting = $false
        return
    }

    try { $script:UiControls.Timer.Stop() } catch { }
    try { if ($script:UiControls.ContainsKey('ActivateTimer')) { $script:UiControls.ActivateTimer.Stop(); $script:UiControls.ActivateTimer.Dispose() } } catch { }
    try { if ($script:UiControls.ContainsKey('ActivateEvent')) { $script:UiControls.ActivateEvent.Dispose() } } catch { }
    try { if ($script:UiControls.ContainsKey('ActivateAckEvent')) { $script:UiControls.ActivateAckEvent.Dispose() } } catch { }
    try { Stop-BackgroundUpdateCheck } catch { }
    try { Stop-WorkbookInfoBackgroundScan } catch { }
    try { Stop-RefreshApprovalBackgroundCheck } catch { }
    $script:UiLogonPromptDone = $true
    try {
        if ($script:UiControls.ContainsKey('LogonTimer')) {
            $script:UiControls.LogonTimer.Stop()
            $script:UiControls.LogonTimer.Dispose()
        }
    }
    catch { }
    try {
        $script:UiControls.Tray.Visible = $false
        $script:UiControls.Tray.Dispose()
    }
    catch { }

    [System.Windows.Forms.Application]::ExitThread()
    [System.Windows.Forms.Application]::Exit()
}

# ------------------------------------------------------------------------------
# Region: periodic refresh (runs on the UI thread only)
# ------------------------------------------------------------------------------

function Get-RuleListRowKey {
    <#  Identifies a row across rebuilds so the selection survives.  #>
    param($Tag)
    if ($null -eq $Tag) { return '' }
    if ($Tag.Kind -eq 'Workbook') { return ('W|{0}|{1}' -f $Tag.RuleId, $Tag.Path) }
    return ('R|{0}' -f $Tag.RuleId)
}

function Get-RuleWorkbookDisplayText {
    param([hashtable]$Rule)

    $names = New-Object System.Collections.ArrayList
    foreach ($action in @($Rule.actions)) {
        $path = [string]$action.path
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        [void]$names.Add((Split-Path -Leaf $path))
    }
    if ($names.Count -eq 0) { return '-' }
    if ($names.Count -eq 1) { return [string]$names[0] }
    # A comma-separated list is truncated in the dashboard column and looks
    # like one long filename. Make the number of actions the primary signal;
    # every filename and full path is available in the row tooltip.
    return ('{0} workbooks' -f $names.Count)
}

function Get-RelativeTimeText {
    <#
        "12 minutes ago" rather than "08/25 19:00:57". One is read; the other
        has to be worked out. The exact stamp goes on the tooltip.
    #>
    param([datetime]$When)

    $span = (Get-Date) - $When
    if ($span.TotalSeconds -lt 0)    { return $When.ToString('MMM d, HH:mm') }
    if ($span.TotalSeconds -lt 45)   { return 'just now' }
    if ($span.TotalMinutes -lt 2)    { return 'a minute ago' }
    if ($span.TotalMinutes -lt 60)   { return ('{0:0} minutes ago' -f $span.TotalMinutes) }
    if ($span.TotalHours   -lt 2)    { return 'an hour ago' }
    if ($span.TotalHours   -lt 24)   { return ('{0:0} hours ago' -f $span.TotalHours) }
    if ($span.TotalDays    -lt 2)    { return ('yesterday at {0}' -f $When.ToString('HH:mm')) }
    if ($span.TotalDays    -lt 7)    { return ('{0:0} days ago' -f $span.TotalDays) }
    return $When.ToString('MMM d, HH:mm')
}

function Get-FailureExplanation {
    <#
        What went wrong, in a sentence, plus the one thing worth trying next.
        A status code tells you what the engine saw; this tells you what to do.

        Returns @{ Sentence; ActionLabel; Action }  - Action is '' when the
        only sensible next step is to look at the log.
    #>
    param([string]$ErrorType, [string]$Message)

    switch ($ErrorType) {
        'ConnectionUnreachable' {
            return @{ Sentence = 'A data source this file uses could not be reached, so nothing was refreshed.'
                      ActionLabel = 'Check its data sources'; Action = 'Diagnose' }
        }
        'WorkbookNotFound' {
            return @{ Sentence = 'The file is not where it used to be. It may have been moved, renamed, or its drive is disconnected.'
                      ActionLabel = 'Choose the file again'; Action = 'Edit' }
        }
        'WorkbookLocked' {
            return @{ Sentence = 'The file is already open somewhere else, so the refresh was not started.'
                      ActionLabel = 'Try again'; Action = 'Retry' }
        }
        'FileReadyTimeout' {
            return @{ Sentence = 'The file that set this off was still being written when time ran out.'
                      ActionLabel = 'Try again'; Action = 'Retry' }
        }
        'MacroSecurityError' {
            return @{ Sentence = 'Excel would not open this file with macros disabled, and macros are not allowed by default.'
                      ActionLabel = 'Open its settings'; Action = 'Edit' }
        }
        'ExcelOpenError' {
            return @{ Sentence = 'Excel could not open the file. It may be damaged, or on a drive that is not available.'
                      ActionLabel = 'Check its data sources'; Action = 'Diagnose' }
        }
        'ExcelRefreshError' {
            return @{ Sentence = 'Excel started but could not finish the refresh.'
                      ActionLabel = 'Check its data sources'; Action = 'Diagnose' }
        }
        'ExcelDialogLoop' {
            return @{ Sentence = 'Excel kept asking for a connection or a password, which nobody can answer while it runs in the background.'
                      ActionLabel = 'Check its data sources'; Action = 'Diagnose' }
        }
        'Cancelled' {
            return @{ Sentence = 'The update was stopped before it finished.'
                      ActionLabel = 'Try again'; Action = 'Retry' }
        }
    }

    $detail = [string]$Message
    if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'No further detail was reported.' }
    if ($detail.Length -gt 220) { $detail = $detail.Substring(0, 217) + '...' }
    return @{ Sentence = $detail; ActionLabel = 'Open the log'; Action = 'Log' }
}

function Get-RuleLastRunTime {
    <#
        The moment this rule last finished, as a DateTime.

        Get-RuleLastRunInfo formats "MM/dd HH:mm:ss" for display, and reading
        that back is not safe - it carries no year, so it fails outright under
        some regional settings and could be read as another date under others.
        The history record is read directly instead.
    #>
    param([hashtable]$Rule, $State)

    $paths = @{}
    foreach ($action in @($Rule.actions)) {
        $path = [string]$action.path
        if (-not [string]::IsNullOrWhiteSpace($path)) { $paths[$path.ToLowerInvariant()] = $true }
    }

    foreach ($record in (@(Copy-SharedList $script:UiShared.History) | Sort-Object -Property { [string]$_.finishedAt } -Descending)) {
        $matched = ([string]$record.ruleId -eq [string]$Rule.id)
        if (-not $matched -and $paths.Count -gt 0) {
            foreach ($result in @($record.results)) {
                $wb = [string]$result.workbook
                if (-not [string]::IsNullOrWhiteSpace($wb) -and $paths.ContainsKey($wb.ToLowerInvariant())) { $matched = $true; break }
            }
        }
        if (-not $matched) { continue }

        $when = [DateTime]::MinValue
        foreach ($stamp in @([string]$record.finishedAt, [string]$record.startedAt)) {
            if (-not [string]::IsNullOrWhiteSpace($stamp) -and [DateTime]::TryParse($stamp, [ref]$when)) { return $when }
        }
        return [DateTime]::MinValue
    }

    if ($null -ne $State) {
        $when = [DateTime]::MinValue
        if ($null -ne $State.LastRun -and [DateTime]::TryParse([string]$State.LastRun, [ref]$when)) { return $when }
    }
    return [DateTime]::MinValue
}

function Get-RuleStatusText {
    <#
        The Status column in words rather than in the engine's vocabulary.
        'PathUnavailable' tells a developer what happened; 'Folder not found'
        tells the person who has to fix it.

        Returns @{ Text; Explanation }
    #>
    param([string]$Status)

    switch ([string]$Status) {
        'Active'            { return @{ Text = 'Watching';        Explanation = 'The folder is being watched. A matching file will set this rule off.' } }
        'Waiting for time'  { return @{ Text = 'Waiting for time'; Explanation = 'Nothing to do until the scheduled time comes round. The Next Run column says when.' } }
        'Waiting for logon' { return @{ Text = 'Waiting to sign in'; Explanation = 'This rule runs when you sign in to Windows, and is asked about shortly afterwards.' } }
        'Manual only'       { return @{ Text = 'Manual only';     Explanation = 'Nothing sets this rule off by itself. Use Run Now when you want it.' } }
        'Manual'            { return @{ Text = 'Manual only';     Explanation = 'Nothing sets this rule off by itself. Use Run Now when you want it.' } }
        'Paused'            { return @{ Text = 'Paused';          Explanation = 'Monitoring is paused for the whole application. Press Resume Monitoring to carry on.' } }
        'Disabled'          { return @{ Text = 'Turned off';      Explanation = 'This rule is switched off. Use Enable / Disable to switch it back on.' } }
        'PathUnavailable'   { return @{ Text = 'Folder not found'; Explanation = 'The folder or drive being watched cannot be reached right now. It is retried automatically.' } }
        'Misconfigured'     { return @{ Text = 'Needs fixing';    Explanation = 'Something in this rule is incomplete - open it with Edit and check the trigger.' } }
        'Error'             { return @{ Text = 'Problem';         Explanation = 'Watching could not be started. The activity list below has the reason.' } }
    }
    return @{ Text = 'Not started yet'; Explanation = 'No status has been reported for this rule yet. It settles within a few seconds of starting.' }
}

function Get-RuleNextRunText {
    <#
        The next moment a Scheduled rule is due. Everything else has no
        predictable next time, so the column stays empty for it.
    #>
    param([hashtable]$Rule, [hashtable]$State)

    if (-not (ConvertTo-BoolValue $Rule.enabled $true)) { return 'disabled' }
    if (-not (Test-TriggerIsScheduled ([string]$Rule.trigger.type))) { return '' }

    $time = ConvertTo-ScheduleTime ([string]$Rule.trigger.scheduleTime)
    if ($null -eq $time) { return 'check the time' }

    $days = @($Rule.trigger.scheduleDays)
    if ($days.Count -eq 0) { return 'no days set' }

    $due = Get-RuleNextRunTime -Rule $Rule -State $State
    if ($due -le [DateTime]::MinValue) { return '' }

    $offset = [int]($due.Date - (Get-Date).Date).TotalDays
    if ($offset -eq 0) { return ('today {0}' -f $due.ToString('HH:mm')) }
    if ($offset -eq 1) { return ('tomorrow {0}' -f $due.ToString('HH:mm')) }
    return ('{0} {1}' -f $due.ToString('ddd dd/MM'), $due.ToString('HH:mm'))
}

function Get-RuleNextRunTime {
    <# Returns the actual next due time so the column sorts chronologically. #>
    param([hashtable]$Rule, [hashtable]$State)

    if (-not (ConvertTo-BoolValue $Rule.enabled $true)) { return [DateTime]::MinValue }
    if (-not (Test-TriggerIsScheduled ([string]$Rule.trigger.type))) { return [DateTime]::MinValue }

    $time = ConvertTo-ScheduleTime ([string]$Rule.trigger.scheduleTime)
    if ($null -eq $time) { return [DateTime]::MinValue }

    $days = @($Rule.trigger.scheduleDays)
    if ($days.Count -eq 0) { return [DateTime]::MinValue }

    $now  = Get-Date
    $last = ''
    if ($null -ne $State) { $last = [string]$State.LastScheduledRun }

    for ($offset = 0; $offset -le 7; $offset++) {
        $day = $now.Date.AddDays($offset)
        if ($days -notcontains $day.DayOfWeek.ToString()) { continue }
        $due = $day.AddHours($time.Hour).AddMinutes($time.Minute)
        if ($due -lt $now) { continue }
        # A slot already honoured today must not be offered again.
        if ($due.ToString('yyyy-MM-dd HH:mm') -eq $last) { continue }

        return $due
    }
    return [DateTime]::MinValue
}

$script:WorkbookInfoScanTask = $null
$script:WorkbookInfoLastCompleted = [DateTime]::MinValue
$script:WorkbookInfoRefreshPending = $false
$script:WorkbookInfoRefreshInProgress = $false

function Get-ConfiguredWorkbookPaths {
    <#
        Excludes any workbook the running job has open. Reading a workbook holds
        a file handle on it, and Excel finishes a save by replacing the original
        file, so a metadata read that overlaps a save makes that save fail.
    #>
    $busy = @{}
    try {
        if ([bool]$script:UiShared.CurrentJob.Active) {
            $busyPath = [string]$script:UiShared.CurrentJob.Workbook
            if (-not [string]::IsNullOrWhiteSpace($busyPath)) { $busy[$busyPath.ToLowerInvariant()] = $true }
        }
    }
    catch { }

    $paths = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($rule in @($script:UiConfig.rules)) {
        foreach ($action in @($rule.actions)) {
            $path = [string]$action.path
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $key = $path.ToLowerInvariant()
            if ($seen.ContainsKey($key)) { continue }
            if ($busy.ContainsKey($key) -or $busy.ContainsKey((Split-Path -Leaf $path).ToLowerInvariant())) { continue }
            $seen[$key] = $true
            [void]$paths.Add($path)
        }
    }
    return @($paths.ToArray())
}

function Start-WorkbookInfoBackgroundScan {
    <#  Reads workbook files outside the UI thread. A disconnected share or a
        large workbook can therefore never make the Dashboard stop responding. #>
    param([switch]$Force)

    if ($null -ne $script:WorkbookInfoScanTask) { return $false }
    if (-not $Force -and $script:WorkbookInfoLastCompleted -gt [DateTime]::MinValue -and
        ((Get-Date) - $script:WorkbookInfoLastCompleted).TotalSeconds -lt 60) { return $false }

    $paths = @(Get-ConfiguredWorkbookPaths)
    if ($paths.Count -eq 0) {
        $script:WorkbookInfoLastCompleted = Get-Date
        return $false
    }

    $progress = [hashtable]::Synchronized(@{
        Current     = 0
        Total       = $paths.Count
        CurrentPath = ''
        LastResult  = ''
        Succeeded   = 0
        Failed      = 0
        Completed   = $false
        Error       = ''
    })
    $results = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

    # Each scan runs in a brand new runspace, so the reader's own cache starts
    # empty every time and would re-read every workbook in full. The stamps
    # already known to the dashboard are therefore handed in, and an unchanged
    # workbook is never opened.
    $knownStamps = @{}
    foreach ($cachedKey in @($script:WorkbookInfoCache.Keys)) {
        try { $knownStamps[[string]$cachedKey] = [string]$script:WorkbookInfoCache[$cachedKey].Stamp } catch { }
    }

    $ps = [powershell]::Create()
    $worker = {
        param($AppRoot, $WorkbookPaths, $Progress, $Results, $KnownStamps)
        Set-StrictMode -Version 1.0
        $ErrorActionPreference = 'Stop'
        . (Join-Path (Join-Path $AppRoot 'src') 'Common.ps1')
        . (Join-Path (Join-Path $AppRoot 'src') 'UIDiagnostics.ps1')
        . (Join-Path (Join-Path $AppRoot 'src') 'UIWorkbookInfo.ps1')

        try {
            foreach ($pathValue in @($WorkbookPaths)) {
                $path = [string]$pathValue
                $Progress.CurrentPath = $path
                try {
                    $knownStamp = ''
                    $pathKey = $path.ToLowerInvariant()
                    if ($KnownStamps.ContainsKey($pathKey)) { $knownStamp = [string]$KnownStamps[$pathKey] }
                    if (-not [string]::IsNullOrWhiteSpace($knownStamp)) {
                        $stat = $null
                        try { $stat = Get-Item -LiteralPath $path -ErrorAction Stop } catch { $stat = $null }
                        if ($null -ne $stat -and
                            ('{0}|{1}' -f $stat.LastWriteTimeUtc.Ticks, $stat.Length) -eq $knownStamp) {
                            $Progress.Succeeded = [int]$Progress.Succeeded + 1
                            $Progress.LastResult = ('Unchanged: {0}' -f (Split-Path -Leaf $path))
                            $Progress.Current = [int]$Progress.Current + 1
                            continue
                        }
                    }
                    $info = Get-WorkbookInfo -Path $path -Refresh $true
                    $key = $path.ToLowerInvariant()
                    $stamp = ''
                    if ($script:WorkbookInfoCache.ContainsKey($key)) {
                        $stamp = [string]$script:WorkbookInfoCache[$key].Stamp
                    }
                    $Results.Enqueue([pscustomobject]@{ Path = $path; Stamp = $stamp; Info = $info })
                    if ([bool]$info.Exists) {
                        $Progress.Succeeded = [int]$Progress.Succeeded + 1
                        $Progress.LastResult = ('Read: {0}' -f (Split-Path -Leaf $path))
                    }
                    else {
                        $Progress.Failed = [int]$Progress.Failed + 1
                        $Progress.LastResult = ('Unavailable: {0}' -f (Split-Path -Leaf $path))
                    }
                }
                catch {
                    $info = New-EmptyWorkbookInfo
                    $info.Reason = [string]$_.Exception.Message
                    $Results.Enqueue([pscustomobject]@{ Path = $path; Stamp = ''; Info = $info })
                    $Progress.Failed = [int]$Progress.Failed + 1
                    $Progress.LastResult = ('Could not read: {0}' -f (Split-Path -Leaf $path))
                }
                $Progress.Current = [int]$Progress.Current + 1
            }
        }
        catch { $Progress.Error = [string]$_.Exception.Message }
        finally { $Progress.Completed = $true }
    }
    [void]$ps.AddScript($worker)
    [void]$ps.AddArgument([string]$script:UiPaths.AppRoot)
    [void]$ps.AddArgument($paths)
    [void]$ps.AddArgument($progress)
    [void]$ps.AddArgument($results)
    [void]$ps.AddArgument($knownStamps)
    $script:WorkbookInfoScanTask = @{
        PowerShell = $ps
        Handle     = $ps.BeginInvoke()
        Progress   = $progress
        Results    = $results
    }
    return $true
}

function Receive-WorkbookInfoBackgroundScan {
    $status = @{
        Running = $false; Updated = $false; Completed = $false
        Current = 0; Total = 0; CurrentPath = ''; LastResult = ''
        Succeeded = 0; Failed = 0; Error = ''
    }
    $task = $script:WorkbookInfoScanTask
    if ($null -eq $task) { return $status }

    $status.Running     = (-not $task.Handle.IsCompleted)
    $status.Current     = [int]$task.Progress.Current
    $status.Total       = [int]$task.Progress.Total
    $status.CurrentPath = [string]$task.Progress.CurrentPath
    $status.LastResult  = [string]$task.Progress.LastResult
    $status.Succeeded   = [int]$task.Progress.Succeeded
    $status.Failed      = [int]$task.Progress.Failed
    $status.Error       = [string]$task.Progress.Error

    $result = $null
    while ($task.Results.TryDequeue([ref]$result)) {
        try {
            $key = ([string]$result.Path).ToLowerInvariant()
            $before = Get-WorkbookInfo -Path ([string]$result.Path)
            # Objects handed across PowerShell runspaces may arrive as a plain
            # hashtable or as a PSObject wrapper. Normalize both forms instead
            # of silently dropping the Data updated result on a failed cast.
            $after = ConvertTo-HashtableDeep $result.Info
            if ($after -isnot [hashtable]) { throw 'Workbook metadata was returned in an unexpected format.' }
            $script:WorkbookInfoCache[$key] = @{ Stamp = [string]$result.Stamp; Info = $after }
            if ($before.SavedAt -ne $after.SavedAt -or $before.DataRefreshedAt -ne $after.DataRefreshedAt -or
                $before.Location -ne $after.Location -or $before.Exists -ne $after.Exists) {
                $status.Updated = $true
            }
        }
        catch {
            $status.Error = ('Could not apply workbook information for {0}: {1}' -f [string]$result.Path, $_.Exception.Message)
            Write-AppLog -Level 'WARN' -Message $status.Error
        }
        $result = $null
    }

    if ($task.Handle.IsCompleted) {
        try { [void]$task.PowerShell.EndInvoke($task.Handle) } catch { $status.Error = [string]$_.Exception.Message }
        try { $task.PowerShell.Dispose() } catch { }
        $script:WorkbookInfoScanTask = $null
        $script:WorkbookInfoLastCompleted = Get-Date
        $status.Running = $false
        $status.Completed = $true
    }
    return $status
}

function Stop-WorkbookInfoBackgroundScan {
    if ($null -eq $script:WorkbookInfoScanTask) { return }
    try { $script:WorkbookInfoScanTask.PowerShell.Stop() } catch { }
    try { $script:WorkbookInfoScanTask.PowerShell.Dispose() } catch { }
    $script:WorkbookInfoScanTask = $null
}

function Get-UiWorkbookPathKey {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    try { return ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\\').ToLowerInvariant() }
    catch { return $Path.Trim().TrimEnd('\\').ToLowerInvariant() }
}

function Get-SuccessfulQueryRefreshHistoryMap {
    <#
        Local history is the authoritative fallback for refreshes performed by
        this application. Workbook custom properties are useful when Excel lets
        us write them, but some workbooks reject that COM operation. Keeping the
        exact successful refresh time locally preserves Data updated and the
        recent-refresh guard without treating file save time as query refresh.
    #>
    $map = @{}
    foreach ($record in @(Copy-SharedList $script:UiShared.History)) {
        foreach ($result in @($record.results)) {
            if (-not (ConvertTo-BoolValue $result.success $false)) { continue }
            $path = [string]$result.workbook
            $key = Get-UiWorkbookPathKey -Path $path
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            $raw = [string]$result.queryRefreshedAt
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            $when = [DateTime]::MinValue
            if (-not [DateTime]::TryParse($raw, [ref]$when)) { continue }
            if (-not $map.ContainsKey($key) -or $when -gt [DateTime]$map[$key]) { $map[$key] = $when }
        }
    }
    return $map
}

function Get-SuccessfulQueryRefreshFromHistory {
    param([string]$Path)
    $map = Get-SuccessfulQueryRefreshHistoryMap
    $key = Get-UiWorkbookPathKey -Path $Path
    if (-not [string]::IsNullOrWhiteSpace($key) -and $map.ContainsKey($key)) { return [DateTime]$map[$key] }
    return [DateTime]::MinValue
}

function Get-RuleWorkbookInfoSummary {
    <#
        The Where and Data updated cells for one rule, from whatever has
        already been read. Never opens a file, so the grid can be redrawn as
        often as it likes.

        Tooltips are deliberately split by column. "Where" explains location;
        "Data updated" explains query-refresh and save timestamps. This avoids
        the old whole-row tooltip that mixed unrelated information together.
    #>
    param([hashtable]$Rule)

    $summary = @{
        LocationText    = ''
        DataAgeText     = ''
        DataSortTicks   = [int64]0
        LocationTooltip = ''
        DataTooltip     = ''
    }
    $paths = @(@($Rule.actions) | ForEach-Object { [string]$_.path } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($paths.Count -eq 0) { return $summary }

    $locations = New-Object System.Collections.ArrayList
    $oldest    = [DateTime]::MaxValue
    $known     = 0
    $refreshUnknown = 0
    $locationLines = New-Object System.Collections.ArrayList
    $dataLines     = New-Object System.Collections.ArrayList

    foreach ($path in $paths) {
        $info = Get-WorkbookInfo -Path $path
        if (-not $info.Exists) { continue }
        $known++

        $fileName = Split-Path -Leaf $path
        $location = [string]$info.Location
        if ($info.SharedWorkbook -and $location -eq 'Local') { $location = 'Shared' }
        if ($locations -notcontains $location) { [void]$locations.Add($location) }

        if ($paths.Count -gt 1) { [void]$locationLines.Add($fileName) }
        [void]$locationLines.Add(('Location: {0}' -f $location))
        if (-not [string]::IsNullOrWhiteSpace([string]$info.LocationDetail)) {
            [void]$locationLines.Add([string]$info.LocationDetail)
        }
        if ($info.SharedWorkbook) {
            [void]$locationLines.Add('Excel marks this as a shared workbook.')
        }
        if ($paths.Count -gt 1) { [void]$locationLines.Add('') }

        $when = $info.DataRefreshedAt
        $refreshWhoValue = [string]$info.RefreshedBy
        $historyWhen = Get-SuccessfulQueryRefreshFromHistory -Path $path
        if ($historyWhen -gt $when) {
            $when = $historyWhen
            $refreshWhoValue = 'Excel Query Trigger Manager (local history)'
        }
        if ($when -gt [DateTime]::MinValue -and $when -lt $oldest) { $oldest = $when }
        if ($when -le [DateTime]::MinValue) { $refreshUnknown++ }

        if ($paths.Count -gt 1) { [void]$dataLines.Add($fileName) }
        if ($when -gt [DateTime]::MinValue) {
            $refreshWho = $(if ([string]::IsNullOrWhiteSpace($refreshWhoValue)) {
                ''
            } else {
                ' by ' + $refreshWhoValue
            })
            [void]$dataLines.Add(('Query data refreshed: {0}{1}' -f `
                $when.ToString('dddd d MMMM, HH:mm:ss'), $refreshWho))
        }
        else {
            [void]$dataLines.Add('Query data refreshed: not recorded')
        }

        if ($info.SavedAt -gt [DateTime]::MinValue) {
            $saveWho = $(if ([string]::IsNullOrWhiteSpace([string]$info.SavedBy)) {
                ''
            } else {
                ' by ' + [string]$info.SavedBy
            })
            [void]$dataLines.Add(('File saved: {0}{1} (not used for Data updated)' -f `
                $info.SavedAt.ToString('dddd d MMMM, HH:mm:ss'), $saveWho))
        }
        if ($paths.Count -gt 1) { [void]$dataLines.Add('') }
    }

    if ($known -eq 0) {
        $summary.LocationTooltip = 'Workbook location information is not available.'
        $summary.DataTooltip = 'Workbook refresh information is not available.'
        return $summary
    }

    $summary.LocationText = $(if ($locations.Count -eq 1) { [string]$locations[0] } else { 'Mixed' })
    if ($refreshUnknown -gt 0) { $summary.DataAgeText = 'Not recorded' }
    elseif ($oldest -lt [DateTime]::MaxValue) {
        $summary.DataAgeText = Get-RelativeTimeText -When $oldest
        $summary.DataSortTicks = [int64]$oldest.Ticks
    }

    if ($locationLines.Count -gt 0) {
        $summary.LocationTooltip = (@($locationLines.ToArray()) -join [Environment]::NewLine).Trim()
    }
    if ($dataLines.Count -gt 0) {
        $summary.DataTooltip = (@($dataLines.ToArray()) -join [Environment]::NewLine).Trim()
    }
    return $summary
}

function Get-RuleTriggerFieldTooltip {
    param([hashtable]$Rule)

    $trigger = $Rule.trigger
    if ($null -eq $trigger) { return 'Trigger is not configured.' }
    $type = [string]$trigger.type
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add(('Trigger: {0}' -f (Get-TriggerTypeLabel $type)))

    if (Test-TriggerIsScheduled $type) {
        [void]$lines.Add(('Time: {0}' -f [string]$trigger.scheduleTime))
        [void]$lines.Add(('Days: {0}' -f (Format-ScheduleDays @($trigger.scheduleDays))))
    }
    elseif (Test-TriggerIsLogon $type) {
        $behavior = $(if ([string]$trigger.logonBehavior -eq 'Automatic') {
            'Refresh automatically after sign-in'
        } else {
            'Ask before refreshing after sign-in'
        })
        [void]$lines.Add($behavior)
    }
    elseif (Test-TriggerUsesWatcher $type) {
        $folder = ''
        if (Test-TriggerUsesFolder $type) {
            $folder = [string]$trigger.path
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$trigger.path)) {
            try { $folder = Split-Path -Parent ([string]$trigger.path) } catch { $folder = '' }
        }
        if (-not [string]::IsNullOrWhiteSpace($folder)) { [void]$lines.Add(('Folder: {0}' -f $folder)) }
        if ($type -eq 'FileChangedSpecific' -and -not [string]::IsNullOrWhiteSpace([string]$trigger.path)) {
            [void]$lines.Add(('File: {0}' -f (Split-Path -Leaf ([string]$trigger.path))))
        }
        elseif (Test-TriggerUsesFolder $type) {
            $filter = [string]$trigger.filter
            if ([string]::IsNullOrWhiteSpace($filter)) { $filter = '*.*' }
            [void]$lines.Add(('Filter: {0}' -f $filter))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$trigger.contains)) {
            [void]$lines.Add(('Filename contains: {0}' -f [string]$trigger.contains))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$trigger.exclude)) {
            [void]$lines.Add(('Filename excludes: {0}' -f [string]$trigger.exclude))
        }
    }
    else {
        [void]$lines.Add('Runs only when started manually.')
    }

    return (@($lines.ToArray()) -join [Environment]::NewLine)
}

function Get-RuleQueryFieldTooltip {
    param([hashtable]$Rule)

    $actions = @($Rule.actions)
    if ($actions.Count -eq 0) { return 'No workbook action is configured.' }

    $lines = New-Object System.Collections.ArrayList
    foreach ($action in $actions) {
        $workbook = Split-Path -Leaf ([string]$action.path)
        if ([string]::IsNullOrWhiteSpace($workbook)) { $workbook = 'Workbook' }

        if ([string]$action.refreshMethod -eq 'SelectedQueries') {
            $queries = @($action.selectedQueries | ForEach-Object { [string]$_ } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            [void]$lines.Add(('{0}: {1} selected' -f $workbook, $queries.Count))
            foreach ($query in $queries) { [void]$lines.Add(('  - {0}' -f $query)) }
        }
        else {
            [void]$lines.Add(('{0}: all queries' -f $workbook))
        }
        if ($actions.Count -gt 1) { [void]$lines.Add('') }
    }
    return (@($lines.ToArray()) -join [Environment]::NewLine).Trim()
}

function Get-RuleFailureExplanation {
    <#  The plain sentence for this rule's most recent failed run.  #>
    param([hashtable]$Rule)

    $history = @(Copy-SharedList $script:UiShared.History)
    for ($i = $history.Count - 1; $i -ge 0; $i--) {
        $record = $history[$i]
        if ([string]$record.ruleId -ne [string]$Rule.id) { continue }
        if ([string]$record.status -ne 'Failed') { continue }
        foreach ($result in @($record.results)) {
            if (-not [bool]$result.success) {
                return (Get-FailureExplanation -ErrorType ([string]$result.errorType) -Message ([string]$result.message))
            }
        }
        return $null
    }
    return $null
}

function Get-RuleDurationBaseline {
    <#
        The middle duration of this rule's recent finished jobs, in seconds.
        Returns 0 until there are enough runs to say anything useful.
    #>
    param([hashtable]$Rule)

    $seconds = New-Object System.Collections.ArrayList
    foreach ($record in (Copy-SharedList $script:UiShared.History)) {
        if ([string]$record.ruleId -ne [string]$Rule.id) { continue }
        $value = ConvertTo-DurationSeconds ([string]$record.elapsed)
        if ($value -gt 0) { [void]$seconds.Add($value) }
    }
    if ($seconds.Count -lt 3) { return 0 }

    $sorted = @($seconds.ToArray() | Sort-Object)
    return [double]$sorted[[int]([math]::Floor($sorted.Count / 2))]
}

function ConvertTo-DurationSeconds {
    <#  Accepts hh:mm:ss, mm:ss or a plain number of seconds.  #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 0 }
    $trimmed = $Text.Trim()
    $parts = @($trimmed -split ':')
    if ($parts.Count -ge 2) {
        $total = 0.0
        foreach ($part in $parts) {
            $number = 0.0
            if (-not [double]::TryParse($part, [ref]$number)) { return 0 }
            $total = ($total * 60) + $number
        }
        return $total
    }
    $single = 0.0
    if ([double]::TryParse($trimmed, [ref]$single)) { return $single }
    return 0
}

function Get-RuleLastRunInfo {
    <#
        Returns the latest execution timestamp and whole-job duration for a
        rule. Current-session manual one-workbook refreshes are included when
        they refreshed a workbook belonging to the rule. Persisted rule state
        is used after an application restart.
    #>
    param([hashtable]$Rule, [hashtable]$State)

    $paths = @{}
    foreach ($action in @($Rule.actions)) {
        $path = [string]$action.path
        if (-not [string]::IsNullOrWhiteSpace($path)) { $paths[$path.ToLowerInvariant()] = $true }
    }

    $history = @(Copy-SharedList $script:UiShared.History)
    for ($i = $history.Count - 1; $i -ge 0; $i--) {
        $record = $history[$i]
        $matched = ([string]$record.ruleId -eq [string]$Rule.id)
        if (-not $matched -and $paths.Count -gt 0) {
            foreach ($result in @($record.results)) {
                $wb = [string]$result.workbook
                if (-not [string]::IsNullOrWhiteSpace($wb) -and $paths.ContainsKey($wb.ToLowerInvariant())) {
                    $matched = $true
                    break
                }
            }
        }
        if (-not $matched) { continue }

        $when = [DateTime]::MinValue
        $timeText = '-'
        if ([DateTime]::TryParse([string]$record.finishedAt, [ref]$when)) {
            $timeText = $when.ToString('MM/dd HH:mm:ss')
        }
        elseif ([DateTime]::TryParse([string]$record.startedAt, [ref]$when)) {
            $timeText = $when.ToString('MM/dd HH:mm:ss')
        }

        $durationText = [string]$record.elapsed
        if ([string]::IsNullOrWhiteSpace($durationText)) { $durationText = '-' }
        return @{ Time = $timeText; Duration = $durationText; Status = [string]$record.status }
    }

    if ($null -ne $State) {
        $timeText = '-'
        $when = [DateTime]::MinValue
        if ($null -ne $State.LastRun -and [DateTime]::TryParse([string]$State.LastRun, [ref]$when)) {
            $timeText = $when.ToString('MM/dd HH:mm:ss')
        }
        $durationText = [string]$State.LastDuration
        if ([string]::IsNullOrWhiteSpace($durationText)) {
            # Backward compatibility with state files written before LastDuration existed.
            $legacy = [string]$State.LastResultText
            if ($legacy -match '\((\d{2}:\d{2}:\d{2})\)\s*$') { $durationText = $Matches[1] }
        }
        if ([string]::IsNullOrWhiteSpace($durationText)) { $durationText = '-' }
        return @{ Time = $timeText; Duration = $durationText; Status = [string]$State.LastRunStatus }
    }

    return @{ Time = '-'; Duration = '-'; Status = '' }
}

function Update-WorkbookInfoAndRuleList {
    <#  Polls the worker without touching a workbook on the UI thread. No new
        scan is started while a job is running: Data updated can wait a few
        minutes, but a workbook save that collides with a metadata read cannot
        be retried without repeating the whole refresh.  #>
    $scan = Receive-WorkbookInfoBackgroundScan
    if ($scan.Updated) { Update-RuleList }
    if ([bool]$script:UiShared.CurrentJob.Active) {
        # A scan that began before the job must not keep reading now. Anything
        # it had already collected is simply read again after the job.
        if ($scan.Running) { Stop-WorkbookInfoBackgroundScan }
        return
    }
    if (-not $scan.Running) { [void](Start-WorkbookInfoBackgroundScan) }
}

function Update-RuleList {
    <#
        One row per rule. The workbook column shows only workbook file names;
        full paths and trigger source details remain available in the tooltip.
    #>
    if ($script:UiRuleDragActive) {
        $script:UiRuleListRefreshPending = $true
        return
    }
    $script:UiRuleListRefreshPending = $false
    $listView    = $script:UiControls.Rules
    $listView.InsertionMark.Index = -1
    $selectedKey = ''
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedKey = Get-RuleListRowKey $listView.SelectedItems[0].Tag
    }

    $renderedItems = New-Object System.Collections.ArrayList
    $originalIndex = 0
    $listView.BeginUpdate()
    try {
        $listView.Items.Clear()
        foreach ($rule in @($script:UiConfig.rules)) {
            $ruleId  = [string]$rule.id
            $enabled = ConvertTo-BoolValue $rule.enabled $true

            $watcherStatus = 'Unknown'
            $state = $null
            if ($script:UiShared.RuleState.ContainsKey($ruleId)) {
                $state = $script:UiShared.RuleState[$ruleId]
                $watcherStatus = [string]$state.WatcherStatus
            }

            $triggerText = Get-TriggerTypeLabel ([string]$rule.trigger.type) -Short
            $workbookText = Get-RuleWorkbookDisplayText -Rule $rule

            $item = New-Object System.Windows.Forms.ListViewItem([string]$rule.name)
            [void]$item.SubItems.Add($triggerText)
            $workbookSubItem = $item.SubItems.Add($workbookText)

            # Whatever has already been read about these workbooks. Nothing is
            # opened here: the background metadata worker does that, so
            # redrawing the grid never touches a slow network share.
            $shareInfo = Get-RuleWorkbookInfoSummary -Rule $rule
            [void]$item.SubItems.Add([string]$shareInfo.LocationText)

            $statusInfo = Get-RuleStatusText -Status $watcherStatus
            [void]$item.SubItems.Add([string]$statusInfo.Text)

            $queryScopeText = Get-RuleQueryScopeText -Rule $rule
            [void]$item.SubItems.Add($queryScopeText)

            $lastRun = Get-RuleLastRunInfo -Rule $rule -State $state
            $lastWhen = Get-RuleLastRunTime -Rule $rule -State $state
            $lastRunCell = [string]$lastRun.Time
            $lastRunExactText = ''
            if ($lastWhen -gt [DateTime]::MinValue) {
                $lastRunCell = Get-RelativeTimeText -When $lastWhen
                $lastRunExactText = 'Last run: {0}' -f $lastWhen.ToString('dddd d MMMM, HH:mm:ss')
            }
            [void]$item.SubItems.Add($lastRunCell)

            $nextRunText = Get-RuleNextRunText -Rule $rule -State $state
            $nextRunWhen = Get-RuleNextRunTime -Rule $rule -State $state
            [void]$item.SubItems.Add($nextRunText)

            $durationItem = $item.SubItems.Add([string]$lastRun.Duration)
            [void]$item.SubItems.Add([string]$shareInfo.DataAgeText)

            $lastSeconds = ConvertTo-DurationSeconds ([string]$lastRun.Duration)
            $sortKeys = @{
                0 = ([string]$rule.name).ToLowerInvariant()
                1 = ([string]$triggerText).ToLowerInvariant()
                2 = ([string]$workbookText).ToLowerInvariant()
                3 = ([string]$shareInfo.LocationText).ToLowerInvariant()
                4 = ([string]$statusInfo.Text).ToLowerInvariant()
                5 = ([string]$queryScopeText).ToLowerInvariant()
                6 = [int64]$lastWhen.Ticks
                7 = [int64]$nextRunWhen.Ticks
                8 = [double]$lastSeconds
                9 = [int64]$shareInfo.DataSortTicks
            }
            $item.Tag  = @{ Kind = 'Rule'; RuleId = $ruleId; FieldTooltips = @{}
                            SortKeys = $sortKeys; OriginalIndex = $originalIndex }
            $item.Font = $script:UiFonts.Bold
            # Off, so the Duration cell can be coloured on its own. The running
            # and queued shading therefore has to paint every cell explicitly.
            $item.UseItemStyleForSubItems = $false

            $actionCount = @($rule.actions).Count
            if ($actionCount -gt 1) {
                $workbookSubItem.Font = Get-UiFont 9 'Bold'
                $workbookSubItem.ForeColor = [System.Drawing.Color]::FromArgb(35, 90, 165)
            }

            # A run that took far longer than this rule usually does is worth
            # noticing before it becomes a hang.
            $baseline = Get-RuleDurationBaseline -Rule $rule
            $slowDurationText = ''
            if ($baseline -gt 0 -and $lastSeconds -gt 20 -and $lastSeconds -ge ($baseline * 2)) {
                $durationItem.ForeColor = [System.Drawing.Color]::FromArgb(170, 110, 10)
                $slowDurationText = 'This rule usually takes about {0:0}s.' -f $baseline
            }

            # ---- column-specific hover details ----------------------------
            # Indexes match the ten columns created in New-DashboardForm.
            $fieldTips = $item.Tag.FieldTooltips

            $fieldTips[0] = ('Rule: {0}{1}Drag this row to change its saved position.{1}Enabled: {2}{1}Ask before automatic refresh: {3}' -f `
                [string]$rule.name, [Environment]::NewLine,
                $(if ($enabled) { 'Yes' } else { 'No' }),
                $(if (Get-RuleAskBeforeRefreshValue -Rule $rule) { 'Yes' } else { 'No' }))

            $fieldTips[1] = Get-RuleTriggerFieldTooltip -Rule $rule

            $workbookPaths = @($rule.actions | ForEach-Object { [string]$_.path } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $workbookTooltipLines = New-Object System.Collections.ArrayList
            if ($workbookPaths.Count -eq 0) {
                [void]$workbookTooltipLines.Add('No workbook is configured.')
            }
            else {
                for ($workbookIndex = 0; $workbookIndex -lt $workbookPaths.Count; $workbookIndex++) {
                    $workbookName = Split-Path -Leaf $workbookPaths[$workbookIndex]
                    [void]$workbookTooltipLines.Add(('{0}. {1}' -f ($workbookIndex + 1), $workbookName))
                    [void]$workbookTooltipLines.Add(('   {0}' -f $workbookPaths[$workbookIndex]))
                }
            }
            $fieldTips[2] = @($workbookTooltipLines.ToArray()) -join [Environment]::NewLine

            $fieldTips[3] = [string]$shareInfo.LocationTooltip
            $fieldTips[4] = [string]$statusInfo.Explanation
            $fieldTips[5] = Get-RuleQueryFieldTooltip -Rule $rule

            $lastRunStatus = [string]$lastRun.Status
            $lastRunLines = New-Object System.Collections.ArrayList
            if (-not [string]::IsNullOrWhiteSpace($lastRunExactText)) {
                [void]$lastRunLines.Add($lastRunExactText)
            }
            else {
                [void]$lastRunLines.Add('This rule has not recorded a completed run yet.')
            }
            if (-not [string]::IsNullOrWhiteSpace($lastRunStatus)) {
                [void]$lastRunLines.Add(('Result: {0}' -f $lastRunStatus))
            }
            if ($lastRunStatus -eq 'Failed') {
                $failure = Get-RuleFailureExplanation -Rule $rule
                if ($null -ne $failure) { [void]$lastRunLines.Add([string]$failure.Sentence) }
            }
            $fieldTips[6] = @($lastRunLines.ToArray()) -join [Environment]::NewLine

            if (Test-TriggerIsScheduled ([string]$rule.trigger.type)) {
                $nextLines = New-Object System.Collections.ArrayList
                if (-not [string]::IsNullOrWhiteSpace($nextRunText)) {
                    [void]$nextLines.Add(('Next run: {0}' -f $nextRunText))
                }
                [void]$nextLines.Add(('Schedule: {0} on {1}' -f `
                    [string]$rule.trigger.scheduleTime,
                    (Format-ScheduleDays @($rule.trigger.scheduleDays))))
                $fieldTips[7] = @($nextLines.ToArray()) -join [Environment]::NewLine
            }
            else {
                $fieldTips[7] = ''
            }

            $durationLines = New-Object System.Collections.ArrayList
            if (-not [string]::IsNullOrWhiteSpace([string]$lastRun.Duration) -and [string]$lastRun.Duration -ne '-') {
                [void]$durationLines.Add(('Last duration: {0}' -f [string]$lastRun.Duration))
            }
            else {
                [void]$durationLines.Add('No completed duration is recorded yet.')
            }
            if (-not [string]::IsNullOrWhiteSpace($slowDurationText)) {
                [void]$durationLines.Add($slowDurationText)
            }
            elseif ($baseline -gt 0) {
                [void]$durationLines.Add(('Typical recent duration: about {0:0}s.' -f $baseline))
            }
            $fieldTips[8] = @($durationLines.ToArray()) -join [Environment]::NewLine

            $fieldTips[9] = [string]$shareInfo.DataTooltip

            # Native item tooltips are disabled; all Trigger Rules hover text
            # now comes from FieldTooltips for the field under the pointer.
            $item.ToolTipText = ''

            $rowColor = $null
            if (-not $enabled) {
                $rowColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
            }
            elseif ($watcherStatus -eq 'PathUnavailable' -or $watcherStatus -eq 'Error' -or $watcherStatus -eq 'Misconfigured') {
                $rowColor = [System.Drawing.Color]::FromArgb(180, 60, 40)
            }
            if ($null -ne $rowColor) {
                $item.ForeColor = $rowColor
                foreach ($sub in $item.SubItems) {
                    if (-not [object]::ReferenceEquals($sub, $durationItem) -or $slowDurationText -eq '') { $sub.ForeColor = $rowColor }
                }
            }
            [void]$renderedItems.Add($item)
            $originalIndex++
        }

        $itemsToAdd = @($renderedItems.ToArray())
        $sortColumn = [int]$script:UiRuleSortColumn
        if ($sortColumn -ge 0 -and $sortColumn -lt $script:UiRuleColumnNames.Count) {
            $descending = ($script:UiRuleSortDirection -eq 'Descending')
            $primarySort = @{ Expression = { $_.Tag.SortKeys[$sortColumn] }; Descending = $descending }
            $stableSort  = @{ Expression = { [int]$_.Tag.OriginalIndex }; Ascending = $true }
            $itemsToAdd = @($itemsToAdd | Sort-Object -Property $primarySort, $stableSort)
        }

        $displayedIds = New-Object System.Collections.ArrayList
        foreach ($item in $itemsToAdd) {
            [void]$listView.Items.Add($item)
            [void]$displayedIds.Add([string]$item.Tag.RuleId)
            if ((Get-RuleListRowKey $item.Tag) -eq $selectedKey) { $item.Selected = $true }
        }
        $script:UiDisplayedRuleIds = @($displayedIds.ToArray())
    }
    finally {
        $listView.EndUpdate()
    }
    # Guarded: Update-RuleList is reachable from a few places, and the panel
    # only exists once the dashboard has been built.
    $emptyPanel = $script:UiControls.RulesEmpty
    if ($null -ne $emptyPanel) {
        $isEmpty = ($listView.Items.Count -eq 0)
        if ($emptyPanel.Visible -ne $isEmpty) {
            $emptyPanel.Visible = $isEmpty
            if ($isEmpty) { $emptyPanel.BringToFront() }
        }
    }

    $script:UiHighlightSignature = ''
    Update-RunningRuleHighlight
    Update-ActivityRuleChoices
    Update-SelectedRuleAskBeforeCheck
    # In Show all mode a newly added or removed rule changes the ideal height.
    Set-RulesPaneMode -Mode ([string]$script:UiConfig.appSettings.rulesPaneMode)
}

function Update-RunningRuleHighlight {
    <#
        Shades the rule that is refreshing right now, and more faintly the rules
        that are waiting behind it. Runs on every tick because the rule list
        itself is only rebuilt every few seconds, which is far too slow to show
        a job starting.
    #>
    $listView = $script:UiControls.Rules
    if ($null -eq $listView -or $listView.Items.Count -eq 0) { return }

    $current   = $script:UiShared.CurrentJob
    $runningId = $(if ([bool]$current.Active) { [string]$current.RuleId } else { '' })

    $queuedIds = @{}
    foreach ($job in @(Get-PendingJobSnapshot -Shared $script:UiShared)) {
        $queuedIds[[string]$job.RuleId] = $true
    }

    $signature = $runningId + '#' + ((@($queuedIds.Keys) | Sort-Object) -join ',')
    if ($signature -eq $script:UiHighlightSignature) { return }
    $script:UiHighlightSignature = $signature

    $runningBack = [System.Drawing.Color]::FromArgb(206, 238, 210)
    $queuedBack  = [System.Drawing.Color]::FromArgb(252, 240, 205)
    $normalBack  = $listView.BackColor

    $listView.BeginUpdate()
    try {
        foreach ($item in $listView.Items) {
            $tag = $item.Tag
            if ($null -eq $tag) { continue }
            $ruleId = [string]$tag.RuleId
            $back = $normalBack
            if     (-not [string]::IsNullOrWhiteSpace($runningId) -and $ruleId -eq $runningId) { $back = $runningBack }
            elseif ($queuedIds.ContainsKey($ruleId))                                          { $back = $queuedBack }
            # UseItemStyleForSubItems is off for these rows, so each cell needs
            # painting or only the first column would change.
            $item.BackColor = $back
            foreach ($sub in $item.SubItems) { $sub.BackColor = $back }
        }
    }
    finally {
        $listView.EndUpdate()
    }
}

function Update-RunNowCaption {
    <#
        While a refresh is running, Run Now adds to the queue instead of
        starting immediately, so the button says what it will actually do.
    #>
    $isQueue = [bool]$script:UiShared.CurrentJob.Active
    if ($script:UiRunNowIsQueue -eq $isQueue) { return }
    $script:UiRunNowIsQueue = $isQueue

    $button = $script:UiControls.BtnRunNow
    $button.Text = $(if ($isQueue) { 'Add to Queue' } else { 'Run Now' })
    $script:UiControls.MenuRunWholeRule.Text = $(if ($isQueue) { 'Add whole rule to queue' } else { 'Run whole rule now' })
}

function Update-ActivityDetail {
    $listView = $script:UiControls.Activity
    if ($listView.SelectedItems.Count -eq 0) { return }

    $entry = $listView.SelectedItems[0].Tag
    if ($null -eq $entry) { return }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add(('Time       : {0}' -f ([DateTime]$entry.Time).ToString('yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add(('Level      : {0}' -f $entry.Level))
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.RuleName))  { [void]$lines.Add(('Rule       : {0}' -f $entry.RuleName)) }
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.Workbook))  { [void]$lines.Add(('Workbook   : {0}' -f $entry.Workbook)) }
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.Stage))     { [void]$lines.Add(('Stage      : {0}' -f $entry.Stage)) }
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.ErrorType)) { [void]$lines.Add(('Error type : {0}' -f $entry.ErrorType)) }
    [void]$lines.Add(('Message    : {0}' -f $entry.Message))

    # Attach the most recent finished job for the same rule, when there is one.
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.RuleName)) {
        $record = $null
        foreach ($candidate in (Copy-SharedList $script:UiShared.History)) {
            if ([string]$candidate.rule -eq [string]$entry.RuleName) { $record = $candidate }
        }
        if ($null -ne $record) {
            [void]$lines.Add('')
            [void]$lines.Add(('Last job   : {0}  status={1}  elapsed={2}' -f $record.jobId, $record.status, $record.elapsed))
            [void]$lines.Add(('  trigger  : {0} / {1} at {2}' -f $record.triggerType, $record.triggerSource, $record.triggeredAt))
            if (-not [string]::IsNullOrWhiteSpace([string]$record.triggerFile)) {
                [void]$lines.Add(('  source   : {0}' -f $record.triggerFile))
            }
            [void]$lines.Add(('  started  : {0}   finished: {1}' -f $record.startedAt, $record.finishedAt))
            foreach ($workbookResult in @($record.results)) {
                [void]$lines.Add(('  workbook : {0} -> {1} {2} ({3}s)' -f (Split-Path -Leaf ([string]$workbookResult.workbook)),
                    $(if ($workbookResult.success) { 'OK' } else { 'FAILED' }),
                    [string]$workbookResult.errorType, [string]$workbookResult.elapsed))
                if (@($workbookResult.queries).Count -gt 0) {
                    [void]$lines.Add(('    queries  : {0}' -f (@($workbookResult.queries) -join ', ')))
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$workbookResult.message)) {
                    [void]$lines.Add(('             {0}' -f $workbookResult.message))
                }
            }
        }
    }

    $script:UiControls.Detail.Text = ($lines -join [Environment]::NewLine)
}

function Update-CurrentJobDisplay {
    $current = $script:UiShared.CurrentJob

    if (-not $current.Active) {
        if ($script:UiControls.JobRule.Text -ne '-') {
            $script:UiControls.JobRule.Text     = '-'
            $script:UiControls.JobWorkbook.Text = '-'
            $script:UiControls.JobStage.Text    = 'Idle'
            $script:UiControls.JobQuery.Text    = '-'
            $script:UiControls.JobElapsed.Text  = '00:00:00'
        }
        if ($script:UiControls.ContainsKey('QueryProgressToolTip')) {
            $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobQuery, '')
            $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobElapsed, '')
        }
        $script:UiControls.Progress.Visible = $false
        if ($script:UiControls.BtnCancelJob.Enabled) { $script:UiControls.BtnCancelJob.Enabled = $false }
        return
    }

    $script:UiControls.JobRule.Text     = [string]$current.RuleName
    $script:UiControls.JobWorkbook.Text = [string]$current.Workbook

    $stageText = Get-StageDisplayText ([string]$current.Stage)
    $elapsedText = '00:00:00'
    if ($null -ne $current.StartedAt) {
        $elapsedText = Format-Elapsed ((Get-Date) - [DateTime]$current.StartedAt)
    }

    # Excel exposes individual connection state only while it is able to do so.
    # When available, enrich the ordinary stage/elapsed fields without adding
    # another fixed-height row to the compact Dashboard.
    $queryTotal = 0
    $queryCompleted = 0
    $queryPosition = 0
    $queryName = ''
    $queryElapsed = 0
    $queryDetail = ''
    try {
        if ($current.ContainsKey('QueryTotal'))          { $queryTotal = [int]$current.QueryTotal }
        if ($current.ContainsKey('QueryCompleted'))      { $queryCompleted = [int]$current.QueryCompleted }
        if ($current.ContainsKey('QueryPosition'))       { $queryPosition = [int]$current.QueryPosition }
        if ($current.ContainsKey('QueryName'))           { $queryName = [string]$current.QueryName }
        if ($current.ContainsKey('QueryElapsedSeconds')) { $queryElapsed = [int]$current.QueryElapsedSeconds }
        if ($current.ContainsKey('QueryProgressDetail')) { $queryDetail = [string]$current.QueryProgressDetail }
    }
    catch { }

    $queryText = '-'
    if ($queryTotal -gt 0 -and [string]$current.Stage -in @('Refreshing', 'RefreshingLong')) {
        if ($queryPosition -le 0) {
            $queryPosition = [Math]::Min($queryTotal, [Math]::Max(1, $queryCompleted + 1))
        }

        if (-not [string]::IsNullOrWhiteSpace($queryName)) {
            $queryText = '{0}  ({1}/{2})' -f $queryName, $queryPosition, $queryTotal
        }
        else {
            $queryText = '{0}/{1} complete' -f $queryCompleted, $queryTotal
        }

        if ($queryElapsed -gt 0) {
            $elapsedText += '  |  Query ' + (Format-Elapsed ([TimeSpan]::FromSeconds($queryElapsed)))
        }

        if ($script:UiControls.ContainsKey('QueryProgressToolTip')) {
            $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobQuery, $queryDetail)
            $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobElapsed, $queryDetail)
        }
    }
    elseif ($script:UiControls.ContainsKey('QueryProgressToolTip')) {
        $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobQuery, '')
        $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobElapsed, '')
    }

    if ([string]$current.Stage -eq 'Saving') {
        $saveStartedAt = $null
        $saveMode = ''
        $saveDetail = ''
        try {
            if ($current.ContainsKey('SaveStartedAt'))      { $saveStartedAt = $current.SaveStartedAt }
            if ($current.ContainsKey('SaveMode'))           { $saveMode = [string]$current.SaveMode }
            if ($current.ContainsKey('SaveProgressDetail')) { $saveDetail = [string]$current.SaveProgressDetail }
        }
        catch { }
        if ($null -ne $saveStartedAt) {
            $elapsedText += '  |  Save ' + (Format-Elapsed ((Get-Date) - [DateTime]$saveStartedAt))
        }
        $queryText = 'Refresh complete'
        if ($script:UiControls.ContainsKey('QueryProgressToolTip')) {
            $saveTip = $(if (-not [string]::IsNullOrWhiteSpace($saveDetail)) { $saveDetail } else { $saveMode })
            $script:UiControls.QueryProgressToolTip.SetToolTip($script:UiControls.JobElapsed, $saveTip)
        }
    }

    $script:UiControls.JobStage.Text   = $stageText
    $script:UiControls.JobQuery.Text   = $queryText
    $script:UiControls.JobElapsed.Text = $elapsedText

    if (-not $script:UiControls.Progress.Visible) {
        $script:UiControls.Progress.Visible = $true
    }
    $cancelEnabled = (ConvertTo-BoolValue $current.CanCancel $false) -and ($null -eq $script:UiCancelRequestedAt)
    if ($script:UiControls.BtnCancelJob.Enabled -ne $cancelEnabled) { $script:UiControls.BtnCancelJob.Enabled = $cancelEnabled }
    if ($script:UiControls.Progress.Style -ne 'Marquee') {
        $script:UiControls.Progress.Style = 'Marquee'
    }
}

function New-ActivityListRow {
    param([Parameter(Mandatory = $true)]$Entry)
    $item = New-Object System.Windows.Forms.ListViewItem(([DateTime]$Entry.Time).ToString('HH:mm:ss'))
    [void]$item.SubItems.Add([string]$Entry.Level)
    [void]$item.SubItems.Add([string]$Entry.RuleName)
    [void]$item.SubItems.Add([string]$Entry.Message)
    $item.Tag = $Entry
    switch ([string]$Entry.Level) {
        'ERROR'   { $item.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 35) }
        'WARN'    { $item.ForeColor = [System.Drawing.Color]::FromArgb(170, 110, 10) }
        'SUCCESS' { $item.ForeColor = [System.Drawing.Color]::FromArgb(25, 115, 50) }
        'DEBUG'   { $item.ForeColor = [System.Drawing.Color]::FromArgb(130, 130, 130) }
    }
    return $item
}

function Test-ActivityRowVisible {
    <#  The three filter controls, applied to one entry.  #>
    param([Parameter(Mandatory = $true)]$Entry)

    $level = [string]$Entry.Level
    switch ([int]$script:UiControls.ActivityLevel.SelectedIndex) {
        1 { if ($level -eq 'DEBUG') { return $false } }
        2 { if (@('WARN', 'ERROR') -notcontains $level) { return $false } }
        3 { if ($level -ne 'ERROR') { return $false } }
    }

    $rule = [string]$script:UiControls.ActivityRule.SelectedItem
    if (-not [string]::IsNullOrWhiteSpace($rule) -and $rule -ne 'Any rule') {
        if ([string]$Entry.RuleName -ne $rule) { return $false }
    }

    $find = [string]$script:UiControls.ActivityFind.Text
    if (-not [string]::IsNullOrWhiteSpace($find)) {
        $haystack = ('{0} {1} {2}' -f [string]$Entry.RuleName, [string]$Entry.Workbook, [string]$Entry.Message)
        if ($haystack.IndexOf($find, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false }
    }

    return $true
}

function Update-ActivityFeed {
    <#
        Entries are kept in a buffer so a filter change can redraw the list
        without waiting for new activity to arrive.
    #>
    $listView = $script:UiControls.Activity
    $newEntries = @(Copy-SharedList $script:UiShared.Activity | Where-Object { [int]$_.Seq -gt $script:UiActivitySeq })
    if ($newEntries.Count -eq 0) { return }

    $listView.BeginUpdate()
    try {
        foreach ($entry in $newEntries) {
            $script:UiActivitySeq = [int]$entry.Seq
            [void]$script:UiActivityBuffer.Add($entry)
            if (Test-ActivityRowVisible -Entry $entry) { [void]$listView.Items.Add((New-ActivityListRow -Entry $entry)) }

        }

        while ($script:UiActivityBuffer.Count -gt 300) { $script:UiActivityBuffer.RemoveAt(0) }
        while ($listView.Items.Count -gt 300) { $listView.Items.RemoveAt(0) }
        if ($listView.SelectedItems.Count -eq 0 -and $listView.Items.Count -gt 0) {
            $listView.EnsureVisible($listView.Items.Count - 1)
        }
    }
    finally {
        $listView.EndUpdate()
    }
}

function Update-ActivityFilter {
    <#  Redraws the visible rows from the buffer after a filter change.  #>
    $listView = $script:UiControls.Activity
    $listView.BeginUpdate()
    try {
        $listView.Items.Clear()
        foreach ($entry in @($script:UiActivityBuffer.ToArray())) {
            if (Test-ActivityRowVisible -Entry $entry) { [void]$listView.Items.Add((New-ActivityListRow -Entry $entry)) }
        }
        if ($listView.Items.Count -gt 0) { $listView.EnsureVisible($listView.Items.Count - 1) }
    }
    finally {
        $listView.EndUpdate()
    }
    $script:UiControls.Detail.Text = ''
}

function Reset-ActivityFilter {
    $script:UiControls.ActivityLevel.SelectedIndex = 0
    $script:UiControls.ActivityRule.SelectedIndex  = 0
    $script:UiControls.ActivityFind.Text           = ''
    Update-ActivityFilter
}

function Update-ActivityRuleChoices {
    <#  Keeps the rule filter in step with the configured rules.  #>
    $combo = $script:UiControls.ActivityRule
    $names = @(@($script:UiConfig.rules) | ForEach-Object { [string]$_.name } | Sort-Object)
    $signature = ($names -join '|')
    if ($signature -eq $script:UiActivityRuleChoices) { return }
    $script:UiActivityRuleChoices = $signature

    $selected = [string]$combo.SelectedItem
    $combo.BeginUpdate()
    try {
        $combo.Items.Clear()
        [void]$combo.Items.Add('Any rule')
        foreach ($name in $names) { [void]$combo.Items.Add($name) }
        $restore = $combo.Items.IndexOf($selected)
        $combo.SelectedIndex = $(if ($restore -ge 0) { $restore } else { 0 })
    }
    finally {
        $combo.EndUpdate()
    }
}

function Invoke-ActivityRowActivate {
    <#  Double-clicking a row jumps to the rule it is about.  #>
    $listView = $script:UiControls.Activity
    if ($listView.SelectedItems.Count -eq 0) { return }
    $entry = $listView.SelectedItems[0].Tag
    if ($null -eq $entry) { return }

    $ruleName = [string]$entry.RuleName
    if ([string]::IsNullOrWhiteSpace($ruleName)) { return }

    $target = $null
    foreach ($rule in @($script:UiConfig.rules)) {
        if ([string]$rule.name -eq $ruleName) { $target = $rule; break }
    }
    if ($null -eq $target) { return }

    $rules = $script:UiControls.Rules
    foreach ($item in $rules.Items) {
        $tag = $item.Tag
        if ($null -ne $tag -and [string]$tag.RuleId -eq [string]$target.id) {
            $item.Selected = $true
            $rules.EnsureVisible($item.Index)
            $rules.Focus() | Out-Null
            break
        }
    }
    Invoke-RuleEdit
}

function Get-PendingJobDisplayText {
    param([Parameter(Mandatory = $true)]$Job, [int]$Position)
    $workbooks = @(@($Job.Actions) | ForEach-Object { [string]$_.path } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $what = $(if ($workbooks.Count -eq 1) { Split-Path -Leaf ([string]$workbooks[0]) } else { '{0} workbooks' -f $workbooks.Count })
    return ('{0}. {1}  -  {2}' -f $Position, [string]$Job.RuleName, $what)
}

function Update-PendingList {
    <#
        Rows are whole jobs now rather than individual workbooks, because a row
        is something you can move or remove and a job is the unit that queues.
    #>
    $listBox = $script:UiControls.Pending
    $jobs    = @(Get-PendingJobSnapshot -Shared $script:UiShared)

    $texts = New-Object System.Collections.ArrayList
    $ids   = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $jobs.Count; $index++) {
        [void]$texts.Add((Get-PendingJobDisplayText -Job $jobs[$index] -Position ($index + 1)))
        [void]$ids.Add([string]$jobs[$index].Id)
    }

    $signature = ($ids.ToArray() -join '|') + '#' + ($texts.ToArray() -join '|')
    if ($signature -eq $script:UiPendingSignature) { return }
    $script:UiPendingSignature = $signature
    $script:UiPendingJobIds    = @($ids.ToArray())

    $selectedId = ''
    if ($listBox.SelectedIndex -ge 0 -and $listBox.SelectedIndex -lt @($script:UiPendingJobIds).Count) {
        $selectedId = [string]$script:UiPendingJobIds[$listBox.SelectedIndex]
    }

    $listBox.BeginUpdate()
    try {
        $listBox.Items.Clear()
        foreach ($text in @($texts.ToArray())) { [void]$listBox.Items.Add($text) }
        if (-not [string]::IsNullOrWhiteSpace($selectedId)) {
            $restore = [array]::IndexOf(@($script:UiPendingJobIds), $selectedId)
            if ($restore -ge 0) { $listBox.SelectedIndex = $restore }
        }
    }
    finally {
        $listBox.EndUpdate()
    }
    Update-PendingButtons
}

function Update-PendingButtons {
    $listBox = $script:UiControls.Pending
    $index   = $listBox.SelectedIndex
    $count   = @($script:UiPendingJobIds).Count
    $hasRow  = ($index -ge 0 -and $index -lt $count)
    $script:UiControls.BtnPendingUp.Enabled     = ($hasRow -and $index -gt 0)
    $script:UiControls.BtnPendingDown.Enabled   = ($hasRow -and $index -lt ($count - 1))
    $script:UiControls.BtnPendingRemove.Enabled = $hasRow
}

function Get-SelectedPendingJobId {
    $index = $script:UiControls.Pending.SelectedIndex
    if ($index -lt 0 -or $index -ge @($script:UiPendingJobIds).Count) { return '' }
    return [string]$script:UiPendingJobIds[$index]
}

function Invoke-PendingJobMove {
    param([Parameter(Mandatory = $true)][int]$Delta)
    $jobId = Get-SelectedPendingJobId
    if ([string]::IsNullOrWhiteSpace($jobId)) { return }
    if (Move-PendingJob -Shared $script:UiShared -JobId $jobId -Delta $Delta) {
        Write-AppLog -Level 'INFO' -Message ('{0} was moved {1} in the queue.' -f $jobId, $(if ($Delta -lt 0) { 'up' } else { 'down' }))
        Update-PendingList
        $moved = [array]::IndexOf(@($script:UiPendingJobIds), $jobId)
        if ($moved -ge 0) { $script:UiControls.Pending.SelectedIndex = $moved }
    }
}

function Invoke-PendingJobRemove {
    $jobId = Get-SelectedPendingJobId
    if ([string]::IsNullOrWhiteSpace($jobId)) { return }

    $index = $script:UiControls.Pending.SelectedIndex
    $label = [string]$script:UiControls.Pending.Items[$index]
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ('Take this job out of the queue?' + [Environment]::NewLine + [Environment]::NewLine + $label +
         [Environment]::NewLine + [Environment]::NewLine + 'Nothing is refreshed and no file is changed. The rule itself stays as it is.'),
        'Remove from queue', 'YesNo', 'Question')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $removed = Remove-PendingJob -Shared $script:UiShared -JobId $jobId
    if ($null -eq $removed) {
        # The engine took it in the moment between drawing and clicking.
        Write-AppLog -Level 'INFO' -Message ('{0} had already started, so it was not removed from the queue.' -f $jobId)
    }
    else {
        Write-AppLog -Level 'WARN' -RuleName ([string]$removed.RuleName) `
            -Message ('{0} was removed from the queue before it started.' -f $jobId)
    }
    Update-PendingList
}

function Update-Notifications {
    $shared = $script:UiShared
    while ($shared.Notifications.Count -gt 0) {
        $notification = $null
        try { $notification = $shared.Notifications.Dequeue() } catch { break }
        if ($null -eq $notification) { continue }

        $iconType = [System.Windows.Forms.ToolTipIcon]::Info
        if ([string]$notification.Level -eq 'Error')   { $iconType = [System.Windows.Forms.ToolTipIcon]::Error }
        if ([string]$notification.Level -eq 'Warning') { $iconType = [System.Windows.Forms.ToolTipIcon]::Warning }

        try {
            $script:UiControls.Tray.ShowBalloonTip(5000, [string]$notification.Title, [string]$notification.Text, $iconType)
        }
        catch { }
    }
}

function Show-ManualRefreshResultDialog {
    param([Parameter(Mandatory = $true)]$Record)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Manual Excel Refresh Result'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize = New-Object System.Drawing.Size(650,440)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterParent'
    $form.MaximizeBox=$false; $form.MinimizeBox=$false
    $form.Font = Get-UiFont

    $statusText = [string]$Record.status
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location=New-Object System.Drawing.Point(16,14); $lbl.AutoSize=$true
    $lbl.Font=Get-UiFont 11 'Bold'; $lbl.Text=$(if($statusText -eq 'Completed'){'Refresh completed'}else{'Refresh finished: '+$statusText})
    $lbl.ForeColor=$(if($statusText -eq 'Completed'){[System.Drawing.Color]::FromArgb(30,120,55)}else{[System.Drawing.Color]::FromArgb(180,60,40)})
    $form.Controls.Add($lbl)

    $txt=New-Object System.Windows.Forms.TextBox
    $txt.Location=New-Object System.Drawing.Point(16,52); $txt.Size=New-Object System.Drawing.Size(618,330)
    $txt.Multiline=$true; $txt.ReadOnly=$true; $txt.ScrollBars='Vertical'; $txt.WordWrap=$true
    $lines=New-Object System.Collections.ArrayList
    [void]$lines.Add(('Job:       {0}' -f $Record.jobId))
    [void]$lines.Add(('Status:    {0}' -f $Record.status))
    [void]$lines.Add(('Started:   {0}' -f $Record.startedAt))
    [void]$lines.Add(('Finished:  {0}' -f $Record.finishedAt))
    [void]$lines.Add(('Elapsed:   {0}' -f $Record.elapsed))
    [void]$lines.Add('')
    foreach($r in @($Record.results)){
        [void]$lines.Add(('Workbook:   {0}' -f [string]$r.workbook))
        [void]$lines.Add(('Result:     {0}' -f $(if($r.success){'SUCCESS'}else{'FAILED'})))
        [void]$lines.Add(('Saved:      {0}' -f $(if($r.saved){'Yes'}else{'No'})))
        [void]$lines.Add(('Queries:    {0}' -f $(if(@($r.queries).Count -gt 0){@($r.queries)-join ', '}else{'No Power Query names reported'})))
        if(-not [string]::IsNullOrWhiteSpace([string]$r.errorType)){[void]$lines.Add(('Error type: {0}' -f $r.errorType))}
        if(-not [string]::IsNullOrWhiteSpace([string]$r.message)){[void]$lines.Add(('Message:    {0}' -f $r.message))}
        [void]$lines.Add('')
    }
    [void]$lines.Add('This result is also stored in the application history and log files.')
    $txt.Text=$lines -join [Environment]::NewLine
    $form.Controls.Add($txt)

    $btnLog=New-Object System.Windows.Forms.Button; $btnLog.Text='Open Log'; $btnLog.Location=New-Object System.Drawing.Point(444,396); $btnLog.Size=New-Object System.Drawing.Size(90,30); $form.Controls.Add($btnLog)
    $btnClose=New-Object System.Windows.Forms.Button; $btnClose.Text='Close'; $btnClose.Location=New-Object System.Drawing.Point(544,396); $btnClose.Size=New-Object System.Drawing.Size(90,30); $btnClose.DialogResult='OK'; $form.Controls.Add($btnClose); $form.AcceptButton=$btnClose
    $btnLog.Add_Click({ Invoke-OpenLog }.GetNewClosure())
    Set-FormWithinWorkingArea -Form $form
    if ($script:UiControls.Form.Visible) { [void]$form.ShowDialog($script:UiControls.Form) }
    else { [void]$form.ShowDialog() }
}

function Update-ManualRefreshResultPopups {
    $history=@(Copy-SharedList $script:UiShared.History)
    if(-not $script:UiManualHistoryInitialized){
        foreach($record in $history){ if([string]$record.ruleId -eq 'MANUAL-REFRESH'){ $script:UiSeenManualHistory[[string]$record.jobId]=$true } }
        $script:UiManualHistoryInitialized=$true
        return
    }
    foreach($record in $history){
        if([string]$record.ruleId -ne 'MANUAL-REFRESH'){continue}
        $jobId=[string]$record.jobId
        if([string]::IsNullOrWhiteSpace($jobId) -or $script:UiSeenManualHistory.ContainsKey($jobId)){continue}
        $script:UiSeenManualHistory[$jobId]=$true
        Show-ManualRefreshResultDialog -Record $record
    }
}


function Get-TriggeredCauseText {
    param([Parameter(Mandatory = $true)]$Record)

    $type = [string]$Record.triggerType
    $source = [string]$Record.triggerSource
    $file = [string]$Record.triggerFile
    $leaf = ''
    if (-not [string]::IsNullOrWhiteSpace($file)) {
        try { $leaf = Split-Path -Leaf $file } catch { $leaf = $file }
    }

    switch ($source) {
        'File' {
            switch ($type) {
                'FileCreated' {
                    if (-not [string]::IsNullOrWhiteSpace($leaf)) { return ('A new matching file was added: {0}' -f $leaf) }
                    return 'A new matching file was added to the monitored folder.'
                }
                'FileChangedSpecific' {
                    if (-not [string]::IsNullOrWhiteSpace($leaf)) { return ('The monitored file was updated: {0}' -f $leaf) }
                    return 'The monitored file was updated.'
                }
                'FileChangedAny' {
                    if (-not [string]::IsNullOrWhiteSpace($leaf)) { return ('A matching monitored file was updated: {0}' -f $leaf) }
                    return 'A matching file in the monitored folder was updated.'
                }
                default {
                    if (-not [string]::IsNullOrWhiteSpace($leaf)) { return ('A monitored file event occurred: {0}' -f $leaf) }
                    return 'A monitored file event occurred.'
                }
            }
        }
        'Schedule' { return ('The scheduled time for rule "{0}" was reached.' -f [string]$Record.rule) }
        'Logon'    { return ('The Windows logon trigger for rule "{0}" ran.' -f [string]$Record.rule) }
        default    { return ('Rule "{0}" was triggered automatically.' -f [string]$Record.rule) }
    }
}

function Show-TriggeredRefreshResultDialog {
    param([Parameter(Mandatory = $true)]$Record)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Triggered Refresh Result'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize = New-Object System.Drawing.Size(680,500)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterScreen'
    $form.MaximizeBox=$false; $form.MinimizeBox=$false
    $form.Font = Get-UiFont
    $form.TopMost = $true

    $statusText = [string]$Record.status
    $success = ($statusText -eq 'Completed')
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location=New-Object System.Drawing.Point(18,14); $lbl.AutoSize=$true
    $lbl.Font=Get-UiFont 11 'Bold'
    $lbl.Text=$(if($success){'Triggered refresh completed'}else{'Triggered refresh finished: '+$statusText})
    $lbl.ForeColor=$(if($success){[System.Drawing.Color]::FromArgb(30,120,55)}else{[System.Drawing.Color]::FromArgb(180,60,40)})
    $form.Controls.Add($lbl)

    $cause = Get-TriggeredCauseText -Record $Record
    $causeLabel = New-Object System.Windows.Forms.Label
    $causeLabel.Location=New-Object System.Drawing.Point(18,46)
    $causeLabel.Size=New-Object System.Drawing.Size(644,48)
    $causeLabel.Text=('Why this ran:  {0}' -f $cause)
    $form.Controls.Add($causeLabel)

    $txt=New-Object System.Windows.Forms.TextBox
    $txt.Location=New-Object System.Drawing.Point(18,96); $txt.Size=New-Object System.Drawing.Size(644,344)
    $txt.Multiline=$true; $txt.ReadOnly=$true; $txt.ScrollBars='Vertical'; $txt.WordWrap=$true
    $lines=New-Object System.Collections.ArrayList
    [void]$lines.Add(('Rule:       {0}' -f [string]$Record.rule))
    [void]$lines.Add(('Triggered:  {0}' -f [string]$Record.triggeredAt))
    [void]$lines.Add(('Started:    {0}' -f [string]$Record.startedAt))
    [void]$lines.Add(('Finished:   {0}' -f [string]$Record.finishedAt))
    [void]$lines.Add(('Elapsed:    {0}' -f [string]$Record.elapsed))
    [void]$lines.Add(('Status:     {0}' -f [string]$Record.status))
    [void]$lines.Add('')
    foreach($r in @($Record.results)){
        $wbName=[string]$r.workbook
        try { if(-not [string]::IsNullOrWhiteSpace($wbName)){ $wbName=Split-Path -Leaf $wbName } } catch { }
        [void]$lines.Add(('Workbook:  {0}' -f $wbName))
        [void]$lines.Add(('Result:    {0}' -f $(if($r.success){'SUCCESS'}else{'FAILED'})))
        [void]$lines.Add(('Saved:     {0}' -f $(if($r.saved){'Yes'}else{'No'})))
        [void]$lines.Add(('Duration:  {0}' -f $(if([int]$r.elapsed -gt 0){Format-Elapsed ([TimeSpan]::FromSeconds([int]$r.elapsed))}else{'00:00:00'})))
        [void]$lines.Add(('Queries:   {0}' -f $(if(@($r.queries).Count -gt 0){@($r.queries)-join ', '}else{'No Power Query names reported'})))
        if(-not [string]::IsNullOrWhiteSpace([string]$r.errorType)){[void]$lines.Add(('Error:     {0}' -f $r.errorType))}
        if(-not [string]::IsNullOrWhiteSpace([string]$r.message)){[void]$lines.Add(('Message:   {0}' -f $r.message))}
        [void]$lines.Add('')
    }
    [void]$lines.Add('The full details are stored in Recent Activity and the application log.')
    $txt.Text=$lines -join [Environment]::NewLine
    $form.Controls.Add($txt)

    $btnLog=New-Object System.Windows.Forms.Button; $btnLog.Text='Open Log'; $btnLog.Location=New-Object System.Drawing.Point(470,454); $btnLog.Size=New-Object System.Drawing.Size(92,30); $form.Controls.Add($btnLog)
    $btnClose=New-Object System.Windows.Forms.Button; $btnClose.Text='Close'; $btnClose.Location=New-Object System.Drawing.Point(570,454); $btnClose.Size=New-Object System.Drawing.Size(92,30); $btnClose.DialogResult='OK'; $form.Controls.Add($btnClose); $form.AcceptButton=$btnClose
    $btnLog.Add_Click({ Invoke-OpenLog }.GetNewClosure())
    $form.Add_Shown({ $form.Activate() }.GetNewClosure())
    Set-FormWithinWorkingArea -Form $form
    [void]$form.ShowDialog()
}

function Get-ApprovalReasonText {
    param([hashtable]$Request)
    $source = [string]$Request.Source
    if ($source -eq 'File') {
        $leaf = Split-Path -Leaf ([string]$Request.TriggerFile)
        if ([string]::IsNullOrWhiteSpace($leaf)) { return 'A matching file/folder event was detected.' }
        return ('A matching file event was detected: {0}' -f $leaf)
    }
    if ($source -eq 'Schedule') { return 'The configured scheduled time was reached.' }
    return ('The rule was triggered ({0}).' -f $source)
}

function Send-RefreshApprovalDecision {
    param([hashtable]$Request, [bool]$Approved, [bool]$Automatic = $false)
    $payload = @{
        RequestId        = [string]$Request.RequestId
        RuleId           = [string]$Request.RuleId
        Source           = [string]$Request.Source
        TriggerFile      = [string]$Request.TriggerFile
        TriggerFiles     = @($Request.TriggerFiles)
        AutomaticApproval = $Automatic
    }
    $command = $(if ($Approved) { 'ApproveTriggeredRun' } else { 'DeclineTriggeredRun' })
    Send-EngineCommand -Shared $script:UiShared -Type $command -Payload $payload
}

function Update-CompletedJobWorkbookInfo {
    <#  A job completion is more important than the one-minute background
        cadence: invalidate its cached paths, then poll a forced worker scan on
        every UI tick until Data updated has been repainted. #>
    $queue = $script:UiShared.WorkbookInfoRefreshRequests
    if ($null -ne $queue) {
        while ($queue.Count -gt 0) {
            $path = ''
            try { $path = [string]$queue.Dequeue() } catch { break }
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $key = $path.ToLowerInvariant()
            [void]$script:WorkbookInfoCache.Remove($key)
            [void]$script:WorkbookInfoBackoff.Remove($key)
            $script:WorkbookInfoRefreshPending = $true
        }
    }

    if (-not $script:WorkbookInfoRefreshPending -and -not $script:WorkbookInfoRefreshInProgress) { return }
    # The next job in the queue can start moments after this one finished.
    # Repainting Data updated waits; holding a handle on its workbook does not.
    if ([bool]$script:UiShared.CurrentJob.Active) { return }

    $scan = Receive-WorkbookInfoBackgroundScan
    if ($scan.Updated -or $scan.Completed) { Update-RuleList }
    if ($scan.Running) {
        $script:WorkbookInfoRefreshInProgress = $true
        return
    }

    $script:WorkbookInfoRefreshInProgress = $false
    if ($script:WorkbookInfoRefreshPending) {
        $script:WorkbookInfoLastCompleted = [DateTime]::MinValue
        $started = Start-WorkbookInfoBackgroundScan -Force
        $script:WorkbookInfoRefreshPending = $false
        $script:WorkbookInfoRefreshInProgress = [bool]$started
    }
}

function Show-RefreshApprovalPrompt {
    param([hashtable]$Request, [object[]]$Recent = @(), [object[]]$Unknown = @())

    $reason = Get-ApprovalReasonText -Request $Request
    $workbooks = @($Request.Workbooks)
    $workbookText = $(if ($workbooks.Count -gt 0) { $workbooks -join ', ' } else { 'No workbook configured' })
    $lines = New-Object System.Collections.ArrayList
    if (@($Recent).Count -gt 0) {
        [void]$lines.Add('One or more workbooks may already be up to date.')
    }
    else {
        [void]$lines.Add('A trigger is ready to refresh Excel.')
    }
    [void]$lines.Add('')
    [void]$lines.Add(('Why this ran: {0}' -f $reason))
    [void]$lines.Add(('Rule: {0}' -f [string]$Request.RuleName))
    [void]$lines.Add(('Workbook(s): {0}' -f $workbookText))

    if (@($Recent).Count -gt 0) {
        $threshold = ConvertTo-IntValue $Request.RecentRefreshPromptMinutes 0 0
        $thresholdText = $(if ($threshold -ge 60 -and ($threshold % 60) -eq 0) {
            '{0} hour(s)' -f [int]($threshold / 60)
        } else { '{0} minute(s)' -f $threshold })
        [void]$lines.Add('')
        [void]$lines.Add(('Recent-refresh window: {0}' -f $thresholdText))
        foreach ($item in @($Recent)) {
            $who = $(if ([string]::IsNullOrWhiteSpace([string]$item.Who)) { '' } else { ' by ' + [string]$item.Who })
            [void]$lines.Add(('- {0}: {1} {2}{3}' -f [string]$item.Name, [string]$item.Kind,
                ([datetime]$item.When).ToString('yyyy-MM-dd HH:mm'), $who))
        }
    }
    if (@($Unknown).Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add('Could not verify these workbooks, so confirmation is required:')
        foreach ($item in @($Unknown)) {
            $why = [string]$item.Reason
            if ([string]::IsNullOrWhiteSpace($why)) { $why = 'No query refresh time is recorded in the workbook.' }
            [void]$lines.Add(('- {0}: {1}' -f [string]$item.Name, $why))
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add('Refresh now?')
    [void]$lines.Add('')
    [void]$lines.Add('Yes = Queue the refresh')
    [void]$lines.Add('No  = Skip this trigger only')

    $choice = [System.Windows.Forms.MessageBox]::Show(
        ($lines.ToArray() -join [Environment]::NewLine),
        'Confirm triggered refresh',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
    Send-RefreshApprovalDecision -Request $Request -Approved ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Start-RefreshApprovalBackgroundCheck {
    param([Parameter(Mandatory = $true)][hashtable]$Request)
    if ($null -ne $script:UiRefreshApprovalCheckTask) { return $false }

    $paths = @($Request.WorkbookPaths)
    $threshold = ConvertTo-IntValue $Request.RecentRefreshPromptMinutes 0 0
    $localRefreshHistory = Get-SuccessfulQueryRefreshHistoryMap
    $results = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $progress = [hashtable]::Synchronized(@{ Completed = $false; Error = '' })
    $ps = [powershell]::Create()
    $worker = {
        param($AppRoot, $WorkbookPaths, $ThresholdMinutes, $LocalRefreshHistory, $Progress, $Results)
        Set-StrictMode -Version 1.0
        $ErrorActionPreference = 'Stop'
        . (Join-Path (Join-Path $AppRoot 'src') 'Common.ps1')
        . (Join-Path (Join-Path $AppRoot 'src') 'UIDiagnostics.ps1')
        . (Join-Path (Join-Path $AppRoot 'src') 'UIWorkbookInfo.ps1')
        try {
            foreach ($pathValue in @($WorkbookPaths)) {
                $path = [string]$pathValue
                $name = $(if ([string]::IsNullOrWhiteSpace($path)) { 'Unnamed workbook' } else { Split-Path -Leaf $path })
                try {
                    $info = Get-WorkbookInfo -Path $path -Refresh $true
                    $when = [DateTime]::MinValue
                    $kind = ''
                    $who = ''
                    if ($info.DataRefreshedAt -gt [DateTime]::MinValue) {
                        $when = [datetime]$info.DataRefreshedAt; $kind = 'data refreshed'; $who = [string]$info.RefreshedBy
                    }
                    $historyKey = ''
                    try { $historyKey = ([System.IO.Path]::GetFullPath($path)).TrimEnd('\').ToLowerInvariant() }
                    catch { $historyKey = $path.Trim().TrimEnd('\').ToLowerInvariant() }
                    if ($LocalRefreshHistory.ContainsKey($historyKey)) {
                        $historyWhen = [DateTime]$LocalRefreshHistory[$historyKey]
                        if ($historyWhen -gt $when) {
                            $when = $historyWhen; $kind = 'data refreshed'; $who = 'Excel Query Trigger Manager (local history)'
                        }
                    }
                    $known = ([bool]$info.Exists -and $when -gt [DateTime]::MinValue)
                    $recent = ($known -and (((Get-Date) - $when).TotalMinutes -le $ThresholdMinutes))
                    $Results.Enqueue([pscustomobject]@{
                        Path = $path; Name = $name; Known = $known; Recent = $recent
                        When = $when; Kind = $kind; Who = $who; Reason = [string]$info.Reason
                    })
                }
                catch {
                    $Results.Enqueue([pscustomobject]@{
                        Path = $path; Name = $name; Known = $false; Recent = $false
                        When = [DateTime]::MinValue; Kind = ''; Who = ''; Reason = [string]$_.Exception.Message
                    })
                }
            }
        }
        catch { $Progress.Error = [string]$_.Exception.Message }
        finally { $Progress.Completed = $true }
    }
    [void]$ps.AddScript($worker)
    [void]$ps.AddArgument([string]$script:UiPaths.AppRoot)
    [void]$ps.AddArgument($paths)
    [void]$ps.AddArgument($threshold)
    [void]$ps.AddArgument($localRefreshHistory)
    [void]$ps.AddArgument($progress)
    [void]$ps.AddArgument($results)
    $script:UiRefreshApprovalCheckTask = @{
        PowerShell = $ps; Handle = $ps.BeginInvoke(); Request = $Request
        Progress = $progress; Results = $results
    }
    return $true
}

function Stop-RefreshApprovalBackgroundCheck {
    if ($null -eq $script:UiRefreshApprovalCheckTask) { return }
    try { $script:UiRefreshApprovalCheckTask.PowerShell.Stop() } catch { }
    try { $script:UiRefreshApprovalCheckTask.PowerShell.Dispose() } catch { }
    $script:UiRefreshApprovalCheckTask = $null
}

function Receive-RefreshApprovalBackgroundCheck {
    $task = $script:UiRefreshApprovalCheckTask
    if ($null -eq $task -or -not $task.Handle.IsCompleted) { return $false }

    $request = $task.Request
    $items = New-Object System.Collections.ArrayList
    $item = $null
    while ($task.Results.TryDequeue([ref]$item)) { [void]$items.Add($item); $item = $null }
    $workerError = [string]$task.Progress.Error
    try { [void]$task.PowerShell.EndInvoke($task.Handle) } catch { if ([string]::IsNullOrWhiteSpace($workerError)) { $workerError = [string]$_.Exception.Message } }
    try { $task.PowerShell.Dispose() } catch { }
    $script:UiRefreshApprovalCheckTask = $null

    $recent = @($items.ToArray() | Where-Object { [bool]$_.Recent })
    $unknown = @($items.ToArray() | Where-Object { -not [bool]$_.Known })
    if (-not [string]::IsNullOrWhiteSpace($workerError)) {
        $unknown += [pscustomobject]@{ Name = 'Workbook check'; Reason = $workerError }
    }
    if ($items.Count -eq 0 -and $unknown.Count -eq 0) {
        $unknown += [pscustomobject]@{ Name = 'Workbook check'; Reason = 'No workbook path was available to check.' }
    }

    if ($recent.Count -gt 0 -or $unknown.Count -gt 0 -or (ConvertTo-BoolValue $request.AlwaysAsk $false)) {
        Show-RefreshApprovalPrompt -Request $request -Recent $recent -Unknown $unknown
    }
    else {
        Send-RefreshApprovalDecision -Request $request -Approved $true -Automatic $true
    }
    return $true
}

function Update-RefreshApprovalPrompts {
    <#  Metadata reads happen in a worker runspace. The UI timer only starts or
        receives that work, so an unavailable share never freezes the Dashboard. #>
    if ($script:UiExiting) { return }
    if ($null -ne $script:UiRefreshApprovalCheckTask) {
        [void](Receive-RefreshApprovalBackgroundCheck)
        return
    }

    $queue = $script:UiShared.RefreshApprovalRequests
    if ($null -eq $queue -or $queue.Count -eq 0) { return }
    $request = $queue.Dequeue()
    if ($null -eq $request) { return }

    $threshold = ConvertTo-IntValue $request.RecentRefreshPromptMinutes 0 0
    if ($threshold -gt 0) {
        [void](Start-RefreshApprovalBackgroundCheck -Request $request)
    }
    else {
        Show-RefreshApprovalPrompt -Request $request
    }
}

function Update-TriggeredRefreshResultPopups {
    if (-not (ConvertTo-BoolValue $script:UiConfig.appSettings.showTriggeredResultPopup $true)) { return }

    $history=@(Copy-SharedList $script:UiShared.History)
    if(-not $script:UiTriggeredHistoryInitialized){
        foreach($record in $history){
            if([string]$record.ruleId -ne 'MANUAL-REFRESH' -and [string]$record.triggerSource -ne 'Manual'){
                $script:UiSeenTriggeredHistory[[string]$record.jobId]=$true
            }
        }
        $script:UiTriggeredHistoryInitialized=$true
        return
    }

    foreach($record in $history){
        if([string]$record.ruleId -eq 'MANUAL-REFRESH' -or [string]$record.triggerSource -eq 'Manual'){ continue }
        $jobId=[string]$record.jobId
        if([string]::IsNullOrWhiteSpace($jobId) -or $script:UiSeenTriggeredHistory.ContainsKey($jobId)){ continue }
        $script:UiSeenTriggeredHistory[$jobId]=$true
        Show-TriggeredRefreshResultDialog -Record $record
    }
}

function Update-DashboardTick {
    <#  The single UI heartbeat. Everything visual happens here.  #>
    try {
        $shared = $script:UiShared
        $status = [string]$shared.Status

        if ($status -ne $script:UiLastStatus -or $status -eq 'Refreshing') {
            $script:UiLastStatus = $status
            $script:UiControls.Status.Text      = ('Status: {0}' -f $status)
            $script:UiControls.Status.ForeColor = Get-StatusColor $status
            $script:UiControls.BtnPause.Text    = $(if ($shared.Paused) { 'Resume Monitoring' } else { 'Pause Monitoring' })
            if ($null -ne $script:UiBottomTips) {
                $script:UiBottomTips.SetToolTip($script:UiControls.BtnPause, $(if ($shared.Paused) {
                    'Starts watching again. Anything that changed while paused is not picked up.'
                } else {
                    'Stops rules setting themselves off, but leaves the application running. Nothing is lost - it simply stops watching until you resume.'
                }))
            }
        }
        $statusDetailText = [string]$shared.StatusDetail
        $script:UiControls.StatusDetail.Text = $statusDetailText
        if ($script:UiControls.ContainsKey('StatusToolTip')) {
            $script:UiControls.StatusToolTip.SetToolTip($script:UiControls.StatusDetail, $statusDetailText)
        }

        $iconKey = Get-StatusIconKey $status
        if ($iconKey -ne $script:UiLastIconKey) {
            $script:UiLastIconKey = $iconKey
            $script:UiControls.Tray.Icon = $script:UiIcons[$iconKey]
        }
        $trayText = 'Excel Query Trigger - {0}' -f $status
        if ($trayText.Length -gt 63) { $trayText = $trayText.Substring(0, 63) }
        if ($script:UiControls.Tray.Text -ne $trayText) { $script:UiControls.Tray.Text = $trayText }

        Update-CurrentJobDisplay
        Update-ExcelWatchdog
        Invoke-BackgroundUpdateCheck
        Update-RunNowCaption
        Update-ActivityFeed
        Update-PendingList
        Update-RunningRuleHighlight
        Update-Notifications
        Update-RefreshApprovalPrompts
        Update-ManualRefreshResultPopups
        Update-TriggeredRefreshResultPopups
        Update-CompletedJobWorkbookInfo

        if ([int]$shared.ConfigVersion -ne $script:UiConfigVersion) {
            $script:UiConfigVersion = [int]$shared.ConfigVersion
            Update-RuleList
            $script:UiLastRuleRefresh = Get-Date
        }
        elseif (((Get-Date) - $script:UiLastRuleRefresh).TotalSeconds -ge 3) {
            Update-RuleList
            $script:UiLastRuleRefresh = Get-Date
        }
        elseif (((Get-Date) - $script:UiLastWorkbookScan).TotalSeconds -ge 10) {
            # Workbook metadata is read a few files at a time, off the redraw
            # path, so a share that has gone away cannot stall the dashboard.
            $script:UiLastWorkbookScan = Get-Date
            Update-WorkbookInfoAndRuleList
        }

        if (-not $shared.EngineAlive -and -not $script:UiExiting -and -not [string]::IsNullOrWhiteSpace([string]$shared.FatalError)) {
            $script:UiControls.Timer.Stop()
            [System.Windows.Forms.MessageBox]::Show(
                ('The monitoring engine stopped unexpectedly:' + [Environment]::NewLine + [string]$shared.FatalError +
                 [Environment]::NewLine + [Environment]::NewLine + 'Please restart the application.'),
                'Excel Query Trigger', 'OK', 'Error') | Out-Null
        }
    }
    catch {
        # A failing tick must never take the window down.
    }
}
