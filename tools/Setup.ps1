# ==============================================================================
#  Setup.ps1  -  install, update or remove Excel Query Trigger Manager
#
#  Everything happens under the signed-in user's own profile:
#    program files   %LOCALAPPDATA%\ExcelQueryTrigger
#    shortcuts       Start Menu, and the desktop if asked
#    start at logon  HKCU\...\Run   (the same value the application writes)
#
#  No administrator rights, no HKLM, no services, nothing installed for other
#  users. Removing it puts all of that back.
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$Uninstall,
    [int]$WaitForPid = 0,
    [switch]$UpdateMode
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

$SourceRoot  = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$InstallRoot = Join-Path $env:LOCALAPPDATA 'ExcelQueryTrigger'
$IconPath    = Join-Path $SourceRoot 'assets\ExcelQueryTrigger.ico'
$StartMenu   = Join-Path ([Environment]::GetFolderPath('Programs')) 'Excel Query Trigger Manager.lnk'
$DesktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Excel Query Trigger Manager.lnk'
$RunKey      = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$RunValue    = 'ExcelQueryTrigger'

# ------------------------------------------------------------------ wording --

$Text = @{
    en = @{
        'title'        = 'Excel Query Trigger Manager - Setup'
        'lang'         = 'Language'
        'intro'        = 'Refreshes Excel queries in the background, on a schedule or when a file arrives, so you do not have to open each workbook yourself.'
        'willInstall'  = 'This will install version {0}.'
        'willUpdate'   = 'This will update the installed version {0} to {1}. Your rules, history and logs are kept.'
        'willReinstall'= 'Version {0} is already installed. Installing again replaces the program files and keeps your rules, history and logs.'
        'destination'  = 'It goes here (no administrator rights needed):'
        'noticeTitle'  = 'Before first use - important'
        'backupNotice' = 'Back up important workbooks before using automatic refresh.'
        'trayNotice'   = 'Closing the Dashboard does not stop the app. It stays running in the system tray.'
        'options'      = 'After installing'
        'optStart'     = 'Open it now'
        'optLogon'     = 'Start it automatically when I sign in to Windows'
        'optDesktop'   = 'Put a shortcut on my desktop'
        'btnInstall'   = 'Install'
        'btnUpdate'    = 'Update'
        'btnRemove'    = 'Remove'
        'btnClose'     = 'Close'
        'btnDone'      = 'Done'
        'running'      = 'Excel Query Trigger Manager is running. Exit it from the tray icon, then press the button again.'
        'runningTitle' = 'Already running'
        'stopFailed'   = 'It could not be closed automatically. Please exit it from the tray icon and try again.'
        'confirmRemove'= 'Remove Excel Query Trigger Manager?{0}{0}The program files, the shortcuts and the sign-in entry are removed.'
        'keepData'     = 'Keep my rules and history?{0}{0}Yes - keep them, so reinstalling picks up where you left off.{0}No  - delete everything.'
        'stepUnblock'  = 'Clearing the downloaded-file mark...'
        'stepCopy'     = 'Copying program files...'
        'stepConfig'   = 'Preparing the settings folder...'
        'stepShortcut' = 'Creating shortcuts...'
        'stepLogon'    = 'Registering the sign-in entry...'
        'stepLogonOff' = 'Removing the sign-in entry...'
        'stepRemove'   = 'Removing program files...'
        'stepLaunch'   = 'Starting the application...'
        'doneInstall'  = 'Installed. It is in your Start menu as "Excel Query Trigger Manager".'
        'doneRemove'   = 'Removed.'
        'failed'       = 'Setup could not finish: {0}'
        'sourceBad'    = 'This copy looks incomplete - {0} was not found next to Setup. Please unzip the whole folder and run Setup again.'
        'openFolder'   = 'Open the folder'
    }
    ja = @{
        'title'        = 'Excel Query Trigger Manager - セットアップ'
        'lang'         = '言語'
        'intro'        = 'Excel のクエリを、時刻やファイルの到着に応じてバックグラウンドで更新します。ブックを1つずつ開いて更新する必要がなくなります。'
        'willInstall'  = 'バージョン {0} をインストールします。'
        'willUpdate'   = 'インストール済みのバージョン {0} を {1} に更新します。ルール・履歴・ログはそのまま残ります。'
        'willReinstall'= 'バージョン {0} はインストール済みです。もう一度実行するとプログラムを入れ直します。ルール・履歴・ログはそのまま残ります。'
        'destination'  = 'インストール先（管理者権限は不要です）:'
        'noticeTitle'  = '初回利用前の重要事項'
        'backupNotice' = '自動更新を使う前に、重要な Excel ファイルを必ずバックアップしてください。'
        'trayNotice'   = 'Dashboard を閉じても終了しません。プログラムはタスクトレイに常駐します。'
        'options'      = 'インストール後の動作'
        'optStart'     = 'すぐに起動する'
        'optLogon'     = 'Windows にサインインしたら自動的に起動する'
        'optDesktop'   = 'デスクトップにショートカットを作る'
        'btnInstall'   = 'インストール'
        'btnUpdate'    = '更新'
        'btnRemove'    = 'アンインストール'
        'btnClose'     = '閉じる'
        'btnDone'      = '完了'
        'running'      = 'Excel Query Trigger Manager が起動しています。タスクトレイのアイコンから終了し、もう一度ボタンを押してください。'
        'runningTitle' = '起動中です'
        'stopFailed'   = '自動で終了できませんでした。タスクトレイのアイコンから終了してから、もう一度お試しください。'
        'confirmRemove'= 'Excel Query Trigger Manager をアンインストールしますか？{0}{0}プログラム本体・ショートカット・サインイン時の自動起動を削除します。'
        'keepData'     = 'ルールと履歴を残しますか？{0}{0}はい － 残します。入れ直したときにそのまま使えます。{0}いいえ － すべて削除します。'
        'stepUnblock'  = 'ダウンロード時のブロックを解除しています...'
        'stepCopy'     = 'プログラムをコピーしています...'
        'stepConfig'   = '設定フォルダーを準備しています...'
        'stepShortcut' = 'ショートカットを作成しています...'
        'stepLogon'    = 'サインイン時の自動起動を登録しています...'
        'stepLogonOff' = 'サインイン時の自動起動を解除しています...'
        'stepRemove'   = 'プログラムを削除しています...'
        'stepLaunch'   = 'アプリを起動しています...'
        'doneInstall'  = 'インストールしました。スタートメニューに「Excel Query Trigger Manager」があります。'
        'doneRemove'   = 'アンインストールしました。'
        'failed'       = 'セットアップを完了できませんでした: {0}'
        'sourceBad'    = 'ファイルが揃っていないようです（{0} が見つかりません）。ZIP をフォルダーごと展開してから、もう一度 Setup を実行してください。'
        'openFolder'   = 'フォルダーを開く'
    }
}

$script:Lang = $(if (([System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName) -eq 'ja') { 'ja' } else { 'en' })
function T([string]$Key) {
    if ($Text[$script:Lang].ContainsKey($Key)) { return $Text[$script:Lang][$Key] }
    return $Text['en'][$Key]
}

# ------------------------------------------------------------------- helpers --

function Get-VersionAt {
    <#  Reads $script:AppVersion out of a copy's Common.ps1. '' when absent.  #>
    param([string]$Root)
    $common = Join-Path $Root 'src\Common.ps1'
    if (-not (Test-Path -LiteralPath $common)) { return '' }
    try {
        $match = [regex]::Match((Get-Content -LiteralPath $common -Raw), "AppVersion\s*=\s*'([^']+)'")
        if ($match.Success) { return $match.Groups[1].Value }
    }
    catch { }
    return ''
}

function Test-AppRunning {
    <#  The application's own single-instance mutex answers this exactly.  #>
    $mutex = $null
    try {
        $mutex = New-Object System.Threading.Mutex($false, 'Local\ExcelQueryTriggerSingleInstance')
        if ($mutex.WaitOne(0, $false)) { $mutex.ReleaseMutex(); return $false }
        return $true
    }
    catch { return $false }
    finally { if ($null -ne $mutex) { try { $mutex.Dispose() } catch { } } }
}

function Wait-ForProcessExit {
    param([int]$ProcessId, [int]$Seconds = 90)
    if ($ProcessId -le 0 -or $ProcessId -eq $PID) { return $true }
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if ($null -eq (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return ($null -eq (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue))
}

function New-AppShortcut {
    param([string]$LinkPath, [string]$TargetRoot)
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($LinkPath)
    $link.TargetPath       = Join-Path $env:WINDIR 'System32\wscript.exe'
    $link.Arguments        = ('"{0}"' -f (Join-Path $TargetRoot 'Start-Hidden.vbs'))
    $link.WorkingDirectory = $TargetRoot
    $link.Description      = 'Refresh Excel queries in the background'
    $icon = Join-Path $TargetRoot 'assets\ExcelQueryTrigger.ico'
    if (Test-Path -LiteralPath $icon) { $link.IconLocation = ('{0},0' -f $icon) }
    $link.Save()
}

function Copy-Program {
    param([string]$From, [string]$To, [scriptblock]$Report)

    if (-not (Test-Path -LiteralPath $To)) { New-Item -ItemType Directory -Path $To -Force | Out-Null }

    # Everything except the user's own data. config\ and logs\ are handled after.
    foreach ($entry in (Get-ChildItem -LiteralPath $From -Force)) {
        if ($entry.Name -in @('config', 'logs')) { continue }
        $target = Join-Path $To $entry.Name
        if ($entry.PSIsContainer) {
            if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
            Copy-Item -LiteralPath $entry.FullName -Destination $target -Recurse -Force
        }
        else {
            Copy-Item -LiteralPath $entry.FullName -Destination $target -Force
        }
    }

    & $Report (T 'stepConfig')
    $configDir = Join-Path $To 'config'
    if (-not (Test-Path -LiteralPath $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
    $rules = Join-Path $configDir 'rules.json'
    if (-not (Test-Path -LiteralPath $rules)) {
        $template = Join-Path $From 'config\rules.json'
        if (Test-Path -LiteralPath $template) { Copy-Item -LiteralPath $template -Destination $rules -Force }
    }
    $logDir = Join-Path $To 'logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
}

# An in-app update launches this copy first, then closes itself cleanly. Wait for
# that exact process instead of searching for and killing PowerShell processes.
if ($WaitForPid -gt 0 -and -not (Wait-ForProcessExit -ProcessId $WaitForPid)) {
    [System.Windows.Forms.MessageBox]::Show((T 'stopFailed'), (T 'title'), 'OK', 'Warning') | Out-Null
    return
}
for ($wait = 0; $wait -lt 50 -and (Test-AppRunning); $wait++) { Start-Sleep -Milliseconds 200 }

# ---------------------------------------------------------------------- form --

$installedVersion = Get-VersionAt $InstallRoot
$sourceVersion    = Get-VersionAt $SourceRoot
$isInstalled      = -not [string]::IsNullOrWhiteSpace($installedVersion)
$isFirstInstall   = -not $isInstalled

if ([string]::IsNullOrWhiteSpace($sourceVersion)) {
    [System.Windows.Forms.MessageBox]::Show(([string]::Format((T 'sourceBad'), 'src\Common.ps1')),
        (T 'title'), 'OK', 'Warning') | Out-Null
    return
}

$form = New-Object System.Windows.Forms.Form
$form.Text            = T 'title'
$formHeight           = $(if ($isFirstInstall) { 524 } else { 420 })
$form.ClientSize      = New-Object System.Drawing.Size(560, $formHeight)
$form.FormBorderStyle = 'FixedDialog'
$form.StartPosition   = 'CenterScreen'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
if (Test-Path -LiteralPath $IconPath) { try { $form.Icon = New-Object System.Drawing.Icon($IconPath) } catch { } }

$uiFont = $(
    try { New-Object System.Drawing.Font($(if ($script:Lang -eq 'ja') { 'Meiryo UI' } else { 'Segoe UI' }), 9) }
    catch { [System.Drawing.SystemFonts]::MessageBoxFont })
$form.Font = $uiFont

$banner = New-Object System.Windows.Forms.Label
$banner.Text = 'Excel Query Trigger Manager'
$banner.Font = New-Object System.Drawing.Font($uiFont.FontFamily, 14, [System.Drawing.FontStyle]::Bold)
$banner.Location = New-Object System.Drawing.Point(20, 18)
$banner.Size = New-Object System.Drawing.Size(400, 30)
$form.Controls.Add($banner)

$cmbLang = New-Object System.Windows.Forms.ComboBox
$cmbLang.DropDownStyle = 'DropDownList'
$cmbLang.Location = New-Object System.Drawing.Point(430, 20)
$cmbLang.Size = New-Object System.Drawing.Size(110, 24)
[void]$cmbLang.Items.AddRange(@('English', '日本語'))
$cmbLang.SelectedIndex = $(if ($script:Lang -eq 'ja') { 1 } else { 0 })
$form.Controls.Add($cmbLang)

$lblIntro = New-Object System.Windows.Forms.Label
$lblIntro.Location = New-Object System.Drawing.Point(20, 54)
$lblIntro.Size = New-Object System.Drawing.Size(520, 40)
$form.Controls.Add($lblIntro)

$lblPlan = New-Object System.Windows.Forms.Label
$lblPlan.Location = New-Object System.Drawing.Point(20, 100)
$lblPlan.Size = New-Object System.Drawing.Size(520, 36)
$lblPlan.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblPlan)

$lblWhere = New-Object System.Windows.Forms.Label
$lblWhere.Location = New-Object System.Drawing.Point(20, 142)
$lblWhere.Size = New-Object System.Drawing.Size(520, 18)
$form.Controls.Add($lblWhere)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(20, 162)
$txtPath.Size = New-Object System.Drawing.Size(520, 22)
$txtPath.ReadOnly = $true
$txtPath.Text = $InstallRoot
$form.Controls.Add($txtPath)

$grpNotice = New-Object System.Windows.Forms.GroupBox
$grpNotice.Location = New-Object System.Drawing.Point(20, 196)
$grpNotice.Size = New-Object System.Drawing.Size(520, 104)
$grpNotice.Visible = $isFirstInstall
$form.Controls.Add($grpNotice)

$lblBackup = New-Object System.Windows.Forms.Label
$lblBackup.Location = New-Object System.Drawing.Point(16, 22)
$lblBackup.Size = New-Object System.Drawing.Size(488, 28)
$lblBackup.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$lblBackup.ForeColor = [System.Drawing.Color]::FromArgb(170, 45, 35)
$grpNotice.Controls.Add($lblBackup)

$lblTray = New-Object System.Windows.Forms.Label
$lblTray.Location = New-Object System.Drawing.Point(16, 56)
$lblTray.Size = New-Object System.Drawing.Size(488, 36)
$lblTray.Font = New-Object System.Drawing.Font($uiFont, [System.Drawing.FontStyle]::Bold)
$lblTray.ForeColor = [System.Drawing.Color]::FromArgb(30, 70, 125)
$grpNotice.Controls.Add($lblTray)

$contentOffset = $(if ($isFirstInstall) { 104 } else { 0 })

$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Location = New-Object System.Drawing.Point(20, (196 + $contentOffset))
$grpOptions.Size = New-Object System.Drawing.Size(520, 110)
$form.Controls.Add($grpOptions)

$chkStart = New-Object System.Windows.Forms.CheckBox
$chkStart.Location = New-Object System.Drawing.Point(16, 24); $chkStart.Size = New-Object System.Drawing.Size(490, 22)
$chkStart.Checked = $true
$grpOptions.Controls.Add($chkStart)

$chkLogon = New-Object System.Windows.Forms.CheckBox
$chkLogon.Location = New-Object System.Drawing.Point(16, 50); $chkLogon.Size = New-Object System.Drawing.Size(490, 22)
$chkLogon.Checked = $(if ($isInstalled) {
    try { -not [string]::IsNullOrWhiteSpace([string](Get-ItemPropertyValue -LiteralPath $RunKey -Name $RunValue -ErrorAction Stop)) } catch { $false }
} else { $true })
$grpOptions.Controls.Add($chkLogon)

$chkDesktop = New-Object System.Windows.Forms.CheckBox
$chkDesktop.Location = New-Object System.Drawing.Point(16, 76); $chkDesktop.Size = New-Object System.Drawing.Size(490, 22)
$chkDesktop.Checked = (Test-Path -LiteralPath $DesktopLink)
$grpOptions.Controls.Add($chkDesktop)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, (316 + $contentOffset))
$lblStatus.Size = New-Object System.Drawing.Size(520, 36)
$form.Controls.Add($lblStatus)

$btnPrimary = New-Object System.Windows.Forms.Button
$btnPrimary.Location = New-Object System.Drawing.Point(320, (366 + $contentOffset)); $btnPrimary.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($btnPrimary)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Location = New-Object System.Drawing.Point(20, (366 + $contentOffset)); $btnRemove.Size = New-Object System.Drawing.Size(130, 32)
$btnRemove.Enabled = $isInstalled
$form.Controls.Add($btnRemove)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(440, (366 + $contentOffset)); $btnClose.Size = New-Object System.Drawing.Size(100, 32)
$btnClose.DialogResult = 'Cancel'
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

function Update-Wording {
    $form.Text        = T 'title'
    $lblIntro.Text    = T 'intro'
    $lblWhere.Text    = T 'destination'
    $grpNotice.Text  = T 'noticeTitle'
    $lblBackup.Text  = T 'backupNotice'
    $lblTray.Text    = T 'trayNotice'
    $grpOptions.Text  = T 'options'
    $chkStart.Text    = T 'optStart'
    $chkLogon.Text    = T 'optLogon'
    $chkDesktop.Text  = T 'optDesktop'
    $btnRemove.Text   = T 'btnRemove'
    $btnClose.Text    = T 'btnClose'
    $btnPrimary.Text  = $(if ($isInstalled -and $installedVersion -ne $sourceVersion) { T 'btnUpdate' } else { T 'btnInstall' })
    $lblPlan.Text     = $(
        if (-not $isInstalled) { [string]::Format((T 'willInstall'), $sourceVersion) }
        elseif ($installedVersion -ne $sourceVersion) { [string]::Format((T 'willUpdate'), $installedVersion, $sourceVersion) }
        else { [string]::Format((T 'willReinstall'), $installedVersion) })
}

$cmbLang.Add_SelectedIndexChanged({
    $script:Lang = $(if ($cmbLang.SelectedIndex -eq 1) { 'ja' } else { 'en' })
    try {
        $name = $(if ($script:Lang -eq 'ja') { 'Meiryo UI' } else { 'Segoe UI' })
        $form.Font = New-Object System.Drawing.Font($name, 9)
    }
    catch { }
    Update-Wording
})

$setBusy = {
    param([bool]$Busy)
    foreach ($control in @($btnPrimary, $btnRemove, $btnClose, $cmbLang, $chkStart, $chkLogon, $chkDesktop)) { $control.Enabled = (-not $Busy) }
    if (-not $Busy) { $btnRemove.Enabled = $isInstalled }
    $form.UseWaitCursor = $Busy
    [System.Windows.Forms.Application]::DoEvents()
}
$report = {
    param([string]$Message)
    $lblStatus.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}
$script:LaunchAfterSetup = $false

$btnPrimary.Add_Click({
    if (Test-AppRunning) {
        [System.Windows.Forms.MessageBox]::Show((T 'running'), (T 'runningTitle'), 'OK', 'Information') | Out-Null
        return
    }

    $installSucceeded = $false
    $script:LaunchAfterSetup = $false
    & $setBusy $true
    try {
        # A zip from the network carries a block mark on every file, which stops
        # PowerShell from running them.
        & $report (T 'stepUnblock')
        try { Get-ChildItem -LiteralPath $SourceRoot -Recurse -File | Unblock-File -ErrorAction SilentlyContinue } catch { }

        & $report (T 'stepCopy')
        Copy-Program -From $SourceRoot -To $InstallRoot -Report $report
        try { Get-ChildItem -LiteralPath $InstallRoot -Recurse -File | Unblock-File -ErrorAction SilentlyContinue } catch { }

        & $report (T 'stepShortcut')
        New-AppShortcut -LinkPath $StartMenu -TargetRoot $InstallRoot
        if ($chkDesktop.Checked) { New-AppShortcut -LinkPath $DesktopLink -TargetRoot $InstallRoot }
        elseif (Test-Path -LiteralPath $DesktopLink) { Remove-Item -LiteralPath $DesktopLink -Force -ErrorAction SilentlyContinue }

        if ($chkLogon.Checked) {
            & $report (T 'stepLogon')
            if (-not (Test-Path -LiteralPath $RunKey)) { New-Item -Path $RunKey -Force | Out-Null }
            Set-ItemProperty -LiteralPath $RunKey -Name $RunValue `
                -Value ('"{0}" "{1}"' -f (Join-Path $env:WINDIR 'System32\wscript.exe'), (Join-Path $InstallRoot 'Start-AtLogon.vbs'))
        }

        & $report (T 'doneInstall')
        $script:installedVersion = $sourceVersion
        $script:isInstalled = $true
        $btnPrimary.Text = T 'btnInstall'
        $script:LaunchAfterSetup = $chkStart.Checked
        $installSucceeded = $true
    }
    catch {
        & $report ([string]::Format((T 'failed'), $_.Exception.Message))
    }
    finally {
        & $setBusy $false
    }
    if ($installSucceeded) {
        # Close Setup immediately. If requested, the app is started only after
        # ShowDialog returns so its splash cannot be hidden behind Setup.
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }
})

$btnRemove.Add_Click({
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ([string]::Format((T 'confirmRemove'), [Environment]::NewLine)), (T 'title'), 'YesNo', 'Warning')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $keepData = [System.Windows.Forms.MessageBox]::Show(
        ([string]::Format((T 'keepData'), [Environment]::NewLine)), (T 'title'), 'YesNo', 'Question')

    if (Test-AppRunning) {
        [System.Windows.Forms.MessageBox]::Show((T 'running'), (T 'runningTitle'), 'OK', 'Information') | Out-Null
        return
    }

    & $setBusy $true
    try {
        & $report (T 'stepLogonOff')
        try { Remove-ItemProperty -LiteralPath $RunKey -Name $RunValue -ErrorAction SilentlyContinue } catch { }
        foreach ($link in @($StartMenu, $DesktopLink)) {
            if (Test-Path -LiteralPath $link) { Remove-Item -LiteralPath $link -Force -ErrorAction SilentlyContinue }
        }

        & $report (T 'stepRemove')
        if ($keepData -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($entry in (Get-ChildItem -LiteralPath $InstallRoot -Force -ErrorAction SilentlyContinue)) {
                if ($entry.Name -in @('config', 'logs')) { continue }
                Remove-Item -LiteralPath $entry.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        elseif (Test-Path -LiteralPath $InstallRoot) {
            Remove-Item -LiteralPath $InstallRoot -Recurse -Force -ErrorAction SilentlyContinue
        }

        & $report (T 'doneRemove')
        $script:isInstalled = $false
        $script:installedVersion = ''
        $btnPrimary.Text = T 'btnInstall'
    }
    catch {
        & $report ([string]::Format((T 'failed'), $_.Exception.Message))
    }
    finally {
        & $setBusy $false
    }
})

Update-Wording
if ($Uninstall) { $form.Add_Shown({ $btnRemove.PerformClick() }) }
elseif ($UpdateMode) {
    # Install it was already confirmed inside the running application. Show
    # progress, but do not require a second click in Setup.
    $form.Add_Shown({ $btnPrimary.PerformClick() })
}
[void]$form.ShowDialog()
$form.Dispose()
if ($script:LaunchAfterSetup) {
    $launchArguments = ('"{0}"' -f (Join-Path $InstallRoot 'Start-Hidden.vbs'))
    if ($isFirstInstall -or $UpdateMode) { $launchArguments += ' -ShowWindow' }
    Start-Process -FilePath (Join-Path $env:WINDIR 'System32\wscript.exe') `
        -ArgumentList $launchArguments `
        -WorkingDirectory $InstallRoot | Out-Null
}
