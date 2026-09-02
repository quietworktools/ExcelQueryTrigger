# ==============================================================================
#  StartupManager.ps1
#  HKCU\...\Run only - never HKLM, so no administrator rights are needed.
#  The command points at Start-AtLogon.vbs so Windows does not flash a console
#  window at every logon, and so the application can tell a logon start apart
#  from someone launching it by hand during the day.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:StartupKeyPath   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:StartupValueName = 'ExcelQueryTrigger'
# Windows records "the user switched this off in Settings > Startup Apps" here
# rather than deleting the Run value. The Run value therefore keeps looking
# perfectly healthy while nothing actually starts.
$script:StartupApprovedKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'

function Test-StartupApprovedDisabled {
    <#
        Has Windows explicitly disabled our startup item?

        The value is a 12-byte blob whose first byte carries the enabled flag:
        even (0x02 / 0x06) means enabled, odd (0x03 / 0x07) means the user
        turned it off. The remaining bytes are a FILETIME of when that happened
        and are deliberately not read - that part is the undocumented half.

        This is read-only and fails open in every ambiguous case: no key, no
        value, an unreadable value or an unexpected length all return $false,
        which is exactly what the application assumed before this function
        existed. The only new answer it can give is a positive one, so it can
        report a disabled item but can never invent a broken registration.
    #>
    try {
        if (-not (Test-Path -LiteralPath $script:StartupApprovedKeyPath)) { return $false }
        $property = Get-ItemProperty -LiteralPath $script:StartupApprovedKeyPath `
            -Name $script:StartupValueName -ErrorAction SilentlyContinue
        if ($null -eq $property) { return $false }

        $bytes = $property.$($script:StartupValueName)
        if ($null -eq $bytes) { return $false }
        if (-not ($bytes -is [byte[]])) { return $false }
        if ($bytes.Length -lt 1) { return $false }

        return (($bytes[0] -band 1) -eq 1)
    }
    catch {
        return $false
    }
}

function Get-StartupRegistrationHealth {
    <#
        Returns @{
            Registered      = [bool]   the Run value exists and is not empty
            Command         = [string] what it actually runs
            LauncherValid   = [bool]   the file that command points at exists
            PointsToCurrent = [bool]   it points at this installation
            DisabledByWindows = [bool] Startup Apps has switched it off
            Healthy         = [bool]   all of the above are in order
            Reason          = [string] '' when healthy, otherwise one sentence
        }

        Test-StartupRegistration keeps its original meaning (the entry exists)
        so nothing that already depends on it changes behaviour.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Paths)

    $health = @{
        Registered        = $false
        Command           = ''
        LauncherValid     = $false
        PointsToCurrent   = $false
        DisabledByWindows = $false
        Healthy           = $false
        Reason            = 'The Windows startup entry is missing.'
    }

    try {
        if (-not (Test-StartupRegistration)) { return $health }

        $health.Registered = $true
        $health.Command = [string](Get-ItemProperty -LiteralPath $script:StartupKeyPath `
            -Name $script:StartupValueName -ErrorAction Stop).$($script:StartupValueName)

        $expected = Get-StartupCommand -Paths $Paths
        $health.PointsToCurrent = ($health.Command -eq $expected)

        # Pull the launcher out of the command line: it is the last quoted
        # argument before any switches, which covers both the wscript.exe form
        # and the powershell.exe fallback.
        $launcher = ''
        $quoted = [regex]::Matches($health.Command, '"([^"]+)"')
        foreach ($match in $quoted) {
            $candidate = [string]$match.Groups[1].Value
            if ($candidate -match '\.(vbs|ps1)$') { $launcher = $candidate }
        }
        if ([string]::IsNullOrWhiteSpace($launcher) -and $quoted.Count -gt 0) {
            $launcher = [string]$quoted[0].Groups[1].Value
        }
        $health.LauncherValid = (-not [string]::IsNullOrWhiteSpace($launcher)) -and `
                                (Test-Path -LiteralPath $launcher)

        $health.DisabledByWindows = Test-StartupApprovedDisabled

        if ($health.DisabledByWindows) {
            $health.Reason = 'Windows has switched this startup item off in Settings > Apps > Startup.'
        }
        elseif (-not $health.LauncherValid) {
            $health.Reason = 'The Windows startup entry points at a launcher file that no longer exists.'
        }
        elseif (-not $health.PointsToCurrent) {
            $health.Reason = 'The Windows startup entry points at a different copy of the application.'
        }
        else {
            $health.Reason  = ''
            $health.Healthy = $true
        }
    }
    catch {
        $health.Reason = ('The Windows startup entry could not be read: {0}' -f $_.Exception.Message)
    }

    return $health
}

function Get-StartupCommand {
    param([Parameter(Mandatory = $true)][hashtable]$Paths)

    if (Test-Path -LiteralPath $Paths.LogonLauncherPath) {
        $wscriptExe = Join-Path $env:WINDIR 'System32\wscript.exe'
        return ('"{0}" "{1}"' -f $wscriptExe, $Paths.LogonLauncherPath)
    }
    if (Test-Path -LiteralPath $Paths.LauncherPath) {
        $wscriptExe = Join-Path $env:WINDIR 'System32\wscript.exe'
        return ('"{0}" "{1}"' -f $wscriptExe, $Paths.LauncherPath)
    }

    $mainScript = Join-Path $Paths.AppRoot 'ExcelQueryTrigger.ps1'
    $powershellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    return ('"{0}" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "{1}" -StartedFromLogon' -f $powershellExe, $mainScript)
}

function Update-StartupRegistration {
    <#
        Silently repoints an existing entry at the current launcher. Needed
        after an upgrade: an installation registered before Start-AtLogon.vbs
        existed would otherwise keep starting without -StartedFromLogon, and the
        logon prompt would never appear.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Paths)

    try {
        if (-not (Test-StartupRegistration)) { return $false }
        $current = [string](Get-ItemProperty -LiteralPath $script:StartupKeyPath -Name $script:StartupValueName -ErrorAction Stop).$($script:StartupValueName)
        $wanted  = Get-StartupCommand -Paths $Paths
        if ($current -eq $wanted) { return $false }

        Set-ItemProperty -LiteralPath $script:StartupKeyPath -Name $script:StartupValueName -Value $wanted -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Test-StartupRegistration {
    try {
        if (-not (Test-Path -LiteralPath $script:StartupKeyPath)) { return $false }
        $property = Get-ItemProperty -LiteralPath $script:StartupKeyPath -Name $script:StartupValueName -ErrorAction SilentlyContinue
        return ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.$($script:StartupValueName)))
    }
    catch {
        return $false
    }
}

function Set-StartupRegistration {
    <#  Returns @{ Success = [bool]; Message = [string] }  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Paths,
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    try {
        if (-not (Test-Path -LiteralPath $script:StartupKeyPath)) {
            New-Item -Path $script:StartupKeyPath -Force -ErrorAction Stop | Out-Null
        }

        if ($Enabled) {
            $command = Get-StartupCommand -Paths $Paths
            New-ItemProperty -LiteralPath $script:StartupKeyPath -Name $script:StartupValueName `
                -Value $command -PropertyType String -Force -ErrorAction Stop | Out-Null
            return @{ Success = $true; Message = ('Registered to start with Windows: {0}' -f $command) }
        }

        if (Test-StartupRegistration) {
            Remove-ItemProperty -LiteralPath $script:StartupKeyPath -Name $script:StartupValueName -Force -ErrorAction Stop
        }
        return @{ Success = $true; Message = 'Removed from Windows startup.' }
    }
    catch {
        return @{ Success = $false; Message = ('Could not update the startup entry: {0}' -f $_.Exception.Message) }
    }
}
