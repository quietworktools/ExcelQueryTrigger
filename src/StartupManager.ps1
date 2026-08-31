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
