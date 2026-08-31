$ErrorActionPreference = 'Stop'

$launchers = @('Start-Hidden.vbs', 'Start-AtLogon.vbs')
foreach ($path in $launchers) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing launcher: $path"
    }

    $bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $path))
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "$path has a UTF-8 BOM. Windows Script Host may fail at line 1, char 1."
    }

    $text = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $path), [System.Text.Encoding]::UTF8)
    if ($text -notmatch '(?im)^\s*Option Explicit\s*$') {
        throw "$path must keep Option Explicit."
    }
    if ($text -notmatch '(?im)^\s*Dim\s+[^\r\n]*\bpowershellExe\b') {
        throw "$path uses powershellExe but does not declare it."
    }
    if ($text -notmatch '%WINDIR%\\System32\\WindowsPowerShell\\v1\.0\\powershell\.exe') {
        throw "$path does not use the absolute Windows PowerShell path."
    }
}

Write-Host 'PASS  VBS launchers are BOM-free and declare all added launcher variables.'
