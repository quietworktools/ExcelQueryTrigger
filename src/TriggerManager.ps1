# ==============================================================================
#  TriggerManager.ps1  (Engine runspace only)
#
#  Watchers are registered with Register-ObjectEvent WITHOUT -Action on purpose.
#  A -Action script block has to be executed inside a runspace; the engine
#  runspace is busy running the worker loop (and may be inside a 5 minute Excel
#  refresh), so an -Action handler would either block or be dropped. Without
#  -Action the events are pushed into the runspace event queue by the raising
#  thread and the worker loop drains them with Get-Event whenever it comes back
#  around. Nothing is lost while Excel is busy.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:Watchers      = @{}   # ruleId -> watcher entry
$script:TriggerState  = @{}   # ruleId -> debounce / cooldown bookkeeping
$script:TriggerShared = $null

# ---- network monitoring state ------------------------------------------------
# Online          - adapters up and every watched network folder answered
# NetworkOffline  - Windows reports no usable adapter (Wi-Fi off, lid closed,
#                   docking station unplugged)
# CheckingPaths   - adapters came back; waiting a moment before believing it
# WaitingForPaths - there is a network, but the corporate folders are not on it
#                   (home Wi-Fi, hotel, VPN not connected yet)
#
# NetworkOffline is the only global verdict. Everything else is decided per
# watched folder: one dead file server must not stop the monitors that point at
# a different one.
$script:NetworkState          = 'Online'
$script:NetworkStateSince     = [DateTime]::MinValue
$script:NetworkDetail         = ''
$script:NetworkPathKindCache  = @{}
# folderKey -> @{ Folder; Reachable; LastProbe }
$script:NetworkFolderState    = @{}
# How long to let the adapters settle before touching a share. Reconnecting
# Wi-Fi reports "available" before DNS and SMB are usable.
$script:NetworkStabilizeSeconds = 4
# How often to re-test the corporate folders while waiting for them. Long
# enough that an evening at home costs a handful of quiet probes.
$script:NetworkPathRetrySeconds = 45
# Windows can briefly report "no network" while switching from Wi-Fi to
# Ethernet (or between VPN routes) even though connectivity is restored almost
# immediately. Do not expose that transient handover as "Network lost".
#
# NetworkAdapterDownSince  - when the adapter first reported unavailable, or
#                            $null while Windows reports an adapter.
# NetworkAdapterGraceActive- $true only while an unavailable adapter is being
#                            deliberately hidden from the state machine, i.e.
#                            monitoring is still running. It is cleared the
#                            moment the outage is confirmed, so the "handover
#                            completed" note can never claim that monitoring
#                            survived when it was in fact suspended.
$script:NetworkAdapterDownSince = $null
$script:NetworkAdapterGraceActive = $false
$script:NetworkAdapterDownConfirmSeconds = 5

function Initialize-TriggerManager {
    param([Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:TriggerShared = $Shared
    $script:Watchers     = @{}
    $script:TriggerState = @{}
    $script:NetworkState         = 'Online'
    $script:NetworkStateSince    = Get-Date
    $script:NetworkDetail        = ''
    $script:NetworkPathKindCache = @{}
    $script:NetworkFolderState   = @{}
    $script:NetworkAdapterDownSince = $null
    $script:NetworkAdapterGraceActive = $false
}

# ------------------------------------------------------------------------------
# Region: network awareness
#
# Why this polls instead of subscribing to NetworkChange:
# NetworkAvailabilityChanged is a static event, and the only way to receive one
# in Windows PowerShell 5.1 is a script block that has to execute inside a
# runspace. This runspace is the engine loop, and it is regularly parked inside
# a multi-minute Excel COM call - the same reason the watchers themselves are
# registered without -Action further down this file. A handler would therefore
# fire late, or not at all, exactly when the machine is busiest.
# GetIsNetworkAvailable() reads the same adapter state, costs nothing, and is
# sampled by the loop that actually owns the watchers. No subscription exists,
# so there is nothing to duplicate and nothing to dispose at shutdown.
# ------------------------------------------------------------------------------

function Test-PathIsNetworkLocation {
    <#  UNC path, or a drive letter that is mapped to one.  #>
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $key = $Path.ToLowerInvariant()
    if ($script:NetworkPathKindCache.ContainsKey($key)) { return [bool]$script:NetworkPathKindCache[$key] }

    $isNetwork = $false
    try {
        if ($Path.StartsWith('\\')) {
            $isNetwork = $true
        }
        else {
            $root = [System.IO.Path]::GetPathRoot($Path)
            if (-not [string]::IsNullOrWhiteSpace($root)) {
                $drive = New-Object System.IO.DriveInfo($root)
                $isNetwork = ($drive.DriveType -eq [System.IO.DriveType]::Network)
            }
        }
    }
    catch { $isNetwork = $false }

    $script:NetworkPathKindCache[$key] = $isNetwork
    return $isNetwork
}

function Test-RuleUsesNetworkPath {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    if (-not (Test-TriggerUsesWatcher $Rule.trigger.type)) { return $false }
    return (Test-PathIsNetworkLocation (Get-WatchFolderForRule -Rule $Rule))
}

function Test-NetworkAdapterAvailable {
    <#
        Does Windows believe there is any usable adapter? This is only a hint:
        a VPN client, Hyper-V switch or docking-station NIC can keep it true
        with no route to the office at all. It is here because it turns "Wi-Fi
        is off" into an instant answer; the path probe decides everything else.
        Kept as its own function so it can be replaced in a test harness.
    #>
    try { return [bool][System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable() }
    catch { return $true }   # fail open: never suspend monitoring on a guess
}

function Get-NetworkFolderKey {
    param([string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Folder)) { return '' }
    return $Folder.TrimEnd('\\').ToLowerInvariant()
}

function Get-NetworkFolderEntry {
    param([string]$Folder)
    $key = Get-NetworkFolderKey $Folder
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    if (-not $script:NetworkFolderState.ContainsKey($key)) {
        $script:NetworkFolderState[$key] = @{
            Folder    = $Folder
            Reachable = $true
            LastProbe = [DateTime]::MinValue
        }
    }
    return $script:NetworkFolderState[$key]
}

function Get-NetworkMonitoringState {
    <#
        The overall verdict for the status header. Waiting is reported when any
        single watched folder is waiting, but that says nothing about the ones
        that are still working.
    #>
    return @{ State = [string]$script:NetworkState; Detail = [string]$script:NetworkDetail }
}

function Get-UnreachableNetworkFolders {
    <#
        The monitored folders currently marked unreachable by the per-folder
        logic. Read-only: it inspects the existing NetworkFolderState entries
        and changes nothing, so it cannot affect suspension, recovery or the
        retry interval. Used only to decide whether a handover message would
        be truthful.
    #>
    $folders = @()
    foreach ($key in @($script:NetworkFolderState.Keys)) {
        $entry = $script:NetworkFolderState[$key]
        if ($null -ne $entry -and -not [bool]$entry.Reachable) { $folders += [string]$entry.Folder }
    }
    return $folders
}

function Test-NetworkAdapterOffline {
    return ($script:NetworkState -eq 'NetworkOffline')
}

function Test-NetworkAdapterHandoverActive {
    <#
        True only inside the confirmation window, while Windows reports no
        adapter and the application is deliberately keeping monitoring alive.
        A path probe taken in this window proves nothing: the share may be
        unreachable purely because the route is being rebuilt. Callers use this
        to postpone a verdict, never to suppress one - if the outage is real,
        the state machine commits to NetworkOffline within a few seconds and
        suspends the watchers itself.
    #>
    return [bool]$script:NetworkAdapterGraceActive
}

function Test-RuleNetworkSuspended {
    <#
        Is this one rule deliberately not being watched because of the network?
        Everything that used to ask the global state asks this instead, so a
        rule pointing at a healthy server keeps running while another waits.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Rule)

    if (-not (Test-RuleUsesNetworkPath -Rule $Rule)) { return $false }
    if ($script:NetworkState -eq 'NetworkOffline') { return $true }
    $entry = Get-NetworkFolderEntry (Get-WatchFolderForRule -Rule $Rule)
    if ($null -eq $entry) { return $false }
    return (-not [bool]$entry.Reachable)
}

function Get-RuleNetworkWaitingStatus {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    if ($script:NetworkState -eq 'NetworkOffline') { return 'NetworkOffline' }
    return 'WaitingForNetwork'
}

function Set-NetworkStateInternal {
    param([string]$State, [string]$Detail)
    $script:NetworkState      = $State
    $script:NetworkStateSince = Get-Date
    $script:NetworkDetail     = $Detail
    if ($null -ne $script:TriggerShared) {
        $script:TriggerShared.NetworkState       = $State
        $script:TriggerShared.NetworkStateDetail = $Detail
    }
}

function Update-NetworkOverallState {
    <#  Rolls the per-folder verdicts up into the one word the header shows.  #>
    param([AllowNull()][array]$NetworkRules = @())

    if ($script:NetworkState -eq 'NetworkOffline' -or $script:NetworkState -eq 'CheckingPaths') { return }

    $waiting = New-Object System.Collections.ArrayList
    foreach ($rule in $NetworkRules) {
        $entry = Get-NetworkFolderEntry (Get-WatchFolderForRule -Rule $rule)
        if ($null -ne $entry -and -not [bool]$entry.Reachable) {
            if (-not $waiting.Contains([string]$entry.Folder)) { [void]$waiting.Add([string]$entry.Folder) }
        }
    }

    if ($waiting.Count -eq 0) {
        if ($script:NetworkState -ne 'Online') { Set-NetworkStateInternal -State 'Online' -Detail '' }
        return
    }

    Set-NetworkStateInternal -State 'WaitingForPaths' `
        -Detail ('The network is available, but {0} monitored location(s) cannot be reached: {1}' -f `
            $waiting.Count, ((@($waiting.ToArray()) -join ', ')))
}

function Suspend-NetworkRuleWatchers {
    <#
        Takes down network-backed watchers without the per-rule error trail. A
        watcher on a share that has gone is not a fault to report, it is a
        laptop that left the building. Limited to one folder when a folder is
        named, so an outage on one server leaves the others alone.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [string]$WatcherStatus = 'NetworkOffline',
        [string]$OnlyFolder = ''
    )

    $onlyKey   = Get-NetworkFolderKey $OnlyFolder
    $suspended = 0
    foreach ($rule in $Rules) {
        if (-not (Test-RuleUsesNetworkPath -Rule $rule)) { continue }
        if (-not [string]::IsNullOrWhiteSpace($onlyKey)) {
            if ((Get-NetworkFolderKey (Get-WatchFolderForRule -Rule $rule)) -ne $onlyKey) { continue }
        }
        $ruleId = [string]$rule.id
        if ($script:Watchers.ContainsKey($ruleId)) {
            Stop-RuleWatcher -RuleId $ruleId
            $suspended++
        }
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = $WatcherStatus } | Out-Null
    }
    return $suspended
}

function Invoke-NetworkPathProbe {
    <#
        Tests watched network folders and re-arms the watchers whose folder
        answered. -OnlyUnreachable is the quiet retry: folders that are already
        working are not touched, so a healthy server is never probed on behalf
        of a broken one.

        Returns @{ Probed; Recovered; Lost; Restarted; Waiting }.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [switch]$OnlyUnreachable,
        [int]$RetrySeconds = 0
    )

    $now       = Get-Date
    $probed    = 0
    $restarted = 0
    $recovered = New-Object System.Collections.ArrayList
    $lost      = New-Object System.Collections.ArrayList
    $waiting   = New-Object System.Collections.ArrayList
    $results   = @{}

    foreach ($rule in $Rules) {
        if (-not (Test-RuleUsesNetworkPath -Rule $rule)) { continue }
        $folder = Get-WatchFolderForRule -Rule $rule
        $entry  = Get-NetworkFolderEntry $folder
        if ($null -eq $entry) { continue }
        $key = Get-NetworkFolderKey $folder

        if (-not $results.ContainsKey($key)) {
            $skip = $false
            if ($OnlyUnreachable -and [bool]$entry.Reachable) { $skip = $true }
            if ((-not $skip) -and $RetrySeconds -gt 0 -and
                ($now - [DateTime]$entry.LastProbe).TotalSeconds -lt $RetrySeconds) { $skip = $true }

            if ($skip) {
                $results[$key] = [bool]$entry.Reachable
            }
            else {
                $wasReachable = [bool]$entry.Reachable
                $ok = $false
                try { $ok = Test-Path -LiteralPath $folder } catch { $ok = $false }
                $entry.Reachable = $ok
                $entry.LastProbe = $now
                $results[$key]   = $ok
                $probed++
                if ($ok -and -not $wasReachable) { [void]$recovered.Add($folder) }
                if ((-not $ok) -and $wasReachable) { [void]$lost.Add($folder) }
            }
        }

        if ($results[$key]) {
            if (-not $script:Watchers.ContainsKey([string]$rule.id)) {
                if (Start-RuleWatcher -Rule $rule -QuietUnavailable -IgnoreNetworkSuspension) { $restarted++ }
            }
        }
        else {
            if (-not $waiting.Contains($folder)) { [void]$waiting.Add($folder) }
            Set-RuleState -Shared $script:TriggerShared -RuleId ([string]$rule.id) `
                -Values @{ WatcherStatus = (Get-RuleNetworkWaitingStatus -Rule $rule) } | Out-Null
        }
    }

    return @{ Probed = $probed; Restarted = $restarted
              Recovered = @($recovered.ToArray()); Lost = @($lost.ToArray())
              Waiting = @($waiting.ToArray()) }
}

function Set-NetworkPathUnavailable {
    <#
        Called when one watched network folder stops answering while the
        adapters are still up - the home-Wi-Fi or one-dead-server case. Only
        that folder's watchers are suspended.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [Parameter(Mandatory = $true)][string]$Folder
    )

    $entry = Get-NetworkFolderEntry $Folder
    if ($null -eq $entry) { return }
    $alreadyWaiting = (-not [bool]$entry.Reachable)
    $entry.Reachable = $false
    $entry.LastProbe = Get-Date

    $suspended = Suspend-NetworkRuleWatchers -Rules $Rules -WatcherStatus 'WaitingForNetwork' -OnlyFolder $Folder
    Update-NetworkOverallState -NetworkRules $Rules

    if (-not $alreadyWaiting) {
        Write-AppLog -Level 'INFO' -Message `
            ('Network is available, but this monitored location is not reachable: {0}. {1} monitor(s) suspended for it; other locations keep watching.' -f $Folder, $suspended)
    }
}

function Request-NetworkPathRecheck {
    <#
        Lets the event intake tell the monitor that a folder just misbehaved,
        so the next network check probes it immediately instead of waiting for
        the retry interval.
    #>
    param([string]$Folder)
    $entry = Get-NetworkFolderEntry $Folder
    if ($null -ne $entry) { $entry.LastProbe = [DateTime]::MinValue }
}

function Write-NetworkProbeOutcome {
    <#  One line for what a probe changed, or nothing when it changed nothing. #>
    param([hashtable]$Probe)

    $recovered = @($Probe.Recovered)
    $waiting   = @($Probe.Waiting)

    if ($recovered.Count -gt 0) {
        if ($waiting.Count -gt 0) {
            Write-AppLog -Level 'INFO' -Message `
                ('Network location(s) available again: {0}. Restarted {1} watcher(s). Still waiting for: {2}' -f `
                    ($recovered -join ', '), [int]$Probe.Restarted, ($waiting -join ', '))
        }
        else {
            Write-AppLog -Level 'INFO' -Message `
                ('Network locations available. Restarted {0} watcher(s).' -f [int]$Probe.Restarted)
        }
        return
    }

    if (@($Probe.Lost).Count -gt 0) {
        Write-AppLog -Level 'INFO' -Message `
            ('Monitored network location(s) no longer reachable: {0}. Other locations keep watching.' -f (@($Probe.Lost) -join ', '))
        return
    }

    # Nothing changed. This is the state a laptop sits in all evening; it is not
    # news every 45 seconds.
    if ($waiting.Count -gt 0) {
        Write-AppLog -Level 'DEBUG' -Message ('Monitored network location(s) still unreachable: {0}' -f ($waiting -join ', ')) -NoActivity
    }
}

function Update-NetworkMonitoringState {
    <#
        The whole state machine, driven from the engine loop. The adapter check
        is free and runs every pass; anything that touches a share keeps to its
        own schedule, and only folders that are already known to be down are
        re-probed on the retry timer.
    #>
    param([AllowNull()][array]$Rules = @(), [bool]$Paused = $false)

    if ($null -eq $Rules) { $Rules = @() }
    $networkRules = @($Rules | Where-Object {
        (ConvertTo-BoolValue $_.enabled $true) -and (Test-RuleUsesNetworkPath -Rule $_)
    })

    # Nothing on the network to protect: never report a network problem.
    if ($networkRules.Count -eq 0) {
        if ($script:NetworkState -ne 'Online') { Set-NetworkStateInternal -State 'Online' -Detail '' }
        return
    }

    # Pausing is the user's decision and owns the watchers for as long as it
    # lasts. Do not fight it, and do not restart anything behind its back.
    if ($Paused) { return }

    # $adapterRaw is what Windows says right now. $adapterUp is that reading
    # after hysteresis, and is used ONLY where the answer would suspend
    # monitoring. The recovery branches below deliberately keep reading
    # $adapterRaw: smoothing a "the adapter came back" answer would let a
    # single flap promote the state machine to Online and probe shares over a
    # link that is not actually there.
    $adapterRaw = Test-NetworkAdapterAvailable
    $adapterUp  = $adapterRaw

    # Hysteresis for normal adapter handovers.
    # Windows may return $false for a few seconds while an Ethernet adapter
    # becomes preferred over Wi-Fi. Treat that as a transient transition, not
    # a real outage. Only a continuous outage longer than the confirmation
    # window is allowed to enter NetworkOffline.
    if (-not $adapterRaw) {
        if ($null -eq $script:NetworkAdapterDownSince) {
            $script:NetworkAdapterDownSince = Get-Date
            # Only claim a handover is being ridden out when monitoring is in
            # fact still running. From NetworkOffline there is nothing to hold.
            if ($script:NetworkState -eq 'Online' -or $script:NetworkState -eq 'WaitingForPaths') {
                $script:NetworkAdapterGraceActive = $true
                Write-AppLog -Level 'DEBUG' `
                    -Message ('Network adapter temporarily unavailable; holding monitoring active for up to {0}s in case this is an adapter handover.' -f $script:NetworkAdapterDownConfirmSeconds) -NoActivity
            }
        }

        $adapterDownSeconds = ((Get-Date) - $script:NetworkAdapterDownSince).TotalSeconds
        if ($adapterDownSeconds -lt $script:NetworkAdapterDownConfirmSeconds) {
            # For state-machine purposes the network remains provisionally up
            # during the grace window. Existing watchers are left untouched.
            $adapterUp = $true
        }
        else {
            # The window has expired. From here on this is a real outage and
            # the suspend branches below are allowed to act on it.
            $script:NetworkAdapterGraceActive = $false
        }
    }
    else {
        if ($null -ne $script:NetworkAdapterDownSince) {
            $briefDownSeconds = ((Get-Date) - $script:NetworkAdapterDownSince).TotalSeconds
            if ($script:NetworkAdapterGraceActive) {
                # The grace flag only records that the GLOBAL suspend was held
                # back. Individual folders can still have been suspended during
                # the same seconds - which is exactly what leaving the office
                # looks like: the adapter blinks, the corporate shares go away,
                # each one is suspended by the per-folder logic, and then some
                # other adapter reports itself available. Claiming "monitoring
                # was not suspended" there is simply untrue.
                #
                # When any monitored folder is currently unreachable, the
                # per-folder INFO messages already tell the whole story, so
                # nothing extra is written and the user is not told the same
                # thing twice.
                if (@(Get-UnreachableNetworkFolders).Count -eq 0) {
                    Write-AppLog -Level 'DEBUG' `
                        -Message ('Network adapter handover completed after {0:N1}s; monitoring was not suspended.' -f $briefDownSeconds) -NoActivity
                }
                else {
                    Write-AppLog -Level 'DEBUG' `
                        -Message ('Network adapter became available after {0:N1}s; monitored location availability is being re-evaluated.' -f $briefDownSeconds) -NoActivity
                }
            }
            $script:NetworkAdapterDownSince   = $null
            $script:NetworkAdapterGraceActive = $false
        }
    }

    switch ($script:NetworkState) {

        'Online' {
            if (-not $adapterUp) {
                Set-NetworkStateInternal -State 'NetworkOffline' `
                    -Detail 'No network connection. Monitoring of network folders is suspended and resumes by itself.'
                foreach ($rule in $networkRules) {
                    $entry = Get-NetworkFolderEntry (Get-WatchFolderForRule -Rule $rule)
                    if ($null -ne $entry) { $entry.Reachable = $false }
                }
                $suspended = Suspend-NetworkRuleWatchers -Rules $networkRules -WatcherStatus 'NetworkOffline'
                $script:NetworkAdapterGraceActive = $false
                Write-AppLog -Level 'INFO' -Message ('Network unavailable for more than {0} seconds. File monitoring suspended ({1} monitor(s)).' -f `
                    $script:NetworkAdapterDownConfirmSeconds, $suspended)
            }
        }

        'WaitingForPaths' {
            if (-not $adapterUp) {
                Set-NetworkStateInternal -State 'NetworkOffline' `
                    -Detail 'No network connection. Monitoring of network folders is suspended and resumes by itself.'
                foreach ($rule in $networkRules) {
                    $entry = Get-NetworkFolderEntry (Get-WatchFolderForRule -Rule $rule)
                    if ($null -ne $entry) { $entry.Reachable = $false }
                }
                $suspended = Suspend-NetworkRuleWatchers -Rules $networkRules -WatcherStatus 'NetworkOffline'
                $script:NetworkAdapterGraceActive = $false
                Write-AppLog -Level 'INFO' -Message ('Network unavailable for more than {0} seconds. File monitoring suspended ({1} monitor(s)).' -f `
                    $script:NetworkAdapterDownConfirmSeconds, $suspended)
            }
            else {
                # Only the folders that are down, and only on the retry timer.
                $probe = Invoke-NetworkPathProbe -Rules $networkRules -OnlyUnreachable -RetrySeconds $script:NetworkPathRetrySeconds
                if ([int]$probe.Probed -gt 0) {
                    Update-NetworkOverallState -NetworkRules $networkRules
                    Write-NetworkProbeOutcome -Probe $probe
                }
            }
        }

        'NetworkOffline' {
            # Raw, not smoothed: leaving an outage must be evidence-based.
            if ($adapterRaw) {
                Set-NetworkStateInternal -State 'CheckingPaths' `
                    -Detail 'Network connection restored. Checking the monitored locations.'
                Write-AppLog -Level 'INFO' -Message 'Network connection restored. Checking monitored locations.'
            }
        }

        'CheckingPaths' {
            # Raw, not smoothed. The stabilize wait (4s) is shorter than the
            # handover window (5s), so a smoothed reading here would let a
            # flapping link pass the stabilize gate, probe every share over a
            # dead route and report them all as unreachable. Nothing is
            # suspended by going back: the watchers are already down.
            if (-not $adapterRaw) {
                # Flapping Wi-Fi. Go back quietly; the loss was already logged.
                Set-NetworkStateInternal -State 'NetworkOffline' `
                    -Detail 'No network connection. Monitoring of network folders is suspended and resumes by itself.'
                Write-AppLog -Level 'DEBUG' -Message 'Network dropped again while checking monitored locations.' -NoActivity
            }
            elseif (((Get-Date) - $script:NetworkStateSince).TotalSeconds -ge $script:NetworkStabilizeSeconds) {
                # Everything gets one probe here: this is the moment the whole
                # picture can have changed.
                $probe = Invoke-NetworkPathProbe -Rules $networkRules
                Set-NetworkStateInternal -State 'Online' -Detail ''
                Update-NetworkOverallState -NetworkRules $networkRules
                if (@($probe.Waiting).Count -gt 0) {
                    Write-AppLog -Level 'INFO' -Message `
                        ('Restarted {0} watcher(s). Monitored location(s) still unreachable: {1}' -f `
                            [int]$probe.Restarted, ((@($probe.Waiting)) -join ', '))
                }
                else {
                    Write-AppLog -Level 'INFO' -Message ('Network locations available. Restarted {0} watcher(s).' -f [int]$probe.Restarted)
                }
            }
        }
    }
}

# ------------------------------------------------------------------------------
# Region: watcher lifecycle
# ------------------------------------------------------------------------------

function Get-WatchFolderForRule {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    $trigger = $Rule.trigger
    if (Test-TriggerUsesFolder $trigger.type) { return [string]$trigger.path }
    if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return '' }
    return (Split-Path -Parent ([string]$trigger.path))
}

function Get-WatchFilterForRule {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    $trigger = $Rule.trigger
    if (Test-TriggerUsesFolder $trigger.type) {
        if ([string]::IsNullOrWhiteSpace([string]$trigger.filter)) { return '*.*' }
        $configuredFilter = [string]$trigger.filter
        # FileSystemWatcher on Windows PowerShell 5.1 supports one Filter only.
        # For a preset such as Excel files we therefore watch all filenames and
        # apply the small set of wildcard patterns after Windows raises an event.
        # This remains event-driven; it does not enumerate or poll the folder.
        if ($configuredFilter.Contains(';')) { return '*.*' }
        return $configuredFilter
    }
    if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return '*.*' }
    return (Split-Path -Leaf ([string]$trigger.path))
}

function Start-RuleWatcher {
    <#  Creates and arms one FileSystemWatcher for a rule. Idempotent.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [switch]$QuietUnavailable,
        # Set by the network monitor itself, which has just proved the folder
        # answers and is the one allowed to bring these watchers back.
        [switch]$IgnoreNetworkSuspension
    )

    $ruleId = [string]$Rule.id
    Stop-RuleWatcher -RuleId $ruleId

    if (-not (Test-TriggerUsesWatcher $Rule.trigger.type)) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Manual' } | Out-Null
        return $true
    }

    # While the network is down, starting a watcher on a share can only fail,
    # and it would fail once per rule per health check. Say nothing and wait.
    if ((-not $IgnoreNetworkSuspension) -and (Test-RuleNetworkSuspended -Rule $Rule)) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId `
            -Values @{ WatcherStatus = (Get-RuleNetworkWaitingStatus -Rule $Rule) } | Out-Null
        return $false
    }

    $folder = Get-WatchFolderForRule -Rule $Rule
    if ([string]::IsNullOrWhiteSpace($folder)) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Misconfigured' } | Out-Null
        Write-AppLog -Level 'WARN' -RuleName $Rule.name -Message 'Watcher not started: no folder configured.'
        return $false
    }

    $reachable = $false
    try { $reachable = Test-Path -LiteralPath $folder } catch { $reachable = $false }
    if (-not $reachable) {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'PathUnavailable' } | Out-Null
        if (-not $QuietUnavailable) {
            Write-AppLog -Level 'WARN' -RuleName $Rule.name -ErrorType 'PathUnavailable' `
                -Message ('Watch folder is not reachable, will retry on the next health check: {0}' -f $folder)
        }
        return $false
    }

    try {
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path                  = $folder
        $watcher.Filter                = (Get-WatchFilterForRule -Rule $Rule)
        $watcher.IncludeSubdirectories = $false
        # 64 KB buffer: network shares can burst several events for one copy.
        $watcher.InternalBufferSize    = 65536
        $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor
                                [System.IO.NotifyFilters]::LastWrite -bor
                                [System.IO.NotifyFilters]::Size

        $eventNames = @('Error')
        switch ($Rule.trigger.type) {
            'FileCreated'         { $eventNames += @('Created', 'Renamed') }
            'FileChangedAny'      { $eventNames += @('Changed') }
            'FileChangedSpecific' { $eventNames += @('Changed', 'Created', 'Renamed') }
        }

        $subscriptions = New-Object System.Collections.ArrayList
        foreach ($eventName in $eventNames) {
            $sourceIdentifier = 'EQT:{0}:{1}' -f $ruleId, $eventName
            $subscription = Register-ObjectEvent -InputObject $watcher -EventName $eventName `
                -SourceIdentifier $sourceIdentifier -MessageData $ruleId
            [void]$subscriptions.Add($sourceIdentifier)
            $subscription | Out-Null
        }

        $watcher.EnableRaisingEvents = $true

        $script:Watchers[$ruleId] = @{
            Watcher       = $watcher
            Subscriptions = @($subscriptions.ToArray())
            Folder        = $folder
            RuleName      = [string]$Rule.name
            LastError     = ''
            StartedAt     = (Get-Date)
        }

        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Active' } | Out-Null
        # A network recovery already reports itself in one line for the whole
        # application, so the per-rule detail goes to the debug log instead of
        # repeating the story once per rule.
        Write-AppLog -Level $(if ($IgnoreNetworkSuspension) { 'DEBUG' } else { 'INFO' }) -RuleName $Rule.name `
            -Message ('Watcher started. Folder={0} Filter={1} Events={2}' -f $folder, $watcher.Filter, (($eventNames | Where-Object { $_ -ne 'Error' }) -join ','))
        return $true
    }
    catch {
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Error' } | Out-Null
        Write-AppLog -Level 'ERROR' -RuleName $Rule.name -ErrorType 'WatcherError' `
            -Message ('Failed to start watcher: {0}' -f $_.Exception.Message)
        return $false
    }
}

function Stop-RuleWatcher {
    param([Parameter(Mandatory = $true)][string]$RuleId)

    if (-not $script:Watchers.ContainsKey($RuleId)) { return }
    $entry = $script:Watchers[$RuleId]

    try {
        if ($null -ne $entry.Watcher) {
            $entry.Watcher.EnableRaisingEvents = $false
        }
    }
    catch { }

    foreach ($sourceIdentifier in @($entry.Subscriptions)) {
        try { Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue } catch { }
        try {
            Get-Event -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue |
                ForEach-Object { Remove-Event -EventIdentifier $_.EventIdentifier -ErrorAction SilentlyContinue }
        }
        catch { }
    }

    try { if ($null -ne $entry.Watcher) { $entry.Watcher.Dispose() } } catch { }
    $script:Watchers.Remove($RuleId)
}

function Restart-RuleWatcher {
    param([Parameter(Mandatory = $true)][hashtable]$Rule)
    Write-AppLog -Level 'INFO' -RuleName $Rule.name -Message 'Restarting watcher.'
    return (Start-RuleWatcher -Rule $Rule)
}

function Stop-AllRuleWatchers {
    foreach ($ruleId in @($script:Watchers.Keys)) { Stop-RuleWatcher -RuleId $ruleId }
}

function Sync-RuleWatchers {
    <#
        Brings the running watchers in line with the configuration. Called at
        startup, after every config change and after resume - so enabling or
        editing a rule never needs an application restart.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [bool]$Paused = $false,
        [switch]$QuietUnavailable
    )

    if ($null -eq $Rules) { $Rules = @() }
    $wanted = @{}
    foreach ($rule in $Rules) {
        if ((ConvertTo-BoolValue $rule.enabled $true) -and (Test-TriggerUsesWatcher $rule.trigger.type) -and -not $Paused) {
            $wanted[[string]$rule.id] = $rule
        }
    }

    foreach ($ruleId in @($script:Watchers.Keys)) {
        if (-not $wanted.ContainsKey($ruleId)) {
            Stop-RuleWatcher -RuleId $ruleId
        }
    }

    $watcherNumber = 0
    foreach ($ruleId in @($wanted.Keys)) {
        $rule = $wanted[$ruleId]
        $script:TriggerShared.StartupMessage = ('Activating monitor: {0}' -f [string]$rule.name)
        if ($script:Watchers.ContainsKey($ruleId)) {
            # Re-arm if the folder/filter changed, or if an object is present
            # but is no longer actually raising events. Sync is used by the
            # startup readiness gate, so it must validate as well as create.
            $entry = $script:Watchers[$ruleId]
            $watcherHealthy = $false
            try {
                $sameFolder = ($entry.Folder -eq (Get-WatchFolderForRule -Rule $rule))
                $sameFilter = ($entry.Watcher.Filter -eq (Get-WatchFilterForRule -Rule $rule))
                $watcherHealthy = $sameFolder -and $sameFilter -and
                    [string]::IsNullOrWhiteSpace([string]$entry.LastError) -and
                    [bool]$entry.Watcher.EnableRaisingEvents
            }
            catch { $watcherHealthy = $false }

            if ($watcherHealthy) {
                $watcherNumber++
                $script:TriggerShared.StartupCurrent = $watcherNumber
                continue
            }

            Stop-RuleWatcher -RuleId $ruleId
        }
        Start-RuleWatcher -Rule $rule -QuietUnavailable:$QuietUnavailable | Out-Null
        $watcherNumber++
        $script:TriggerShared.StartupCurrent = $watcherNumber
    }

    foreach ($rule in $Rules) {
        $ruleId = [string]$rule.id
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Disabled' } | Out-Null
        }
        elseif ($Paused -and (Test-TriggerUsesWatcher $rule.trigger.type)) {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Paused' } | Out-Null
        }
        elseif (-not (Test-TriggerUsesWatcher $rule.trigger.type)) {
            # Scheduled / logon / manual rules have no watcher, but the status
            # column should still say what the rule is waiting for.
            $status = switch ([string]$rule.trigger.type) {
                'Scheduled' { 'Waiting for time' }
                'Logon'     { 'Waiting for logon' }
                default     { 'Manual only' }
            }
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = $status } | Out-Null
        }
    }
}

# ------------------------------------------------------------------------------
# Region: event intake, filtering, debounce, cooldown
# ------------------------------------------------------------------------------

function Test-TriggerFileMatch {
    <#  Applies the conditions the FileSystemWatcher filter cannot express.  #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$FullPath
    )

    $trigger = $Rule.trigger
    $leaf = Split-Path -Leaf $FullPath

    if ($trigger.type -eq 'FileChangedSpecific') {
        if ([string]::IsNullOrWhiteSpace([string]$trigger.path)) { return $false }
        return ([string]::Equals([System.IO.Path]::GetFullPath($FullPath), [System.IO.Path]::GetFullPath([string]$trigger.path), [StringComparison]::OrdinalIgnoreCase))
    }

    # When a rule contains several wildcard patterns (for example the Excel
    # preset), FileSystemWatcher itself is deliberately broad (*.*) and this
    # inexpensive filename test narrows the event back down here.
    if (Test-TriggerUsesFolder $trigger.type) {
        $configuredFilter = [string]$trigger.filter
        if (-not [string]::IsNullOrWhiteSpace($configuredFilter) -and $configuredFilter.Contains(';')) {
            $patternMatched = $false
            foreach ($pattern in $configuredFilter.Split(';')) {
                $trimmed = $pattern.Trim()
                if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
                if ($leaf -like $trimmed) { $patternMatched = $true; break }
            }
            if (-not $patternMatched) { return $false }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$trigger.contains)) {
        if ($leaf.IndexOf([string]$trigger.contains, [StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$trigger.exclude)) {
        if ($leaf.IndexOf([string]$trigger.exclude, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $false }
    }

    # Temporary artefacts of a copy in progress are never a trigger.
    if ($leaf.StartsWith('~$') -or $leaf.EndsWith('.tmp') -or $leaf.EndsWith('.crdownload') -or $leaf.EndsWith('.part')) {
        return $false
    }

    return $true
}

function Read-WatcherEvents {
    <#
        Drains the runspace event queue. Returns an array of
        @{ RuleId; FullPath; ChangeType; EventName } and records watcher errors.
    #>
    $results = New-Object System.Collections.ArrayList

    $queued = @()
    try { $queued = @(Get-Event -ErrorAction SilentlyContinue) } catch { $queued = @() }

    foreach ($queuedEvent in $queued) {
        $identifier = [string]$queuedEvent.SourceIdentifier
        if (-not $identifier.StartsWith('EQT:')) { continue }

        try {
            $parts     = $identifier.Split(':')
            $ruleId    = [string]$queuedEvent.MessageData
            if ([string]::IsNullOrWhiteSpace($ruleId) -and $parts.Count -ge 2) { $ruleId = $parts[1] }
            $eventName = $parts[$parts.Count - 1]

            if ($eventName -eq 'Error') {
                $message       = 'Unknown watcher error'
                $exceptionType = ''
                try {
                    $watcherException = $queuedEvent.SourceArgs[1].GetException()
                    $message = $watcherException.Message
                    $exceptionType = $watcherException.GetType().FullName
                }
                catch { }
                if ($script:Watchers.ContainsKey($ruleId)) { $script:Watchers[$ruleId].LastError = $message }
                [void]$results.Add(@{
                    RuleId       = $ruleId
                    EventName    = 'Error'
                    Message      = $message
                    ExceptionType = $exceptionType
                    FullPath     = ''
                    ChangeType   = 'Error'
                })
            }
            else {
                $eventArgs = $queuedEvent.SourceArgs[1]
                [void]$results.Add(@{
                    RuleId     = $ruleId
                    EventName  = $eventName
                    FullPath   = [string]$eventArgs.FullPath
                    ChangeType = [string]$eventArgs.ChangeType
                    Message    = ''
                })
            }
        }
        catch {
            Write-AppLog -Level 'DEBUG' -Message ('Malformed watcher event ignored: {0}' -f $_.Exception.Message)
        }
        finally {
            try { Remove-Event -EventIdentifier $queuedEvent.EventIdentifier -ErrorAction SilentlyContinue } catch { }
        }
    }

    $events = @($results.ToArray())
    return ,$events
}

function New-TriggerStateEntry {
    <#
        One rule's debounce / cooldown bookkeeping. Every field here exists
        because a current caller reads or writes it; nothing is speculative.

          PendingFiles  Hashtable, lowercased full path -> original full path.
                        Register-TriggerEvent uses ContainsKey/indexer to
                        de-duplicate repeats of the same file inside one
                        window; Get-DueTriggers reads .Values and .Count and
                        resets it with an empty hashtable literal. A list or
                        set would not satisfy both.
          PendingSince  $null when no window is open. Set to the event time
                        when a window opens, and put back to $null (not
                        MinValue) by Get-DueTriggers on dispatch - the reset
                        is what fixes the unset representation.
          LastEventAt   DateTime. Get-DueTriggers subtracts it from now, so it
                        must never be $null. MinValue is safe: the subtraction
                        yields a huge span, and the branch is reached only
                        after PendingFiles.Count > 0 proves a real event set it.
          EventCount    [int]. Incremented per accepted event, zeroed both when
                        a window opens and after dispatch.
          CooldownUntil DateTime, compared with "$now -lt". MinValue means "no
                        cooldown", because every real time is greater than it.
                        $null would make the comparison meaningless.
          PrimaryFile   [string], the file named in the "Trigger detected" log
                        line - the first event that opened the window. Hashtable
                        enumeration order is not insertion order, so PendingFiles
                        cannot answer "which file was first". Without this, the
                        representative file shown in the approval prompt, the
                        Current Job panel and the history triggerFile could name
                        a different file than the one the user was told about.
                        '' when no window is open, matching the existing
                        empty-string fallback in Get-DueTriggers.
    #>
    return @{
        PendingFiles  = @{}
        PendingSince  = $null
        LastEventAt   = [DateTime]::MinValue
        EventCount    = 0
        CooldownUntil = [DateTime]::MinValue
        PrimaryFile   = ''
    }
}

function Get-TriggerState {
    <#
        The per-rule debounce/cooldown state, created on first use.

        Returns the stored hashtable itself, not a copy: Register-TriggerEvent
        accumulates into the same instance that Get-DueTriggers later drains,
        which is the whole point of the accessor. Hashtables are reference
        types, so the caller's mutations are the stored state.

        Lifecycle is owned by Clear-TriggerState, which already exists: it
        resets everything on Resume and can drop a single rule by id. This
        function deliberately adds no lifecycle of its own - it only creates
        the entry it was asked for.
    #>
    param([Parameter(Mandatory = $true)][string]$RuleId)

    if (-not $script:TriggerState.ContainsKey($RuleId)) {
        $script:TriggerState[$RuleId] = New-TriggerStateEntry
    }
    return $script:TriggerState[$RuleId]
}

function Register-TriggerEvent {
    <#
        Records one matching filesystem event against the rule's debounce
        window. Returns $true when the event was accepted, $false when it was
        dropped by a filter or by cooldown.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Rule,
        [Parameter(Mandatory = $true)][string]$FullPath,
        [string]$ChangeType = 'Changed'
    )

    $ruleId = [string]$Rule.id
    $state  = Get-TriggerState -RuleId $ruleId
    $now    = Get-Date

    if (-not (Test-TriggerFileMatch -Rule $Rule -FullPath $FullPath)) {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name -Message ('Event ignored by filename condition: {0}' -f $FullPath)
        return $false
    }

    if ($now -lt $state.CooldownUntil) {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name `
            -Message ('Event dropped by cooldown (until {0:HH:mm:ss}): {1}' -f $state.CooldownUntil, $FullPath)
        return $false
    }

    if ($state.PendingFiles.Count -eq 0) {
        $state.PendingSince = $now
        $state.EventCount   = 0
        # Set alongside the log line below, so the file recorded here is by
        # construction the same file the user is told about.
        $state.PrimaryFile  = $FullPath
        Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ LastTrigger = $now } | Out-Null
        Write-AppLog -Level 'INFO' -RuleName $Rule.name -Stage 'TriggerDetected' `
            -Message ('Trigger detected ({0}): {1}' -f $ChangeType, (Split-Path -Leaf $FullPath))
    }
    else {
        Write-AppLog -Level 'DEBUG' -RuleName $Rule.name `
            -Message ('Debounce window extended by {0} event: {1}' -f $ChangeType, (Split-Path -Leaf $FullPath))
    }

    $fileKey = $FullPath.ToLowerInvariant()
    if (-not $state.PendingFiles.ContainsKey($fileKey)) {
        $state.PendingFiles[$fileKey] = $FullPath
    }
    $state.LastEventAt = $now
    $state.EventCount  = [int]$state.EventCount + 1
    return $true
}

function Get-DueTriggers {
    <#
        Collapses everything accumulated inside the debounce window into a
        single fire per rule, and opens the cooldown window immediately so the
        tail of a noisy copy cannot queue a second job.
    #>
    param([AllowNull()][array]$Rules = @())

    $due = New-Object System.Collections.ArrayList
    $now = Get-Date

    foreach ($rule in $Rules) {
        $ruleId = [string]$rule.id
        if (-not $script:TriggerState.ContainsKey($ruleId)) { continue }

        $state = $script:TriggerState[$ruleId]
        if ($state.PendingFiles.Count -eq 0) { continue }

        $debounce = ConvertTo-IntValue $rule.trigger.debounceSeconds 5 0
        if (($now - $state.LastEventAt).TotalSeconds -lt $debounce) { continue }

        $filePaths = @($state.PendingFiles.Values)
        # The file that opened this window, not whichever one the hashtable
        # happens to enumerate first. Falls back to the old behaviour if
        # PrimaryFile is empty, so state written before this field existed
        # (or by any future caller that skips it) still dispatches.
        $filePath = [string]$state.PrimaryFile
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            $filePath = $(if ($filePaths.Count -gt 0) { [string]$filePaths[0] } else { '' })
        }
        $count    = $state.EventCount

        $state.PendingFiles  = @{}
        $state.PendingSince  = $null
        $state.EventCount    = 0
        $state.PrimaryFile   = ''
        $cooldown = ConvertTo-IntValue $rule.trigger.cooldownSeconds 30 0
        $state.CooldownUntil = $now.AddSeconds($cooldown)

        Write-AppLog -Level 'DEBUG' -RuleName $rule.name `
            -Message ('Debounce elapsed. {0} raw event(s) collapsed into one job. Cooldown until {1:HH:mm:ss}.' -f $count, $state.CooldownUntil)

        [void]$due.Add(@{ Rule = $rule; FilePath = $filePath; FilePaths = $filePaths; EventCount = $count })
    }

    $dueList = @($due.ToArray())
    return ,$dueList
}

function Get-TriggerStateRuleIds {
    <#
        The rule ids that currently hold debounce/cooldown state. Snapshotted
        into a plain array so the caller can Clear-TriggerState while looping
        without mutating the dictionary it is enumerating.

        Returned WITHOUT the ",$array" wrapper used elsewhere in this file: the
        caller writes @(Get-TriggerStateRuleIds), and wrapping a comma-returned
        array in @() nests it one level deep, which would silently make the
        prune loop iterate over a single Object[] and clear nothing.
    #>
    $ids = @($script:TriggerState.Keys | ForEach-Object { [string]$_ })
    return $ids
}

function Reset-TriggerPendingWindow {
    <#
        Drops the debounce window a rule has accumulated, without forgetting
        the rule. Used when the configuration is reloaded: a rule can keep its
        id while its folder, filter, type or event set changes, and files
        collected under the old definition must not be dispatched against the
        new one.

        CooldownUntil is deliberately kept. It records that a refresh recently
        ran, which is still true after a reload and has nothing to do with the
        trigger definition; clearing it would let a save in the Rule Editor
        re-open the door to an immediate second refresh of the same workbook.
        LastEventAt is left alone because it is only ever read while
        PendingFiles has entries, which this function empties.

        Returns $true when there was actually a pending window to drop.
    #>
    param([Parameter(Mandatory = $true)][string]$RuleId)

    if (-not $script:TriggerState.ContainsKey($RuleId)) { return $false }
    $state = $script:TriggerState[$RuleId]
    if ($state.PendingFiles.Count -eq 0) { return $false }

    $state.PendingFiles = @{}
    $state.PendingSince = $null
    $state.EventCount   = 0
    $state.PrimaryFile  = ''
    return $true
}

function Clear-TriggerState {
    param([string]$RuleId)
    if ([string]::IsNullOrWhiteSpace($RuleId)) { $script:TriggerState = @{}; return }
    if ($script:TriggerState.ContainsKey($RuleId)) { $script:TriggerState.Remove($RuleId) }
}

# ------------------------------------------------------------------------------
# Region: source file readiness
# ------------------------------------------------------------------------------

function Test-SourceFileReady {
    <#
        A file that has just appeared on a share is usually still being written.
        Two independent signals must agree before we let Excel read it:
          1. the size has stopped changing between two probes
          2. the file can be opened with FileShare.Read (i.e. nobody holds a
             write handle on it any more)

        Some source systems write a file under a temporary name and then rename
        it. The path we were handed then disappears for good, so waiting the
        whole timeout for it achieves nothing except delaying the next trigger.
        A file that has been absent for several consecutive probes is therefore
        reported as Vanished straight away.

        "The file is gone" and "the share is gone" look identical to Test-Path,
        and only the first one is safe to skip over. Absence is only counted
        while the folder that should contain the file is still answering, so a
        dropped SMB session is reported as Unreachable and keeps its job honest
        instead of being written off as a rename.

        Returns @{ Ready = [bool]; Vanished = [bool]; Unreachable = [bool];
                   Reason = [string]; WaitedSeconds = [int] }
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$IntervalSeconds = 2,
        [int]$TimeoutSeconds = 60,
        [int]$MissingProbesBeforeGivingUp = 3,
        [scriptblock]$ShouldAbort
    )

    $started      = Get-Date
    $deadline     = $started.AddSeconds($TimeoutSeconds)
    $lastSize     = -1L
    $lastWrite    = [DateTime]::MinValue
    $intervalMs   = [Math]::Max(250, $IntervalSeconds * 1000)
    $everSeen     = $false
    $missingCount = 0
    $folderDown   = $false
    # Recorded on every probe so a timeout can say which of the conditions was
    # never met. "Did not become ready" on its own is not diagnosable.
    $lastBlocker  = 'the file could not be read'

    $parentFolder = ''
    try { $parentFolder = [string](Split-Path -Parent $Path) } catch { $parentFolder = '' }

    while ($true) {
        if ($null -ne $ShouldAbort -and (& $ShouldAbort)) {
            return @{ Ready = $false; Vanished = $false; Unreachable = $false
                      Reason = 'Cancelled while waiting for the source file.'
                      WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
        }

        $info = $null
        try {
            if (Test-Path -LiteralPath $Path) { $info = Get-Item -LiteralPath $Path -ErrorAction Stop }
        }
        catch { $info = $null }

        if ($null -ne $info) {
            $everSeen     = $true
            $missingCount = 0
            $folderDown   = $false
            $stable = ($info.Length -eq $lastSize -and $info.LastWriteTimeUtc -eq $lastWrite -and $lastSize -ge 0)
            $lastSize  = $info.Length
            $lastWrite = $info.LastWriteTimeUtc

            if (-not $stable) {
                $lastBlocker = 'it was still being written (size or timestamp kept changing)'
            }
            elseif (Test-FileNotLocked -Path $Path) {
                return @{ Ready = $true; Vanished = $false; Unreachable = $false
                          Reason = 'Size stable and no writer holds the file.'
                          WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
            }
            else {
                $lastBlocker = 'another program still holds it open for writing'
            }
        }
        else {
            # Is the file gone, or the whole share? Only the first is a rename.
            $folderReachable = $false
            if (-not [string]::IsNullOrWhiteSpace($parentFolder)) {
                try { $folderReachable = Test-Path -LiteralPath $parentFolder } catch { $folderReachable = $false }
            }

            if ($folderReachable) {
                $folderDown = $false
                $missingCount++
                $lastBlocker = 'it is no longer present in the folder'
                if ($missingCount -ge $MissingProbesBeforeGivingUp) {
                    $reason = $(if ($everSeen) {
                        'Source file was renamed, replaced or deleted while waiting ({0})' -f $Path
                    } else {
                        'Source file was already gone when the job started, most likely renamed by the system that wrote it ({0})' -f $Path
                    })
                    return @{ Ready = $false; Vanished = $true; Unreachable = $false
                              Reason = $reason
                              WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
                }
            }
            else {
                # A dropped share is never treated as a rename. Keep waiting -
                # these outages have come back on their own - and if the time
                # runs out the job fails rather than refreshing on a guess.
                $folderDown   = $true
                $missingCount = 0
                $lastBlocker  = ('the folder that holds it could not be reached ({0})' -f $parentFolder)
            }
        }

        if ((Get-Date) -ge $deadline) {
            $reason = 'Source file did not become ready within {0} seconds - {1}' -f $TimeoutSeconds, $lastBlocker
            return @{ Ready = $false; Vanished = $false; Unreachable = $folderDown
                      Reason = $reason
                      WaitedSeconds = [int]((Get-Date) - $started).TotalSeconds }
        }

        Start-Sleep -Milliseconds $intervalMs
    }
}

function Test-FileNotLocked {
    param([Parameter(Mandatory = $true)][string]$Path)
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

# ------------------------------------------------------------------------------
# Region: watcher health
# ------------------------------------------------------------------------------

function Test-WatcherHealth {
    <#
        Cheap periodic check - no folder enumeration. Recreates watchers that
        died silently (the classic network share failure) and picks up shares
        that came back. Returns the number of watchers that were recreated.
    #>
    param(
        [AllowNull()][array]$Rules = @(),
        [bool]$Paused = $false
    )

    if ($Paused) { return 0 }
    if ($null -eq $Rules) { $Rules = @() }
    $recreated = 0

    foreach ($rule in $Rules) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (-not (Test-TriggerUsesWatcher $rule.trigger.type)) { continue }

        # The network monitor owns these rules until it says the network is
        # back. Checking them here is what produced a fault report per rule per
        # minute for a laptop that was simply somewhere else.
        $ruleUsesNetwork = Test-RuleUsesNetworkPath -Rule $rule
        if (Test-RuleNetworkSuspended -Rule $rule) { continue }

        $ruleId  = [string]$rule.id
        $folder  = Get-WatchFolderForRule -Rule $rule
        $healthy = $true
        $reason  = ''

        if (-not $script:Watchers.ContainsKey($ruleId)) {
            $healthy = $false
            $reason  = 'watcher missing'
        }
        else {
            $entry = $script:Watchers[$ruleId]
            if (-not [string]::IsNullOrWhiteSpace($entry.LastError)) {
                $healthy = $false
                $reason  = ('watcher reported an error: {0}' -f $entry.LastError)
            }
            else {
                try {
                    if (-not $entry.Watcher.EnableRaisingEvents) {
                        $healthy = $false
                        $reason  = 'EnableRaisingEvents is false'
                    }
                }
                catch {
                    $healthy = $false
                    $reason  = 'watcher object is no longer usable'
                }
            }
        }

        $reachable = $false
        try { $reachable = (-not [string]::IsNullOrWhiteSpace($folder)) -and (Test-Path -LiteralPath $folder) } catch { $reachable = $false }

        if (-not $reachable) {
            if ($ruleUsesNetwork) {
                # There is a network, but not this folder: home Wi-Fi, a hotel,
                # the VPN is not up yet, or one file server is down. One line
                # for that folder, then quiet retries - and every other watched
                # location carries on untouched.
                Set-NetworkPathUnavailable -Rules $Rules -Folder $folder
                continue
            }
            if ($script:Watchers.ContainsKey($ruleId)) {
                Write-AppLog -Level 'WARN' -RuleName $rule.name -ErrorType 'PathUnavailable' `
                    -Message ('Watch folder became unreachable, watcher disposed until it returns: {0}' -f $folder)
                Stop-RuleWatcher -RuleId $ruleId
            }
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'PathUnavailable' } | Out-Null
            continue
        }

        if (-not $healthy) {
            Write-AppLog -Level 'WARN' -RuleName $rule.name -ErrorType 'WatcherError' `
                -Message ('Watcher unhealthy ({0}). Recreating.' -f $reason)
            if (Restart-RuleWatcher -Rule $rule) { $recreated++ }
        }
        else {
            Set-RuleState -Shared $script:TriggerShared -RuleId $ruleId -Values @{ WatcherStatus = 'Active' } | Out-Null
        }
    }

    return $recreated
}

function Get-WatcherSummary {
    <#  Used by the engine to decide between Running and Degraded.  #>
    param([AllowNull()][array]$Rules = @())

    $expected  = 0
    $active    = 0
    $suspended = 0
    $inactive  = New-Object System.Collections.ArrayList
    foreach ($rule in $Rules) {
        if (-not (ConvertTo-BoolValue $rule.enabled $true)) { continue }
        if (-not (Test-TriggerUsesWatcher $rule.trigger.type)) { continue }

        # A watcher that is deliberately down because there is no network is not
        # a missing watcher. Counting it as one turns a train journey into a
        # Degraded application, and would hold the startup splash open forever.
        if (Test-RuleNetworkSuspended -Rule $rule) {
            $suspended++
            continue
        }

        $expected++
        $ruleId = [string]$rule.id
        $isActive = $false
        $watcherStatus = 'not started'

        if ($script:Watchers.ContainsKey($ruleId)) {
            $entry = $script:Watchers[$ruleId]
            try {
                $isActive = [string]::IsNullOrWhiteSpace([string]$entry.LastError) -and
                    [bool]$entry.Watcher.EnableRaisingEvents
                if (-not $isActive) {
                    $watcherStatus = $(if (-not [string]::IsNullOrWhiteSpace([string]$entry.LastError)) {
                        'watcher error: {0}' -f [string]$entry.LastError
                    } else {
                        'watcher is not raising events'
                    })
                }
            }
            catch {
                $isActive = $false
                $watcherStatus = 'watcher object unavailable'
            }
        }
        elseif ($null -ne $script:TriggerShared -and
            $script:TriggerShared.RuleState.ContainsKey($ruleId)) {
            $watcherStatus = [string]$script:TriggerShared.RuleState[$ruleId].WatcherStatus
        }

        if ($isActive) {
            $active++
            continue
        }

        if ($watcherStatus -eq 'PathUnavailable' -and
            -not [string]::IsNullOrWhiteSpace([string]$rule.trigger.path)) {
            $watcherStatus = 'path unavailable: {0}' -f [string]$rule.trigger.path
        }
        [void]$inactive.Add(('{0} ({1})' -f [string]$rule.name, $watcherStatus))
    }

    return @{ Expected = $expected; Active = $active; Suspended = $suspended; Inactive = @($inactive.ToArray()) }
}
