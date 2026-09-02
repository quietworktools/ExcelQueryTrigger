    foreach ($requiredText in @(
            '$Shared.StartupUiReady = $true',
            'Final check - preparing the Dashboard'
        )) {
        if ($managerText -notmatch [regex]::Escape($requiredText)) {
            Add-Failure ('The Dashboard final startup gate is missing: {0}' -f $requiredText)
        }
    }

    # v1.0.3 replaced the Running-only startup gate with a single readiness
    # contract shared by the Dashboard and the engine. 'Waiting for network' is
    # an operational state: a laptop started away from the corporate network
    # must finish startup and show the Dashboard instead of sitting on the
    # splash screen until the share comes back.
    if ($commonText -notmatch '(?m)^\s*function\s+Test-EngineStatusReady\b') {
        Add-Failure 'src/Common.ps1 does not define Test-EngineStatusReady, which is the centralized startup readiness contract.'
    }
    else {
        # Exercise the real helper rather than pattern-matching its body, so a
        # change to the accepted-status list is caught rather than missed.
        #
        # Run it in a child scope: this block sits before the script's own
        # ". $commonPath", and dot-sourcing here would otherwise apply
        # Common.ps1's Set-StrictMode -Version 1.0 to every later check in this
        # file. The & { } keeps that, and the functions it defines, local.
        $readyProblems = @()
        try {
            $readyProblems = & {
                param($CommonScriptPath)
                . $CommonScriptPath
                $problems = @()
                foreach ($state in @('Running', 'Waiting for network')) {
                    if (-not (Test-EngineStatusReady $state)) {
                        $problems += ("'{0}' must be accepted as an operational startup state" -f $state)
                    }
                }
                foreach ($state in @('Starting', 'Stopping', 'Paused', 'Refreshing', 'Degraded', '')) {
                    if (Test-EngineStatusReady $state) {
                        $problems += ("'{0}' must not be accepted as an operational startup state" -f $state)
                    }
                }
                if (Test-EngineStatusReady $null) {
                    $problems += '$null must not be accepted as an operational startup state'
                }
                return ,$problems
            } (Join-Path $Root 'src\Common.ps1')
        }
        catch {
            $readyProblems = @(('Test-EngineStatusReady could not be evaluated: {0}' -f $_.Exception.Message))
        }

        foreach ($readyProblem in @($readyProblems)) {
            Add-Failure ('Test-EngineStatusReady contract is wrong: {0}' -f $readyProblem)
        }
    }

    if ($managerText -notmatch 'Test-EngineStatusReady\s*\(\[string\]\$Shared\.Status\)') {
        Add-Failure 'src/UIManager.ps1 does not gate Dashboard readiness on Test-EngineStatusReady.'
    }

    if ($engineText -notmatch 'Test-EngineStatusReady\s*\(\[string\]\$[Ss]hared\.Status\)') {
        Add-Failure 'src/Engine.ps1 does not gate engine startup readiness on Test-EngineStatusReady.'
    }

    # Regression guard replacing the old literal requirement: instead of
    # requiring a Running comparison, forbid one.
    foreach ($startupGateSource in @(
            @{ Name = 'src/UIManager.ps1'; Text = $managerText },
            @{ Name = 'src/Engine.ps1';    Text = $engineText }
        )) {
        if ($startupGateSource.Text -match '\$[Ss]hared\.Status\s*-(ne|eq)\s*''Running''') {
            Add-Failure ('{0} compares Status directly to Running for startup readiness. That is the pre-network-aware gate and would leave a machine started away from the corporate network waiting on the splash screen.' -f $startupGateSource.Name)
        }
    }
