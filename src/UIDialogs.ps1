# ==============================================================================
#  UIDialogs.ps1
#  Common file / folder pickers for the rule editor.
#
#  Windows PowerShell 5.1 runs on the .NET Framework, whose
#  System.Windows.Forms.FolderBrowserDialog is still the old tree-view control:
#  no address bar, no Quick Access, and pasting a UNC path is painful. Excel,
#  Explorer and every modern application use the Vista-era IFileDialog with the
#  FOS_PICKFOLDERS option instead - the same window as "Open", but returning a
#  folder.
#
#  There is no managed wrapper for that dialog in the .NET Framework, so the COM
#  interface is declared here and compiled once per session. If compilation is
#  blocked (locked-down machines sometimes prevent it) the classic dialog is used
#  instead, so the picker is a cosmetic upgrade and never a hard dependency.
# ==============================================================================

Set-StrictMode -Version 1.0

$script:FolderPickerState = 'Unknown'   # Unknown | Modern | Classic
function Initialize-FolderPicker {
    <#
        .SYNOPSIS
            Compiles the IFileDialog wrapper once. Safe to call repeatedly.
        .OUTPUTS
            $true when the Explorer-style dialog is available.
    #>
    [CmdletBinding()]
    param()

    if ($script:FolderPickerState -ne 'Unknown') { return ($script:FolderPickerState -eq 'Modern') }

    if ('ExcelQueryTrigger.FolderPicker' -as [type]) {
        $script:FolderPickerState = 'Modern'
        return $true
    }

    $source = @'
using System;
using System.Runtime.InteropServices;

namespace ExcelQueryTrigger
{
    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileDialog
    {
        // --- IModalWindow -----------------------------------------------------
        [PreserveSig] int Show(IntPtr parent);
        // --- IFileDialog ------------------------------------------------------
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName(out IntPtr pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close([MarshalAs(UnmanagedType.Error)] int hr);
        void SetClientGuid([In] ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
    }

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    internal class FileOpenDialogCoClass { }

    public static class FolderPicker
    {
        private const uint FOS_NOCHANGEDIR    = 0x00000008;
        private const uint FOS_PICKFOLDERS    = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM= 0x00000040;
        private const uint FOS_PATHMUSTEXIST  = 0x00000800;
        private const uint SIGDN_FILESYSPATH  = 0x80058000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            [In] ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        /// <summary>Shows the Explorer-style folder picker. Returns null when cancelled.</summary>
        public static string Show(IntPtr owner, string title, string okLabel, string initialPath)
        {
            IFileDialog dialog = (IFileDialog)(new FileOpenDialogCoClass());
            try
            {
                uint options;
                dialog.GetOptions(out options);
                dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM
                                          | FOS_PATHMUSTEXIST | FOS_NOCHANGEDIR);

                if (!string.IsNullOrEmpty(title))   { dialog.SetTitle(title); }
                if (!string.IsNullOrEmpty(okLabel)) { dialog.SetOkButtonLabel(okLabel); }

                // A folder that no longer exists (or an offline share) must not stop
                // the dialog from opening, so this is best effort only.
                if (!string.IsNullOrEmpty(initialPath))
                {
                    try
                    {
                        Guid shellItemId = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
                        IShellItem start;
                        SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref shellItemId, out start);
                        if (start != null) { dialog.SetFolder(start); }
                    }
                    catch { }
                }

                // Any non-zero HRESULT means "no selection"; cancelling returns
                // HRESULT_FROM_WIN32(ERROR_CANCELLED).
                if (dialog.Show(owner) != 0) { return null; }

                IShellItem result;
                dialog.GetResult(out result);
                IntPtr buffer = IntPtr.Zero;
                try
                {
                    result.GetDisplayName(SIGDN_FILESYSPATH, out buffer);
                    return Marshal.PtrToStringUni(buffer);
                }
                finally
                {
                    if (buffer != IntPtr.Zero) { Marshal.FreeCoTaskMem(buffer); }
                    Marshal.ReleaseComObject(result);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }
    }
}
'@

    try {
        Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
        $script:FolderPickerState = 'Modern'
    }
    catch {
        $script:FolderPickerState = 'Classic'
        Write-AppLog -Level 'WARN' -Message ('Explorer-style folder picker unavailable, using the classic dialog: {0}' -f $_.Exception.Message)
    }

    return ($script:FolderPickerState -eq 'Modern')
}

function Show-FolderPicker {
    <#
        .SYNOPSIS
            Asks the user for a folder and returns its path, or $null if cancelled.
        .DESCRIPTION
            Uses the same dialog Excel does when it is available. The classic
            fallback is configured as closely as possible: rooted at the desktop
            so network locations stay reachable, and with the new-folder button.
    #>
    [CmdletBinding()]
    param(
        [string]$InitialPath = '',
        [string]$Title       = 'Select a folder',
        [string]$OkLabel     = 'Select Folder',
        [System.Windows.Forms.IWin32Window]$Owner = $null
    )

    if (Initialize-FolderPicker) {
        $handle = [IntPtr]::Zero
        if ($null -ne $Owner) { try { $handle = $Owner.Handle } catch { $handle = [IntPtr]::Zero } }
        try {
            return [ExcelQueryTrigger.FolderPicker]::Show($handle, $Title, $OkLabel, $InitialPath)
        }
        catch {
            # Fall through to the classic dialog rather than losing the click.
            Write-AppLog -Level 'WARN' -Message ('Folder picker failed, falling back: {0}' -f $_.Exception.Message)
            $script:FolderPickerState = 'Classic'
        }
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $dialog.Description         = $Title
        $dialog.ShowNewFolderButton = $true
        $dialog.RootFolder          = [System.Environment+SpecialFolder]::Desktop
        if (-not [string]::IsNullOrWhiteSpace($InitialPath)) {
            try { $dialog.SelectedPath = $InitialPath } catch { }
        }
        $result = if ($null -ne $Owner) { $dialog.ShowDialog($Owner) } else { $dialog.ShowDialog() }
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
        return $null
    }
    finally {
        $dialog.Dispose()
    }
}

function Show-FilePicker {
    <#
        .SYNOPSIS
            Wrapper around OpenFileDialog so both pickers behave the same way.
        .DESCRIPTION
            OpenFileDialog is already the Explorer-style dialog. Centralising it
            here keeps the "start where the current value points" behaviour - and
            the tolerance for a value that points somewhere unreachable - in one
            place.
    #>
    [CmdletBinding()]
    param(
        [string]$InitialPath = '',
        [string]$Title       = 'Select a file',
        [string]$Filter      = 'All files (*.*)|*.*',
        [System.Windows.Forms.IWin32Window]$Owner = $null
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialog.Title            = $Title
        $dialog.Filter           = $Filter
        $dialog.CheckFileExists  = $false
        $dialog.RestoreDirectory = $true
        if (-not [string]::IsNullOrWhiteSpace($InitialPath)) {
            try {
                $parent = Split-Path -Parent $InitialPath
                if (-not [string]::IsNullOrWhiteSpace($parent)) { $dialog.InitialDirectory = $parent }
                $leaf = Split-Path -Leaf $InitialPath
                if (-not [string]::IsNullOrWhiteSpace($leaf)) { $dialog.FileName = $leaf }
            }
            catch { }
        }
        $result = if ($null -ne $Owner) { $dialog.ShowDialog($Owner) } else { $dialog.ShowDialog() }
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
        return $null
    }
    finally {
        $dialog.Dispose()
    }
}

function Remove-UiComReference {
    param([Parameter(ValueFromRemainingArguments = $true)]$ComObjects)
    foreach ($obj in $ComObjects) {
        if ($null -eq $obj) { continue }
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
            }
        }
        catch { }
    }
}

# Only Excel's own plumbing is matched by name. The Power Query helper queries
# ("Sample File", "Transform Sample File", "Parameter1") are NOT listed here on
# purpose: those names are localised - サンプル ファイル, パラメーター1 - so a
# name list would only ever work on an English build. They are hidden because
# they load nowhere, which is true in every language.
$script:UiInternalConnectionPatterns = @(
    '^ThisWorkbookDataModel$',
    '^WorksheetConnection_',
    '^LinkedTable_',
    '^NativeTimeline_',
    '^NativeSlicer_',
    '^ModelConnection'
)

function Test-ConnectionIsExcelInternal {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    foreach ($pattern in $script:UiInternalConnectionPatterns) {
        if ($Name -match $pattern) { return $true }
    }
    return $false
}

function Add-LoadTarget {
    param([hashtable]$Map, [string]$Name, [string]$LoadsTo)
    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    $key = $Name.ToLowerInvariant()
    if (-not $Map.ContainsKey($key)) { $Map[$key] = $LoadsTo }
}

function Get-WorkbookLoadTargets {
    <#
        Where each query or connection actually puts its data. Keyed by both
        connection name and query name, because the two are spelled differently
        and either may be what we hold.

        Four independent signals, because no single one covers every workbook:
          1. a table on a sheet fed by a query table
          2. a PivotTable's cache
          3. Excel's own WorksheetConnection_<file>!<table> connection, which
             only exists when a query was loaded onto a sheet
          4. the Data Model's tables

        An empty result means "could not tell", never "loads nowhere" - the
        caller must not hide anything on the strength of a failed lookup.
    #>
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        $Connections = @()
    )

    $targets = @{}
    $determined = $false

    # 1. tables on sheets
    $sheets = $null
    try { $sheets = $Workbook.Worksheets } catch { $sheets = $null }
    if ($null -ne $sheets) {
        try {
            for ($s = 1; $s -le $sheets.Count; $s++) {
                $sheet = $null
                try {
                    $sheet = $sheets.Item($s)
                    $listObjects = $null
                    try { $listObjects = $sheet.ListObjects } catch { $listObjects = $null }
                    if ($null -eq $listObjects) { continue }
                    try {
                        for ($k = 1; $k -le $listObjects.Count; $k++) {
                            $listObject = $null
                            $queryTable = $null
                            try {
                                $listObject = $listObjects.Item($k)
                                # The table's own name is what a query loaded to
                                # a sheet is called, and it is always readable.
                                try { Add-LoadTarget -Map $targets -Name ([string]$listObject.Name) -LoadsTo 'Table'; $determined = $true } catch { }
                                try { $queryTable = $listObject.QueryTable } catch { $queryTable = $null }
                                if ($null -ne $queryTable) {
                                    try { Add-LoadTarget -Map $targets -Name ([string]$queryTable.WorkbookConnection.Name) -LoadsTo 'Table'; $determined = $true } catch { }
                                    try {
                                        $location = Get-MashupQueryLocation -ConnectionString ([string]$queryTable.Connection)
                                        if (-not [string]::IsNullOrWhiteSpace($location)) { Add-LoadTarget -Map $targets -Name $location -LoadsTo 'Table'; $determined = $true }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                            finally { Remove-UiComReference $queryTable $listObject }
                        }
                    }
                    finally { Remove-UiComReference $listObjects }
                }
                catch { }
                finally { Remove-UiComReference $sheet }
            }
        }
        catch { }
        finally { Remove-UiComReference $sheets }
    }

    # 2. PivotTable caches, taken at workbook level so no sheet is missed
    $caches = $null
    try { $caches = $Workbook.PivotCaches() } catch { $caches = $null }
    if ($null -ne $caches) {
        try {
            for ($c = 1; $c -le $caches.Count; $c++) {
                $cache = $null
                try {
                    $cache = $caches.Item($c)
                    $name = ''
                    try { $name = [string]$cache.WorkbookConnection } catch { $name = '' }
                    if ([string]::IsNullOrWhiteSpace($name)) { try { $name = [string]$cache.WorkbookConnection.Name } catch { $name = '' } }
                    if (-not [string]::IsNullOrWhiteSpace($name)) { Add-LoadTarget -Map $targets -Name $name -LoadsTo 'PivotTable'; $determined = $true }
                }
                catch { }
                finally { Remove-UiComReference $cache }
            }
        }
        catch { }
        finally { Remove-UiComReference $caches }
    }

    # 3. Excel's WorksheetConnection_<file>!<table> - proof of a sheet load
    foreach ($connection in @($Connections)) {
        $sheetTable = [string]$connection.SheetTable
        if (-not [string]::IsNullOrWhiteSpace($sheetTable)) {
            Add-LoadTarget -Map $targets -Name $sheetTable -LoadsTo 'Table'
            $determined = $true
        }
    }

    # 4. the Data Model
    $model = $null
    try { $model = $Workbook.Model } catch { $model = $null }
    if ($null -ne $model) {
        $modelTables = $null
        try { $modelTables = $model.ModelTables } catch { $modelTables = $null }
        if ($null -ne $modelTables) {
            try {
                for ($m = 1; $m -le $modelTables.Count; $m++) {
                    $modelTable = $null
                    try {
                        $modelTable = $modelTables.Item($m)
                        Add-LoadTarget -Map $targets -Name ([string]$modelTable.Name) -LoadsTo 'Data Model'
                        try { Add-LoadTarget -Map $targets -Name ([string]$modelTable.SourceName) -LoadsTo 'Data Model' } catch { }
                        $determined = $true
                    }
                    catch { }
                    finally { Remove-UiComReference $modelTable }
                }
            }
            catch { }
            finally { Remove-UiComReference $modelTables }
        }
        Remove-UiComReference $model
    }

    return @{ Targets = $targets; Determined = $determined }
}

function Get-ExcelUnavailableMessage {
    <#  A COM failure said in words rather than in hexadecimal.  #>
    param($Exception)

    $text = [string]$Exception.Message
    if ($Exception -is [System.Runtime.InteropServices.COMException] -or $text -match '80040154|REGDB_E_CLASSNOTREG|800401F3') {
        return 'Excel could not be started on this computer, so its query list cannot be read. Excel may not be installed here, or its registration may be broken. The rule will still work on a machine that has Excel.'
    }
    if ($text -match '(?i)access is denied|0x80070005') {
        return 'Excel refused to start for this account. This is usually a security policy on the machine.'
    }
    return ('Excel could not be asked for the query list. {0}' -f $text)
}

function Get-QueryCatalogFromFile {
    <#
        The same catalogue, read out of the workbook file instead of out of
        Excel. It cannot tell where each query loads - that needs the object
        model - so everything comes back marked Unknown and nothing is hidden.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $result = @{ Success = $false; Queries = @(); Items = @(); Unrefreshable = @(); Determined = $false; Error = '' }

    $definitions = Get-WorkbookConnectionDefinitions -Path $Path
    if (-not $definitions.Ok) {
        $result.Error = [string]$definitions.Reason
        return $result
    }

    $items     = New-Object System.Collections.ArrayList
    $available = New-Object System.Collections.ArrayList

    foreach ($connection in @($definitions.Connections)) {
        $connectionName = [string]$connection.Name
        if (Test-ConnectionIsExcelInternal -Name $connectionName) {
            [void]$available.Add($connectionName)
            [void]$items.Add(@{
                Name = $connectionName; ConnectionName = $connectionName
                Kind = 'Excel internal'; LoadsTo = 'Internal'; Primary = $false
            })
            continue
        }

        # A Power Query connection names its query in Location=, which is the
        # name a person recognises.
        $location = Get-MashupQueryLocation -ConnectionString ([string]$connection.ConnectionString)
        $displayName = $(if ([string]::IsNullOrWhiteSpace($location)) { $connectionName } else { $location })
        $kind = $(if ([string]::IsNullOrWhiteSpace($location)) { 'Connection' } else { 'Power Query' })

        [void]$available.Add($displayName)
        [void]$items.Add(@{
            Name = $displayName; ConnectionName = $connectionName
            Kind = $kind; LoadsTo = 'Unknown'; Primary = $true
        })
    }

    $result.Items   = @($items.ToArray())
    $result.Queries = @($available.ToArray())
    $result.Success = $true
    return $result
}

function Get-WorkbookQueryCatalogForUi {
    <#
        Opens a workbook read-only in a dedicated hidden Excel instance and
        returns everything in it that can be refreshed on its own, with where
        each item loads its data. Used only while configuring an action; it
        never saves the workbook.

        Returns @{ Success; Items; Queries; Unrefreshable; Determined; Error }
        Determined is false when nothing could be traced to a sheet, a
        PivotTable or the Data Model - in that case "loads nowhere" is unknown
        rather than false, and the caller must not hide anything.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [bool]$AllowWorkbookMacros = $false
    )

    $result = @{ Success = $false; Queries = @(); Items = @(); Unrefreshable = @(); Determined = $false; Error = '' }
    if (-not (Test-Path -LiteralPath $Path)) {
        $result.Error = 'Workbook not found or the path is currently unavailable.'
        return $result
    }

    $excel = $null; $workbooks = $null; $workbook = $null; $workbookOpen = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        if ($AllowWorkbookMacros) {
            try { $excel.AutomationSecurity = 2; $excel.EnableEvents = $true } catch { }
        }
        else {
            try { $excel.AutomationSecurity = 3; $excel.EnableEvents = $false }
            catch {
                $result.Error = 'Excel macro security could not be forced to Disabled, so the workbook was not opened for query discovery.'
                return $result
            }
        }

        $workbooks = $excel.Workbooks
        $missing = [System.Reflection.Missing]::Value
        $workbook = $workbooks.Open($Path, 0, $true, $missing, $missing, $missing, $true,
            $missing, $missing, $missing, $false)
        $workbookOpen = $true

        $release = { param($a, $b) Remove-UiComReference $a $b }
        $connections = @(Get-WorkbookConnectionSummary -Workbook $workbook -Release $release)
        $load = Get-WorkbookLoadTargets -Workbook $workbook -Connections $connections
        $targets = $load.Targets
        $result.Determined = [bool]$load.Determined

        # Location= is the locale-independent link between a query and the
        # connection that refreshes it.
        $connectionForQuery = @{}
        foreach ($connection in $connections) {
            $location = [string]$connection.QueryLocation
            if (-not [string]::IsNullOrWhiteSpace($location)) {
                $connectionForQuery[$location.ToLowerInvariant()] = [string]$connection.Name
            }
        }

        $lookupLoad = {
            param([string[]]$Names)
            foreach ($candidate in $Names) {
                if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
                $key = $candidate.ToLowerInvariant()
                if ($targets.ContainsKey($key)) { return [string]$targets[$key] }
            }
            if ($load.Determined) { return 'Connection only' }
            return 'Unknown'
        }

        $items         = New-Object System.Collections.ArrayList
        $available     = New-Object System.Collections.ArrayList
        $unrefreshable = New-Object System.Collections.ArrayList
        $claimed       = @{}

        $queries = $null
        try { $queries = $workbook.Queries } catch { $queries = $null }
        if ($null -ne $queries) {
            try {
                for ($i = 1; $i -le $queries.Count; $i++) {
                    $query = $null
                    try {
                        $query = $queries.Item($i)
                        $queryName = [string]$query.Name
                        if ([string]::IsNullOrWhiteSpace($queryName)) { continue }

                        $key = $queryName.ToLowerInvariant()
                        $connectionName = ''
                        if ($connectionForQuery.ContainsKey($key)) { $connectionName = [string]$connectionForQuery[$key] }
                        if ([string]::IsNullOrWhiteSpace($connectionName)) { [void]$unrefreshable.Add($queryName); continue }

                        $claimed[$connectionName.ToLowerInvariant()] = $true
                        [void]$available.Add($queryName)
                        $loadsTo = & $lookupLoad @($queryName, $connectionName)
                        [void]$items.Add(@{
                            Name           = $queryName
                            ConnectionName = $connectionName
                            Kind           = 'Power Query'
                            LoadsTo        = $loadsTo
                            Primary        = ($loadsTo -ne 'Connection only')
                        })
                    }
                    catch { }
                    finally { Remove-UiComReference $query }
                }
            }
            finally { Remove-UiComReference $queries }
        }

        foreach ($connection in $connections) {
            $connectionName = [string]$connection.Name
            if ($claimed.ContainsKey($connectionName.ToLowerInvariant())) { continue }

            if (Test-ConnectionIsExcelInternal -Name $connectionName) {
                [void]$available.Add($connectionName)
                [void]$items.Add(@{
                    Name = $connectionName; ConnectionName = $connectionName
                    Kind = 'Excel internal'; LoadsTo = 'Internal'; Primary = $false
                })
                continue
            }

            [void]$available.Add($connectionName)
            $loadsTo = & $lookupLoad @($connectionName)
            [void]$items.Add(@{
                Name           = $connectionName
                ConnectionName = $connectionName
                Kind           = 'Connection'
                LoadsTo        = $loadsTo
                Primary        = ($loadsTo -ne 'Connection only')
            })
        }

        $result.Items = @($items.ToArray())
        $result.Queries = @($available.ToArray())
        $result.Unrefreshable = @($unrefreshable.ToArray())
        $result.Success = $true
    }
    catch {
        # Excel could not be driven - not installed on this machine, blocked by
        # policy, or its COM registration is broken. The list can still be read
        # out of the workbook file itself, so fall back to that rather than
        # showing the person a CLSID.
        $fallback = Get-QueryCatalogFromFile -Path $Path
        if ($fallback.Success) { return $fallback }
        $result.Error = Get-ExcelUnavailableMessage -Exception $_.Exception
    }
    finally {
        if ($workbookOpen -and $null -ne $workbook) { try { $workbook.Close($false) } catch { } }
        Remove-UiComReference $workbook $workbooks
        if ($null -ne $excel) { try { $excel.Quit() } catch { }; Remove-UiComReference $excel }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return $result
}

function Show-QuerySelectionDialog {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [string[]]$SelectedQueries = @(),
        [bool]$AllowWorkbookMacros = $false,
        [System.Windows.Forms.IWin32Window]$Owner
    )

    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try { $catalog = Get-WorkbookQueryCatalogForUi -Path $WorkbookPath -AllowWorkbookMacros:$AllowWorkbookMacros }
    finally { [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }

    if (-not $catalog.Success) {
        [System.Windows.Forms.MessageBox]::Show(
            ('The query list could not be loaded:' + [Environment]::NewLine + [Environment]::NewLine + $catalog.Error),
            'Power Query list', 'OK', 'Warning') | Out-Null
        return $null
    }
    if (@($catalog.Queries).Count -eq 0) {
        $extra = ''
        if (-not $AllowWorkbookMacros -and ([System.IO.Path]::GetExtension($WorkbookPath).ToLowerInvariant() -in @('.xlsm','.xlsb'))) {
            $extra = [Environment]::NewLine + [Environment]::NewLine + 'If Workbook_Open creates the connections, retry after allowing macros for this workbook.'
        }
        [System.Windows.Forms.MessageBox]::Show(
            ('Nothing in this workbook can be refreshed on its own.' + $extra),
            'Query list', 'OK', 'Information') | Out-Null
        return $null
    }

    # Everything the workbook can refresh, in a stable order: the things that
    # actually put data on a sheet first, then the staging and helper items.
    $allItems = @($catalog.Items)
    if ($allItems.Count -eq 0) {
        $allItems = @(@($catalog.Queries) | ForEach-Object {
            @{ Name = [string]$_; ConnectionName = [string]$_; Kind = 'Connection'; LoadsTo = 'Unknown'; Primary = $true } })
    }
    $primaryCount = @($allItems | Where-Object { [bool]$_.Primary }).Count
    # If nothing could be identified as loaded, hiding anything would only get
    # in the way, so everything is shown and the toggle starts on.
    $startExpanded = ($primaryCount -eq 0)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select queries and connections'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize = New-Object System.Drawing.Size(600,470)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterParent'
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    try { $form.Font = Get-UiFont } catch { }

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(14,12)
    $lbl.Size = New-Object System.Drawing.Size(570,40)
    $lbl.Text = $(if ($catalog.Determined) {
        'Choose what to refresh. Items that load nowhere are hidden - refreshing a query still evaluates the upstream queries it depends on.'
    } else {
        'Choose what to refresh. This list was read from the workbook file, so where each item loads is not known and everything is shown.'
    })
    $form.Controls.Add($lbl)

    $list = New-Object System.Windows.Forms.ListView
    $list.Location = New-Object System.Drawing.Point(14,56)
    $list.Size = New-Object System.Drawing.Size(570,280)
    $list.View = 'Details'
    $list.CheckBoxes = $true
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.HideSelection = $false
    [void]$list.Columns.Add('Name', 300)
    [void]$list.Columns.Add('Type', 120)
    [void]$list.Columns.Add('Loads to', 130)
    $form.Controls.Add($list)

    $chkShowAll = New-Object System.Windows.Forms.CheckBox
    $chkShowAll.Location = New-Object System.Drawing.Point(14,344)
    $chkShowAll.Size = New-Object System.Drawing.Size(570,22)
    $chkShowAll.Text = ('Also show staging, helper and internal items ({0} hidden)' -f ($allItems.Count - $primaryCount))
    $chkShowAll.Checked = $startExpanded
    $chkShowAll.Enabled = ($allItems.Count -ne $primaryCount)
    $form.Controls.Add($chkShowAll)

    # The checked set is kept here rather than read off the rows, because rows
    # come and go as the toggle changes and a hidden row must not be forgotten.
    $checkedNames = New-Object System.Collections.ArrayList
    foreach ($name in @($SelectedQueries)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$name)) { [void]$checkedNames.Add([string]$name) }
    }

    # ItemChecked and rebuild are separate event closures. A shared reference
    # keeps the temporary suppression flag visible to both of them.
    $selectionState = @{ Suspend = $false }
    $rebuild = {
        $selectionState.Suspend = $true
        $list.BeginUpdate()
        try {
            $list.Items.Clear()
            foreach ($item in $allItems) {
                $name = [string]$item.Name
                $isChecked = (@($checkedNames.ToArray()) -contains $name)
                # An item that is already selected always stays visible, or it
                # could never be unselected again.
                if (-not $chkShowAll.Checked -and -not [bool]$item.Primary -and -not $isChecked) { continue }
                $row = New-Object System.Windows.Forms.ListViewItem($name)
                [void]$row.SubItems.Add([string]$item.Kind)
                [void]$row.SubItems.Add([string]$item.LoadsTo)
                if (-not [bool]$item.Primary) { $row.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120) }
                $row.Checked = $isChecked
                [void]$list.Items.Add($row)
            }
        }
        finally {
            $list.EndUpdate()
            $selectionState.Suspend = $false
        }
    }.GetNewClosure()

    $list.Add_ItemChecked({
        param($sender, $eventArgs)
        if ([bool]$selectionState.Suspend) { return }
        $name = [string]$eventArgs.Item.Text
        $index = @($checkedNames.ToArray()).IndexOf($name)
        if ($eventArgs.Item.Checked) { if ($index -lt 0) { [void]$checkedNames.Add($name) } }
        elseif ($index -ge 0) { $checkedNames.RemoveAt($index) }
    }.GetNewClosure())

    $chkShowAll.Add_CheckedChanged($rebuild)
    & $rebuild

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text='Select All'; $btnAll.Location=New-Object System.Drawing.Point(14,376); $btnAll.Size=New-Object System.Drawing.Size(96,28)
    $form.Controls.Add($btnAll)
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text='Clear'; $btnClear.Location=New-Object System.Drawing.Point(118,376); $btnClear.Size=New-Object System.Drawing.Size(84,28)
    $form.Controls.Add($btnClear)
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text='Copy list'; $btnCopy.Location=New-Object System.Drawing.Point(210,376); $btnCopy.Size=New-Object System.Drawing.Size(96,28)
    $form.Controls.Add($btnCopy)
    $btnCopy.Add_Click({
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add(('Workbook: {0}' -f $WorkbookPath))
        [void]$lines.Add(('Load targets identified: {0}' -f $catalog.Determined))
        [void]$lines.Add('')
        [void]$lines.Add(('{0}`t{1}`t{2}`t{3}' -f 'Name', 'Type', 'Loads to', 'Connection'))
        foreach ($item in $allItems) {
            [void]$lines.Add(('{0}`t{1}`t{2}`t{3}' -f $item.Name, $item.Kind, $item.LoadsTo, $item.ConnectionName))
        }
        if (@($catalog.Unrefreshable).Count -gt 0) {
            [void]$lines.Add('')
            [void]$lines.Add(('No workbook connection: {0}' -f (@($catalog.Unrefreshable) -join ', ')))
        }
        try { [System.Windows.Forms.Clipboard]::SetText(($lines.ToArray() -join [Environment]::NewLine)) } catch { }
    }.GetNewClosure())
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text='OK'; $btnOk.Location=New-Object System.Drawing.Point(400,424); $btnOk.Size=New-Object System.Drawing.Size(88,30)
    $form.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text='Cancel'; $btnCancel.Location=New-Object System.Drawing.Point(496,424); $btnCancel.Size=New-Object System.Drawing.Size(88,30); $btnCancel.DialogResult='Cancel'
    $form.Controls.Add($btnCancel); $form.CancelButton=$btnCancel

    # Select All and Clear act on what is on screen, not on what is hidden.
    $btnAll.Add_Click({ foreach($row in $list.Items){ $row.Checked=$true } }.GetNewClosure())
    $btnClear.Add_Click({ foreach($row in $list.Items){ $row.Checked=$false } }.GetNewClosure())
    $btnOk.Add_Click({
        if (@($checkedNames.ToArray()).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Select at least one item.','Query list','OK','Warning') | Out-Null; return
        }
        $form.Tag=@($checkedNames.ToArray()); $form.DialogResult='OK'; $form.Close()
    }.GetNewClosure())

    try { Set-FormWithinWorkingArea -Form $form } catch { }
    if ($null -ne $Owner) { $dialogResult=$form.ShowDialog($Owner) } else { $dialogResult=$form.ShowDialog() }
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) { return @($form.Tag) }
    return $null
}

function Show-AdhocRefreshConfirmation {
    <# Final safety confirmation for Refresh Any Excel File. #>
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [bool]$AllowWorkbookMacros = $false,
        [int]$WarningSeconds = 300,
        [System.Windows.Forms.IWin32Window]$Owner
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Confirm Excel Refresh'
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.ClientSize = New-Object System.Drawing.Size(600,350)
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterParent'
    $form.MaximizeBox=$false; $form.MinimizeBox=$false
    try { $form.Font = Get-UiFont } catch { }

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text='Review before refresh'; try { $lblTitle.Font = Get-UiFont 9 'Bold' } catch { $lblTitle.Font = $form.Font }
    $lblTitle.Location=New-Object System.Drawing.Point(16,14); $lblTitle.AutoSize=$true
    $form.Controls.Add($lblTitle)

    $labels = @(
        @{T='Workbook:';Y=50}, @{T='Location:';Y=78}, @{T='Macros:';Y=122}, @{T='Refresh:';Y=152}, @{T='Warn after:';Y=210}
    )
    foreach($spec in $labels){ $l=New-Object System.Windows.Forms.Label; $l.Text=$spec.T; $l.Location=New-Object System.Drawing.Point(18,$spec.Y); $l.Size=New-Object System.Drawing.Size(100,22); $form.Controls.Add($l) }

    $lblName=New-Object System.Windows.Forms.Label; $lblName.Text=(Split-Path -Leaf $WorkbookPath); $lblName.Location=New-Object System.Drawing.Point(120,50); $lblName.Size=New-Object System.Drawing.Size(460,22); $form.Controls.Add($lblName)
    $txtPath=New-Object System.Windows.Forms.TextBox; $txtPath.Text=$WorkbookPath; $txtPath.Location=New-Object System.Drawing.Point(120,76); $txtPath.Size=New-Object System.Drawing.Size(460,24); $txtPath.ReadOnly=$true; $form.Controls.Add($txtPath)
    $lblMacros=New-Object System.Windows.Forms.Label; $lblMacros.Text=$(if($AllowWorkbookMacros){'Allowed - Excel Trust Center policy still applies'}else{'Blocked (recommended)'}); $lblMacros.Location=New-Object System.Drawing.Point(120,122); $lblMacros.Size=New-Object System.Drawing.Size(460,22); $form.Controls.Add($lblMacros)

    $cbo=New-Object System.Windows.Forms.ComboBox; $cbo.Location=New-Object System.Drawing.Point(120,150); $cbo.Size=New-Object System.Drawing.Size(190,24); $cbo.DropDownStyle='DropDownList'; [void]$cbo.Items.Add('All queries'); [void]$cbo.Items.Add('Selected queries'); $cbo.SelectedIndex=0; $form.Controls.Add($cbo)
    $btnQueries=New-Object System.Windows.Forms.Button; $btnQueries.Text='Choose queries...'; $btnQueries.Location=New-Object System.Drawing.Point(320,147); $btnQueries.Size=New-Object System.Drawing.Size(130,28); $btnQueries.Enabled=$false; $form.Controls.Add($btnQueries)
    $lblQueries=New-Object System.Windows.Forms.Label; $lblQueries.Location=New-Object System.Drawing.Point(120,180); $lblQueries.Size=New-Object System.Drawing.Size(460,22); $lblQueries.ForeColor=[System.Drawing.Color]::FromArgb(95,95,95); $lblQueries.Text='All workbook queries/connections will be refreshed.'; $form.Controls.Add($lblQueries)
    $lblWarn=New-Object System.Windows.Forms.Label; $lblWarn.Text=('{0} seconds (warning only - refresh continues)' -f $WarningSeconds); $lblWarn.Location=New-Object System.Drawing.Point(120,210); $lblWarn.Size=New-Object System.Drawing.Size(460,22); $form.Controls.Add($lblWarn)

    $note=New-FormWrappedLabel -Text 'Excel opens in a dedicated hidden instance. The workbook is saved only after the refresh has completed. You can cancel here without opening Excel.' -X 18 -Y 242 -Width 562; $form.Controls.Add($note)

    $btnRefresh=New-Object System.Windows.Forms.Button; $btnRefresh.Text='Refresh'; $btnRefresh.Location=New-Object System.Drawing.Point(294,304); $btnRefresh.Size=New-Object System.Drawing.Size(88,30); $form.Controls.Add($btnRefresh)
    $btnAnother=New-Object System.Windows.Forms.Button; $btnAnother.Text='Choose Another'; $btnAnother.Location=New-Object System.Drawing.Point(388,304); $btnAnother.Size=New-Object System.Drawing.Size(104,30); $form.Controls.Add($btnAnother)
    $btnCancel=New-Object System.Windows.Forms.Button; $btnCancel.Text='Cancel'; $btnCancel.Location=New-Object System.Drawing.Point(498,304); $btnCancel.Size=New-Object System.Drawing.Size(82,30); $form.Controls.Add($btnCancel); $form.CancelButton=$btnCancel

    # Keep selected queries in a shared reference because each WinForms event
    # handler has its own PowerShell scope.
    $queryState=@{ Selected=@() }
    $update={
        $isSelected=($cbo.SelectedIndex -eq 1); $btnQueries.Enabled=$isSelected
        if(-not $isSelected){ $lblQueries.Text='All workbook queries/connections will be refreshed.' }
        elseif(@($queryState.Selected).Count -eq 0){ $lblQueries.Text='No queries selected yet.' }
        elseif(@($queryState.Selected).Count -le 4){ $lblQueries.Text=(@($queryState.Selected)-join ', ') }
        else { $lblQueries.Text=('{0} queries selected' -f @($queryState.Selected).Count) }
    }.GetNewClosure()
    $cbo.Add_SelectedIndexChanged($update)
    $btnQueries.Add_Click({
        $chosen=Show-QuerySelectionDialog -WorkbookPath $WorkbookPath -SelectedQueries @($queryState.Selected) -AllowWorkbookMacros:$AllowWorkbookMacros -Owner $form
        if($null -ne $chosen){$queryState.Selected=@($chosen); & $update}
    }.GetNewClosure())
    $btnAnother.Add_Click({$form.Tag=@{Decision='ChooseAnother'};$form.DialogResult=[System.Windows.Forms.DialogResult]::Retry;$form.Close()}.GetNewClosure())
    $btnCancel.Add_Click({$form.Tag=@{Decision='Cancel'};$form.DialogResult=[System.Windows.Forms.DialogResult]::Cancel;$form.Close()}.GetNewClosure())
    $btnRefresh.Add_Click({
        if($cbo.SelectedIndex -eq 1 -and @($queryState.Selected).Count -eq 0){[System.Windows.Forms.MessageBox]::Show('Select at least one query, or choose All queries.','Confirm Excel Refresh','OK','Warning')|Out-Null;return}
        $action=Get-DefaultAction
        $action['path']=$WorkbookPath
        $action['timeoutSeconds']=$WarningSeconds
        $action['allowWorkbookMacros']=$AllowWorkbookMacros
        $action['refreshMethod']=$(if($cbo.SelectedIndex -eq 1){'SelectedQueries'}else{'RefreshAll'})
        $action['selectedQueries']=@($queryState.Selected)
        $form.Tag=@{Decision='Refresh';Action=$action};$form.DialogResult=[System.Windows.Forms.DialogResult]::OK;$form.Close()
    }.GetNewClosure())

    try { Set-FormWithinWorkingArea -Form $form } catch { }
    if($null -ne $Owner){[void]$form.ShowDialog($Owner)}else{[void]$form.ShowDialog()}
    if($null -eq $form.Tag){return @{Decision='Cancel'}}
    return $form.Tag
}
