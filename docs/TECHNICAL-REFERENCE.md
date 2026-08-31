# Technical Reference

## Runtime model

Excel Query Trigger Manager is a PowerShell/WinForms application with separate responsibilities for configuration, trigger monitoring, job management, Excel automation, logging, and UI.

Refresh jobs are serialized through a queue. The same application does not intentionally refresh the same workbook through two app-owned Excel instances at the same time.

## File monitoring

File and folder triggers use Windows file-system events rather than repeatedly enumerating every file in a folder.

Conceptually:

`Windows file event → type/name filter → debounce → file-ready check → cooldown → job/confirmation`

### Debounce

A single physical file operation can generate several Windows events. Debounce
combines closely spaced matching changes into one rule run while retaining every
distinct source-file path. File readiness is checked for every retained path.

### Cooldown / minimum repeat interval

After a rule fires, the minimum repeat interval prevents the same rule from firing again immediately.

### File-ready check

For rules configured to wait for a complete source file, the application checks whether the file appears ready before starting the refresh workflow.

### Monitoring health check

The application periodically verifies that monitoring infrastructure is still active. This is not a per-second folder scan.

### No historical scan

The application does not maintain a snapshot of folder contents for later comparison. Starting monitoring establishes the point from which new Windows events are observed.

## Job queue

Triggered and manual refresh work is queued and processed sequentially.

Queue coalescing uses the normalized workbook path plus query scope, selected
queries, macro/background policy, failure policy, and warning threshold. Two
different refresh operations on the same workbook are not silently combined.

Ask-before-refresh happens before the refresh is queued/opened in Excel. A skipped confirmation does not create an Excel refresh job.

The optional recent-refresh guard also runs before queueing. Workbook metadata is
read on a worker runspace so an unavailable network share cannot block the UI.
`pivotCacheDefinition*.xml` refresh timestamps and the application's portable
`ExcelQueryTriggerLastQueryRefreshUtc` custom property are used. The workbook
save timestamp is deliberately excluded because saving and refreshing are
different events. A recent query-refresh timestamp or an unreadable/unknown
timestamp requires confirmation, while an older known timestamp is approved
automatically. This is advisory duplicate prevention, not a distributed lock
between computers.

## Workbook metadata reading

Workbook metadata for the Dashboard is read from the file itself rather than
through Excel. The reader copies the workbook with a share mode of
`ReadWrite | Delete` so the file can still be renamed or replaced while the copy
is in progress, which is what Excel does at the end of a save.

The background scan runs in a fresh runspace each time, so the known file stamps
(`LastWriteTimeUtc` plus length) are passed into the worker; a workbook whose
stamp is unchanged is not opened. No scan is started while a refresh job is
active, a scan already in progress is stopped when a job starts, and the
workbook the running job has open is excluded from the scan list.

## Excel instance isolation

Refreshes use a dedicated Excel application instance controlled by the tool.

This isolation is why users can normally continue working in other Excel workbooks while a background refresh is running.

The same target workbook should not be opened interactively during the automated refresh.

## Refresh completion detection

The application does not rely on a fixed sleep period.

It observes supported Excel refresh/calculation states as `Busy`, `Quiet`, or
`Unknown` and waits for consistently quiet samples. Repeated state-read failures
or a failed `CalculateUntilAsyncQueriesDone` confirmation stop the job without
saving.

The configured long-refresh value is a warning threshold, not a forced timeout.

## Foreground/synchronous query mode

For supported OLE DB and ODBC connections, the application can temporarily request synchronous refresh behavior (`BackgroundQuery = False`) so completion is easier to detect reliably.

Original supported `BackgroundQuery` values are restored before the workbook is
saved. A restoration failure blocks the save.

This setting concerns query execution; it does not mean that the Excel window must be visible or brought to the foreground.

## Selected-query refresh

Power Query names are discovered from the workbook and mapped to refreshable Excel connections where possible.

Selecting a query controls the refresh entry point. Power Query may still evaluate dependent upstream queries.

## Save and close

Normal successful completion always saves and closes the hidden workbook. These
are safety invariants; legacy or hand-edited configuration cannot override them.

The save is a direct `Workbook.Save`. Before it is entered, the file's timestamp
and size are compared against the values recorded before the refresh started, so
a workbook that a different user saved during a long refresh is not overwritten.

Cancellation and forced process termination are both withheld during the save.
Dialog inspection is not: a modal dialog raised by Excel at that point is closed,
which turns an unanswerable prompt into a reported error.

Cancellation is intentionally different: a cancelled refresh must not fall through the normal successful-save path.

## Macro security

Unattended refreshes block macros by default.

For macro-enabled workbooks, users can explicitly allow macros when required for workbook initialization. Allowing macros does not override Excel Trust Center or enterprise security policy.

If the application cannot establish the requested macro-blocking security state, it should fail safely rather than silently open the workbook with uncertain macro behavior.

## Configuration and state

Rules and application settings are stored as JSON. Runtime state such as last-run information is stored separately.

Configuration/state writes use temporary-file replacement behavior to reduce the risk of leaving a partially written JSON file.

The configuration schema is version 4. When an older configuration is loaded,
retired feedback destinations, credentials and the former pixel-based rules-pane
height are removed from disk during that launch.

## Release updates

GitHub is used only as an anonymous, read-only release source. The repository is
fixed in the application; there is no user-configurable GitHub repository or
token. The release check expects an asset named for the release version and the
installer verifies the extracted package structure and embedded application
version before handing over to Setup.

The daily check runs in a separate runspace so a slow network cannot freeze the
Dashboard. Its timestamp is persisted in the settings file.

## Startup and incremental metadata

The startup splash reports component loading, configuration/rule loading,
monitoring-engine initialization and the initial workbook-metadata scan. During
the scan it displays the current workbook and item count. Metadata I/O runs in a
separate PowerShell runspace and sends results to the UI through a concurrent
queue; the UI thread only drains that queue and repaints cached values. The
foreground dialog remains until the initial scan completes. The Dashboard,
tray icon, UI timer, first-run welcome and logon prompt are all
held back until the worker has returned a result (including an explicit
unavailable result) for every configured workbook. This prevents partially
populated Dashboard state and prevents startup dialogs from competing.

## DPI and display scaling

The application uses a DPI/display strategy intended to remain usable across common Windows scaling settings, including high-DPI/Retina virtual-machine environments.

Dialogs are constrained to the current Windows working area when necessary, with scrolling enabled for oversized forms.

The process is deliberately left DPI-unaware: the forms use fixed 96-DPI coordinates, and system-DPI awareness caused runtime-created controls to disagree with the layout at 150-200%. Sharpness is handled outside the process instead - the launchers set `__COMPAT_LAYER=GdiScaling` (equivalent to *System (Enhanced)*) so Windows renders GDI text at the real resolution without changing any coordinate the code uses. Nothing in the application depends on it.

## Launchers

`Start-Hidden.vbs` is the normal launcher and avoids a visible console.

`tools\Start-WithConsole.cmd` is a diagnostic launcher that exposes PowerShell startup errors. `Setup.cmd` runs `tools\Setup.ps1`, a WinForms installer that copies the package to `%LOCALAPPDATA%\ExcelQueryTrigger`, writes the shortcuts and the `HKCU\...\Run` value, and can remove all three again.

The installer copies runtime files to `%LOCALAPPDATA%\ExcelQueryTrigger`.
