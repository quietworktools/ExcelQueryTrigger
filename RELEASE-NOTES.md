# Release Notes

## v1.0.3 — Reliability and dashboard improvements

**2026-09-02**

- Improved file-trigger reliability, including better handling of files that are
  renamed or replaced shortly after they are created.
- Improved trigger grouping so several related file events arriving together are
  handled as a single refresh request instead of separate ones.
- Improved network-folder monitoring so temporary Wi-Fi, Ethernet and VPN changes
  recover automatically without unnecessary error messages.
- A network folder that becomes unavailable now only pauses the rules that use
  it. Other network folders and all local folders keep watching, and monitoring
  resumes on its own once the folder can be reached again.
- Improved Windows sign-in triggers and startup handling, including a warning if
  a sign-in rule is set up while the app is not configured to start with Windows.
- When a refresh cannot start, the reason is now stated plainly - for example the
  file is still being written, is still open in another program, is no longer
  there, or the folder could not be reached.
- A trigger file that disappears because it was renamed no longer holds up a
  refresh for the full waiting time, and no longer fails the whole job on its own.
- Improved dashboard layout and fixed panels extending beyond the right edge.
- Fixed several Recent Activity and UI errors that could occur after failed or
  incomplete trigger events.
- Additional reliability improvements around Excel refresh completion, logging
  and watcher recovery.

## v1.0.2 — Repository hygiene and privacy cleanup

**2026-09-01**

- Removed runtime-generated logs and local state files from the public source tree and release package.
- Added safeguards that prevent local settings and runtime files from entering future releases.
- Reset the public Git history to remove local paths and personal commit metadata.
- No application behavior changes from v1.0.1.

## v1.0.1 — Stability and rule ordering

**2026-08-31**

- Fixed valid Power Query refreshes being rejected before save when Excel or
  the storage layer changed the workbook's file size or timestamp during the
  refresh.
- Workbook safety now relies on the existing preflight in-use check, Excel's
  writable-open check, and the lock held by the dedicated Excel instance until
  the workbook is saved and closed.
- Trigger Rules can now be dragged into a new order. The order is saved to the
  configuration and restored after restart.
- Every Trigger Rules column can be sorted ascending or descending by clicking
  its heading. Date, next-run and duration columns use their actual values.

## v1.0.0 — Initial release

**2026-08-31**

A Windows tray application that refreshes Excel (Power Query) workbooks when a
file changes, at a scheduled time, at Windows logon, or on demand. No installer,
no administrator rights, no additional PowerShell modules.

### What it does

- **Triggers** — a new file in a folder, a specific file updated, any matching
  file updated, a scheduled time, Windows logon, or manual only.
- **Refresh** — all queries (`RefreshAll`) or selected Power Query connections,
  chosen per workbook.
- **Multiple workbooks** — one rule can refresh several workbooks, one at a
  time, in the order they are listed.
- **Dashboard** — one row per rule, showing where the workbook lives, the query
  scope, the next scheduled run, how long the last run took, and when the data
  itself was last refreshed by anyone.
- **Operation** — runs in the system tray, so closing the Dashboard does not
  stop monitoring. In-app update checking against the public GitHub releases.

### Safety design

- Refreshes run in a dedicated hidden Excel instance, separate from the user's
  own Excel session.
- Workbook availability is checked before Excel is started. A workbook that
  another user or Excel window already has open is not refreshed.
- Macros are blocked by default. Allowing them is an explicit per-workbook
  choice and does not override Excel Trust Center or company policy.
- Every external connection's server is resolved before any query runs, so an
  unreachable data source stops the job instead of letting the provider raise a
  modal dialog inside a COM call.
- A refresh that passes the warning threshold is never killed automatically, and
  a partially refreshed workbook is never saved.
- Temporary `BackgroundQuery` changes are restored before the workbook is saved.
  If restoration cannot be confirmed, the workbook is closed without saving.
- If the workbook changed on disk during a long refresh, the refreshed copy is
  not saved over the newer file.
- Cancellation and forced process termination are both withheld once Excel
  begins writing the workbook. A modal dialog raised at that point is still
  closed, so the job fails with a reported error instead of waiting for an
  answer nobody can give.
- Reading workbook metadata for the Dashboard never blocks Excel from replacing
  a file it is saving, and is not performed at all on a workbook that a running
  job has open.

### Requirements

Windows 10 or 11, Windows PowerShell 5.1, and a desktop installation of Excel
with the workbook's Power Query connections already configured.

### Important

Always keep backups of important workbooks. The application includes safeguards
intended to reduce the risk of file damage, but Excel, Power Query, macros,
network connections, synchronization software, and external data sources can
behave unexpectedly.

Use this software at your own risk.
