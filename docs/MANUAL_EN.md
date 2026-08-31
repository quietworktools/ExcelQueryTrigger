# Excel Query Trigger Manager — User Manual

**Version 1.0.2**

## 1. Overview

Excel Query Trigger Manager is a Windows tray application for refreshing Excel workbooks automatically or semi-automatically.

A rule can:

- watch for a new file;
- watch for a specific file update;
- watch for matching file updates;
- run at a scheduled time;
- run at Windows logon;
- run manually;
- refresh one or several workbooks in sequence;
- refresh all queries or selected queries;
- ask for approval before an automatic refresh.

The application is intended to run without administrator privileges and uses Windows PowerShell 5.1 and Microsoft Excel.

---

## 2. Installation and launch

### Installing

Before first use, back up important workbooks. Closing the Dashboard does not
stop the program; it remains active in the system tray. The first launch after
a new installation opens the Dashboard in front so this behavior is visible.

Unzip the folder anywhere and double-click **`Setup.cmd`**. That is the whole
installation - a small window opens and does the rest.

It installs into `%LOCALAPPDATA%\ExcelQueryTrigger`, which is per-user,
writable without administrator rights, and outside any OneDrive folder. It also
clears the block Windows puts on files from a downloaded zip, adds a Start menu
entry, optionally puts a shortcut on the desktop, and optionally registers the
application to start when you sign in.

Run `Setup.cmd` again to update: your rules, history and logs are kept. The same
window removes the application, asking first whether to keep your rules.

### Starting it

Use the **Start menu** entry, or the desktop shortcut if you asked for one.
Both point at `Start-Hidden.vbs` in the installed folder, which starts the
application with no console window.

`Start-AtLogon.vbs` is the one registered for sign-in. It is the same launcher
with one difference: it tells the application that this was a sign-in start,
which is what allows the logon refresh prompt to appear.

### When it will not start

`tools\Start-WithConsole.cmd` starts the application with the console window
visible, so a PowerShell startup error stays on screen. It is for
troubleshooting only.

### Start automatically at logon

Setup offers this, and Settings inside the application can turn it on or off
later. It writes a single value under `HKCU\...\Run` - never HKLM, and never a
service.

---

## 3. Dashboard

### Status

Shows whether monitoring is running or paused.

### Trigger Rules

Each row represents one saved rule and shows:

- rule name;
- trigger type;
- workbook file name;
- **Where** - whether the workbook is local, on a network share, in a
  OneDrive/SharePoint folder, or a shared workbook;
- status;
- **Queries** - the query scope, `All` or `3 chosen`;
- last run;
- **Next Run** - for scheduled rules;
- duration;
- **Data updated** - when the data itself was last refreshed, by anyone.

Hover a cell for the detail behind it; each column has its own tooltip.

Main commands:

- **Add**
- **Edit**
- **Delete**
- **Enable / Disable**
- **Run Now**
- **Ask before refresh**

**Add** opens a three-step rule wizard. First choose one of the six trigger
types. Then add one or more workbooks; each workbook has its own query selection
and refresh options, and the list order is the execution order. Finally, name
and review the complete rule before **Add it** saves it.

Right-click a rule for additional actions such as opening monitored locations, opening workbook locations, refreshing a workbook, or running the entire rule.

Use **Show all / Compact** at the right of the **Trigger Rules** title. Show all
calculates the height from the number of rules and the available window space;
the selected mode is restored at the next launch.

Drag a rule row to move it anywhere in the list. A blue insertion marker shows
the destination, and the dropped order is saved to `rules.json` and restored
at the next launch.

Click any column heading to sort it ascending, then click the same heading
again for descending. **▲ / ▼** shows the current direction. Text fields use
their displayed content; **Last Run**, **Next Run**, **Duration** and
**Data updated** use real date/time or numeric values. Sorting changes only the
current view and does not rewrite the manual order in `rules.json`. Dragging a
row while sorted turns the visible order into the new saved manual order and
clears the column sort.

### Refresh Any Excel File...

Use this for a one-off refresh that is not tied to a saved rule.

A review screen appears before execution so an accidental file selection can be changed or cancelled.

### Current Job

Shows the active rule/workbook, current stage, elapsed time, and **Cancel Job** while a refresh is active.

### Pending Jobs

Shows jobs waiting behind the current job. Jobs are processed sequentially.

### Recent Activity

Shows monitoring, trigger, refresh, warning, success, skip, cancellation, and error events.

Select a row to see details.

### System tray

- Left-click once: restore and activate the Dashboard.
- Right-click: open the tray menu.
- Hide to Tray: hide the Dashboard while monitoring continues.

---

## 4. Trigger types

### New file added to folder

Runs when Windows reports the arrival of a new matching file.

### Specific file updated

Runs when a configured file changes.

### Matching file updated

Runs when a matching file in a folder changes.

### Scheduled time

Runs at the configured time on selected days while the application is running.

### At Windows logon

Uses application startup after Windows logon.

### Manual only

Never fires automatically.

---

## 5. File monitoring model

Monitoring uses Windows `FileSystemWatcher`.

It is event-driven:

`Windows file event → type/name filter → debounce → file-ready check → cooldown → confirmation/job queue`

The application does not repeatedly scan all files in a monitored folder.

### Debounce

One file copy can generate several Windows events. Debounce combines closely spaced duplicate events into one logical trigger.

### Cooldown

After a rule fires, the minimum repeat interval prevents the same rule from immediately firing again.

### File-ready check

A newly created file may still be copying. The application can wait until the file appears stable and ready before starting Excel.

### Monitoring health check

The application periodically checks that the watcher is still active and that the monitored location remains reachable.

This is not a full folder scan.

---

## 6. Monitoring does not look backward

This is an important operational rule.

The application observes file events only while monitoring is active. It does not compare yesterday's folder contents with today's folder contents at startup.

Example:

- 20:00 — the application stops.
- 22:00 — a CSV file arrives.
- 08:30 next morning — the application starts.

The file added at 22:00 does **not** trigger a refresh at 08:30.

Only matching events that occur after monitoring has started can trigger the rule.

---

## 7. Ask before refresh

Enable **Ask before refresh** for heavy or timing-sensitive jobs.

Flow:

`Trigger detected → approval dialog → Refresh now / Skip`

When the dialog is displayed, Excel has not yet been opened for that job.

### Refresh now

The job enters the normal queue and Excel refresh begins.

### Skip

Only the current trigger occurrence is skipped.

The rule stays enabled and monitoring continues. The next matching event can ask again.

This is useful at the end of the day when a long refresh would be inconvenient.

### Ask only after a recent refresh

On the Add wizard's workbook-selection page, or in the rule editor, enable
**Ask before refreshing if updated within** under **Recent refresh protection**.
Choose a number of minutes or hours. Before an automatic file, schedule, or
logon run, the application reads each target workbook in a background worker:

- when Excel records a query/pivot refresh time, that is used;
- successful refreshes by this application store a separate query-refresh time
  in the workbook so it remains available after restart and on another PC;
- a file-save time is never treated as a query refresh;
- if every workbook is older than the selected window, the job is queued
  automatically;
- if any workbook is recent, or a workbook cannot be checked, a dialog asks
  whether to refresh again.

This reduces duplicate refreshes on shared files, but it is not a cross-PC lock.
Two computers can still start at the same instant before either one saves the
workbook. Use a shared lock/coordination service if strict mutual exclusion is
required. **Run Now** remains an explicit action and does not use this guard.

After a successful automatic or manual job, the Dashboard immediately re-reads
the affected workbook metadata in the background and updates **Data updated**.
It does not wait for the normal periodic metadata scan.

**Data updated** therefore means query/refreshable-data completion time, not the
time someone last edited or saved the workbook. If neither Excel nor this
application has stored a refresh time, the Dashboard shows **Not recorded**.

---

## 8. Excel actions

One rule can contain multiple workbook actions. They are processed in the configured order.

The detailed rule list shows only the choices that can vary by workbook:
workbook, query scope, long-run warning, and what happens if that workbook
fails. Excel runs hidden, saves after a successful refresh, and closes after the
action; these fixed behaviors are not shown as editable columns.

Common settings include:

- workbook path;
- refresh all queries or selected queries;
- warning threshold;
- continue after an error;
- waiting for supported connections before saving;
- macro policy.

Excel always runs hidden for these actions. After a successful refresh, the
workbook is saved and the dedicated Excel instance is closed.

---

## 9. All queries vs selected queries

### All queries

Refreshes the workbook's normal refresh entry points.

### Selected queries

Lets you choose specific Power Query entry points.

Power Query dependencies still apply. If a selected query depends on upstream queries, Power Query may evaluate those dependencies as part of the refresh.

The result popup and logs show query names for practical verification.

---

## 10. Using Excel while a refresh is running

You can normally continue using other Excel workbooks.

The application creates a dedicated Excel instance for its refresh job rather than attaching to your normal interactive Excel session.

Do not manually open or edit the **same workbook** that the application is currently refreshing.

Heavy Power Query processing can consume CPU, memory, disk, or network bandwidth, so other work may temporarily feel slower.

---

## 11. Refresh completion and long-running jobs

The application does not use a fixed sleep and then save blindly.

It checks supported Excel refresh/calculation states and waits for activity to become consistently quiet.

### Warn after

The configured warning time is not a hard timeout.

When reached:

- a warning is logged;
- the job can show that it is taking longer than expected;
- Excel is not automatically killed;
- Power Query is not automatically cancelled;
- the application keeps waiting.

### Choosing individual queries

**Selected queries** lists what the workbook can refresh on its own, one row
each, with the type (**Power Query** or **Connection**) and where it lands:

- **Table** - it fills a table on a sheet
- **PivotTable** - it feeds a PivotTable
- **Data Model** - it loads into the workbook's Data Model
- **Connection only** - it loads nowhere; Power Query's own scaffolding
  (*Sample File*, *Transform Sample File*, *Parameter1*) lands here, whatever
  language your Excel is in
- **Internal** - Excel's own plumbing (*ThisWorkbookDataModel*,
  *WorksheetConnection_...*)
- **Unknown** - this workbook did not let us read where the item loads

Everything except *Connection only* and *Internal* is shown by default; **Also
show staging, helper and internal items** reveals the rest with a count.
Anything already selected stays on screen either way, so it can still be
unselected. Where the workbook exposes nothing about load targets, every item
is listed and marked *Unknown* rather than guessed at.

Refreshing a query still evaluates the upstream queries it depends on, so you
rarely need the hidden ones.

If a row looks wrong, **Copy list** puts the whole table on the clipboard with
each item's type, load target and underlying connection name.

### Test data sources

Right-click a rule and choose **Test data sources...**. For every workbook in
the rule it lists each stored connection, the server it points at, and whether
that server could be reached, with the full connection string shown for the
selected row and a **Copy report** button.

The check reads the workbook file directly - a workbook is a zip, and its
connections live in `xl/connections.xml`. It never opens Excel, so it works on a
workbook Excel cannot open at all and cannot itself raise the credential dialog
it exists to diagnose. It reads through a copy taken in a way that leaves Excel
free to save the original at the same time.

What it can tell you:

- the server name no longer exists on the network
- the server answers but nothing is listening on the port the connection names
- the connection refreshes as the file opens, which is why a prompt can appear
  before the refresh step ever runs
- the connection is a Power Query one, where the real source is inside the
  query and cannot be read from here

What it cannot tell you: whether the cube, table or credential behind a
reachable server is still valid. For that, run the rule and watch the activity
log - a dialog raised by the provider is closed and its title recorded.

### When an update fails

The rule row shows the latest status. Select the corresponding Recent Activity
entry to see the full message, then use **Open Log** when deeper detail is needed.

### Adding a file

**Add** uses three screens: configure one of the six triggers, build the ordered
workbook list, then name and review the rule. Use **Add workbook...** for every
Excel file that should run from the same trigger. **Edit...** changes that
workbook's query selection and refresh options; **Move up** and **Move down**
set the execution order. The trigger screen expands for the selected type.
Folder triggers include folder,
file type/custom pattern, required filename text and excluded filename text;
the specific-file trigger asks for the exact source file; scheduled and logon
triggers show their own timing or confirmation fields. Debounce, cooldown and
the slow-refresh warning keep their defaults. To change those afterwards, select
the rule and use **Edit**.

Scheduled rules run only while this PC is turned on and the user is signed in
to Windows. A scheduled time missed while the PC is off or signed out is not
replayed later.

An identical rule cannot be added twice, even under a different name. Duplicate
detection compares the effective trigger settings and the complete ordered
workbook actions. A different workbook, query scope, refresh option, or action
order remains a distinct rule.

### Watching what is running

The rule being refreshed right now is shaded green in **Trigger Rules**, and any
rule with a job waiting behind it is shaded amber.

While a refresh is running, **Run Now** changes to **Add to Queue**: the job goes
to the back of **Pending Jobs** rather than starting immediately, because only
one Excel instance is driven at a time.

### Reading the activity log

**Show** filters by level: *Everything*, *Hide routine detail* (drops DEBUG),
*Problems only* (warnings and errors) or *Errors only*. **for rule** narrows to
one rule and **containing** matches text in the rule, workbook or message.
**Reset** clears all three. Filtering only changes what is drawn; nothing is
discarded, and the full log file is always on disk.

**Double-click a row** to open the rule it refers to.

### The rule list

**Next Run** shows when a scheduled rule is next due. Other trigger types have
no predictable next time, so the column stays empty for them.

**Duration** turns amber when the last run took at least twice as long as this
rule usually takes, with the comparison in the row's tooltip. The baseline is
the middle of the recent finished runs, so it appears only once a rule has run
at least three times.

### Reordering and cancelling queued jobs

Select a row in **Pending Jobs** and use **Move Up**, **Move Down** or
**Remove**. Each row is one job, which may cover several workbooks.

Removing asks for confirmation. It refreshes nothing, changes no file and leaves
the rule itself exactly as it is - only the queued run goes away. A job that the
engine has already started cannot be removed this way; use **Cancel Job**.

### Cancel Job

If the refresh is genuinely stuck, use **Cancel Job**.

Cancellation is explicit and requires confirmation.

Cancel is available only before saving starts. While the Current Job stage is
**Saving**, **Closing**, or **Releasing Excel**, the button is disabled and Exit
waits for that safety-critical step to finish.

A cancelled refresh does not follow the normal successful-save path.

#### When Cancel does not take effect

Cancellation is checked between calls into Excel. It cannot interrupt Excel while
it is inside one, so **Cancel Job is a request to stop, not an immediate stop**.

With synchronous queries the `RefreshAll` call itself blocks until Excel returns
from it. During that time the request is accepted but cannot be acted on, and the
elapsed time keeps rising.

The usual reason Excel stops answering altogether is a data source provider
raising its own credential or connection dialog. That is handled continuously,
not only when you cancel - see *Dialogs in a hidden Excel* below.

If you press Cancel Job and Excel still does not answer:

- Any dialog it is waiting on is closed **immediately**.
- **After about 12 seconds**, or straight away if closing the dialog did not
  help, you are asked whether to end that one Excel process. Only the instance
  identified from that Excel application's own window handle is ever ended, and
  only while the job is still before the save step,
  so the file on disk is unchanged.

#### Dialogs in a hidden Excel

A dialog inside a hidden, automation-driven Excel is always a dead end: you
cannot see it, nobody can answer it, and while it stands Excel refuses every
automation call - which is exactly why cancellation cannot reach it.

So the dashboard watches for one while Excel is opening the workbook, while the
queries are refreshing, and while the workbook is being saved. A dialog that is
still standing after fifteen seconds is closed, which turns a silent hang into
an ordinary error that is reported like any other. The dialog's title is written
to the activity log, so you can see what was asked.

Closing a dialog and ending the Excel process are separate permissions. During
**Saving** the dialog is still closed - Excel has not written anything while it
waits for an answer - but the process itself is never terminated there, because
that is the one moment when ending it could damage the file.

Only a window that is actually waiting for an answer is closed: it must have
something to click and no progress bar, and captions that name an operation in
progress - *Downloading*, *Opening*, *Saving*, *Contacting* and the like - are
left alone. Excel shows one of those while it fetches a workbook from a network
share, and closing it would abort a refresh that was working perfectly well.

Excel actions run hidden, so an unanswered blocking prompt is always handled by
this safety path instead of being left behind the Dashboard.

If the same question comes back three times in one job, the job is stopped with
an *ExcelDialogLoop* error: a data source that keeps asking will not succeed
unattended, and its connection or credentials need updating.

`autoDismissExcelDialogs`, `excelDialogGraceSeconds` and
`cancelEscalationSeconds` in `appSettings` control this.

With debug logging enabled the log records the process id, window handle and
process start time derived from the dedicated Excel application. The identity is
checked again before termination so a recycled process id is not sufficient.

#### The data source check that runs first

Before any query is touched, every external connection in the workbook is
inspected and its server name is resolved. A server that cannot be found on the
network is what makes the provider raise its wizard in the first place, so the
job stops there instead, with the connection name and server in the message.
A server that resolves but does not answer on the port named in the connection
string is treated the same way; anything less certain is logged as a warning and
the refresh is still attempted.

The check only looks at connections that positively name a network server. A
Power Query connection does not: its connection string reads
`Data Source=$Workbook$` because the real source lives inside the M formula, so
those connections are skipped and recorded in the debug log as *not checkable*.
The same applies to file, UNC and web sources.

Only a definite *no such host* answer, seen twice, stops the job. If name
resolution itself fails - a DNS outage, a VPN that has not come up - that is
logged as a warning and the refresh is attempted anyway.

To turn the check off, set `checkDataSources` to `false` in `appSettings`, or
`checkConnections` to `false` on a single Excel action, in `config/rules.json`.

### Shared-workbook locking

Before a refresh starts, the workbook is checked for use by another Excel
window or another user. If it is in use, Excel is not started and no query is
refreshed. The dedicated Excel instance then opens it for writing and keeps it
open through refresh, save and close. If Excel opens it read-only, the refresh
is not started.

On a normal file server, a user who opens the same workbook during the refresh
will receive a read-only copy. Do not manually edit the same workbook until the
application has saved and closed it.

To investigate a hidden-instance failure, use **Test data sources**, the Current
Job stage, and the activity log. The dialog title is recorded when the
application dismisses a blocking Excel prompt.

### Exit while refreshing

If you attempt to exit while a refresh is active, the application asks whether to cancel the active job and exit.

Choosing No leaves the job running.

---

## 12. Refresh Any Excel File

Use **Refresh Any Excel File...** for a one-off job.

Flow:

`Choose file → review → macro choice if applicable → all/selected queries → Refresh / Choose Another / Cancel`

The review screen prevents a mistaken file selection from immediately opening and changing a workbook.

After completion, a detailed result dialog shows:

- workbook;
- status;
- start/end time;
- elapsed time;
- save result;
- refreshed query names;
- message.

The same activity is also written to the application log.

---

## 13. Macro-enabled workbooks

Macros are blocked by default.

For `.xlsm` and `.xlsb` workbooks, the application can ask whether macros should be allowed.

Some special workbooks use `Workbook_Open` VBA to create or prepare Power Query connections, parameters, or source paths. Blocking macros can prevent those workbooks from initializing correctly.

Allowing macros does not bypass Excel Trust Center or company security policy.

---

## 14. Workbook safety

The application avoids intentionally overwriting a workbook that appears to be in conflicting use.

It checks for conditions such as owner/lock files and read-only opening.

Do not manually work in the same workbook during an automated refresh.

For supported query connections, synchronous execution settings can be applied temporarily and restored before normal save.

### Nothing the app reads can block a save

Excel saves an existing workbook by writing a temporary file next to it and then
replacing the original, so anything holding that file open can stop the save at
its last step. The Dashboard reads workbook files itself, for the **Data
updated** and **Where** columns and for **Test data sources**, and those reads
are deliberately arranged so they can never do that: the file is opened in a
mode that still allows Excel to replace it, a workbook that has not changed is
not opened again at all, and no metadata reading is started while a refresh job
is running.

---

## 15. Logs and result popups

Automatic triggered refreshes can show a detailed result popup explaining:

- why the rule ran;
- what workbook was refreshed;
- which queries were involved;
- success/failure;
- elapsed time.

Manual one-off refreshes also show a detailed result window.

Use **Open Log** for the full history.

Logs are especially useful for distinguishing:

- no trigger occurred;
- trigger occurred but Ask before refresh was skipped;
- refresh started;
- warning threshold was reached;
- refresh succeeded;
- job was cancelled;
- Excel or the data source returned an error.

---

## 16. Settings and storage

Settings are on three tabs, and every one of them carries a sentence saying what
it does:

- **Basics** - whether it starts with Windows, whether it opens its window,
  whether it lives by the clock rather than on the taskbar, and whether it
  checks quietly for updates once a day.
- **Notifications** - the pop-ups near the clock. *Tell me when a refresh fails*
  is the one worth keeping on; the rest get noisy once you trust it.
- **Advanced** - the timing numbers, whether workbooks may run macros, extra
  logging, and buttons that open the settings and log folders. The defaults suit
  almost everyone.

**Restore defaults** puts all three tabs back, without touching your rules.
*Start it when I sign in to Windows* is left alone by that, because Windows -
not the settings file - decides whether that entry exists.

### Startup progress

A progress window is shown while the application loads its components,
settings, rules and monitoring engine, and while the Dashboard reads its initial
workbook metadata. Startup details show each monitoring result, the current file,
item count, and workbook read/unavailable totals. File work runs on a dedicated
background runspace, so an unavailable share cannot freeze the dialog. It remains
visible until every configured workbook has either been read or reported
unavailable. Only then are the prepared Dashboard and tray made available. The
splash is skipped for silent Windows sign-in starts; the Dashboard normally
remains hidden there until opened from the tray.

### Version and updates

Click the version number at the bottom left, or open the **About & updates** tab
from the `i` button, to open **About**. The application
reads the newest public release from GitHub without credentials; it does not
create Issues or send reports. If a newer release is available, **Install it**
downloads the version-matched zip, validates its layout and embedded version,
then starts that package's Setup. Rules, history and logs are retained.

**Settings → Basics → Automatically check for updates** controls the daily
background check and is on by default. Turning it off does not remove manual
update checks from About.

The application runtime is installed under:

`%LOCALAPPDATA%\ExcelQueryTrigger`

Rules/settings and runtime state are stored in the local application configuration area.

The Information & Help window displays the active configuration and log paths.

Configuration schema: **4**

---

## 17. Display scaling and Parallels

The application includes layout/display handling intended for common Windows scale settings and high-DPI virtual-machine environments such as Parallels.

Dialogs are constrained to the available Windows working area and can scroll when necessary.

### Why text is sharp on a scaled display

The application is deliberately DPI-unaware. Its windows are laid out on a fixed 96-DPI grid, and letting Windows scale the whole window is what keeps that layout correct at 150-200%, including Retina virtual machines. The side effect is that Windows stretches the bitmap, which makes text look soft.

The launchers therefore set `__COMPAT_LAYER=GdiScaling` - the same thing as *Override high DPI scaling behaviour: System (Enhanced)* in a shortcut's compatibility settings. Windows then draws GDI text and lines at the real resolution while still handling the window scaling, so the layout is untouched and only the drawing improves. Some bitmap elements may still look scaled; that is expected.

To turn it off, comment out the `__COMPAT_LAYER` line in `Start-Hidden.vbs`, `Start-AtLogon.vbs`, `Setup.cmd` and `tools\Start-WithConsole.cmd`. A value already set by you or by IT is never overwritten. Nothing in the application depends on it - starting `ExcelQueryTrigger.ps1` directly simply gives the softer rendering.

---

## 18. Recommended usage

- Use `Start-Hidden.vbs` for normal operation.
- Keep Ask before refresh ON for heavy or timing-sensitive rules.
- Do not open the same workbook during refresh.
- Use Cancel Job only when you intentionally want to stop a job.
- Check the result popup or log when verifying a one-off/manual refresh.
- Remember that monitoring does not replay events that happened while the application was closed.
