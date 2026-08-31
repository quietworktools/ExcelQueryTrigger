# Safety and Operational Notes

## Safety goals

The application is designed around four principles:

- Do not save a partially refreshed workbook merely because a refresh is taking longer than expected.
- Do not silently run workbook macros when the user intended them to be blocked.
- Do not overwrite a workbook that appears to be in conflicting interactive use.
- Always provide an explicit user-controlled way to cancel a refresh that is genuinely stuck.

## Long refreshes are not automatically killed

The **Warn after** value is a warning threshold.

When it is reached:

- a warning is recorded;
- the Current Job state can indicate that the refresh is taking longer than expected;
- the application continues waiting;
- Excel is not automatically killed;
- Power Query is not automatically cancelled;
- a partially refreshed workbook is not intentionally saved.

## Cancel Job

Because some external systems can hang indefinitely—authentication prompts, network outages, unavailable gateways, or stalled data sources—the Dashboard provides **Cancel Job**.

Cancellation requires explicit confirmation.

A cancelled operation is kept separate from the normal successful-save path.

Cancellation is available only before saving starts. During **Saving**,
**Closing**, and **Releasing Excel**, Cancel and Exit wait for the
safety-critical write/close step to finish.

### When Cancel does not take effect

Cancellation is checked between calls into Excel. It cannot interrupt Excel while
it is inside one, so Cancel Job is a request to stop rather than an immediate stop.

With synchronous queries the `RefreshAll` call itself blocks until Excel returns
from it. During that time the cancellation is accepted but cannot be acted on, and
the elapsed time keeps rising.

When Excel stops answering altogether - typically because a data source provider
raised its own credential dialog - the cancellation escalates, and only after the
user has confirmed Cancel Job:

- Outside cancellation, an answerable dialog may be closed after the configured
  15-second grace period. After cancellation, that grace is reduced to one second.
- After about 12 seconds, the user is asked whether to end the dedicated process.
  The answer is never assumed and the option is unavailable after saving starts.

Closing a dialog and ending the process are separate permissions. A dialog raised
while Excel is writing the workbook is still closed, because Excel has written
nothing while it waits for an answer; the process itself is never terminated at
that point.

Terminating an Excel whose workbook did not confirm closed is otherwise refused.
The exception is limited to a cancellation confirmed to still be before the save
step. A save that started but did not return is never force-terminated.

The owned process id is derived from `Excel.Application.Hwnd`, not from a snapshot
of all Excel processes. Its process start time is recorded and checked again
before termination so a recycled process id is not sufficient.

### Data source check before refreshing

Every external connection is inspected before any query runs. A server name that
does not resolve, or an explicitly named port that refuses connections, stops the
job with a clear message rather than letting the provider raise a modal dialog
inside a COM call.

The check errs towards doing nothing. It acts only on values that positively
look like a host name or an IP address, so Power Query connections
(`Data Source=$Workbook$`), file, UNC and web sources are skipped rather than
guessed at. Only a definite *no such host*, seen twice, blocks a job; any other
resolver failure is a warning and the refresh proceeds. It reads connection
strings and opens a short-lived TCP probe, and never writes to the workbook.
`checkDataSources` in `appSettings` turns it off entirely.

This is a deliberate limitation rather than a defect. There is no safe way to
interrupt a synchronous COM call in progress, and terminating Excel mid-write
risks the very data loss the rest of these rules exist to prevent.

## Exiting during a refresh

The application does not trap the user in an unexitable state.

If Exit is requested before saving, the user is asked whether the active job
should be cancelled and the application closed. During saving/closing, Exit waits
for that step to finish safely.

Declining leaves the refresh running.

## Excel process ownership

The application tracks the dedicated Excel instance it starts. Cleanup logic is intended to act only on the application-owned Excel process, not on the user's normal interactive Excel session.

Forced process cleanup is deliberately conservative.

## Nothing the application reads can block a save

Excel saves an existing workbook by writing a temporary file beside it and then
replacing the original, so any process holding that file open can make the save
fail at its last step. The Dashboard reads workbook files itself for the
`Data updated` and `Where` columns, for the recent-refresh guard and for
`Test data sources`. Those reads open the file in a mode that still permits
Excel to replace it, skip a workbook whose timestamp and size are unchanged, and
are not started at all while a refresh job is running.

## Workbook locking and read-only protection

Before updating a workbook, the application checks for conditions that indicate conflicting use. If Excel opens the target read-only or the workbook cannot be safely written, the refresh should fail rather than overwrite another user's work.

Do not manually open the same workbook while the application is refreshing it.

## BackgroundQuery restoration

Where synchronous refresh is used for reliability, supported `BackgroundQuery`
settings are treated as temporary execution settings and restored before normal
save. If restoration cannot be confirmed, the workbook is closed without saving.

## Macro policy

Default: **blocked**.

Macro-enabled workbooks can require `Workbook_Open` VBA to build connections, parameters, or query source paths. The application therefore allows an explicit opt-in for those workbooks.

The opt-in does not bypass Excel or company macro security.

## Ask before refresh

For expensive or operationally sensitive rules, enable **Ask before refresh**.

The trigger can still be detected automatically, but Excel is not opened until the user approves the refresh.

Choosing No skips only the current occurrence.

## File monitoring boundaries

File monitoring begins when the watcher starts and does not replay historical changes.

This is intentional. Users should not assume that a file delivered overnight while the application was stopped will automatically be processed the next morning.

## Other Excel work

The dedicated Excel instance normally allows other workbooks to remain usable in the user's normal Excel session.

Performance can still be affected by shared CPU, memory, storage, network bandwidth, or common data sources.

## Recommended operational practice

- Use `Start-Hidden.vbs` for normal use.
- Keep **Ask before refresh** enabled for very heavy or timing-sensitive jobs.
- Review the result popup or log when running an unsaved one-off refresh.
- Use **Cancel Job** only when you intentionally want to stop an active refresh.
- Never edit the same workbook while it is being refreshed by the application.
