# Excel Query Trigger Manager

A Windows utility for automatically refreshing Excel Power Query workbooks in the background.

**Current version: v1.0.3**

## Download

### [Download the latest version](https://github.com/quietworktools/ExcelQueryTrigger/releases/latest)

Download the release ZIP, extract it, and double-click `Setup.cmd`.

## 1. What it does

Excel Query Trigger Manager automates Excel workbooks that would otherwise need to be opened and refreshed manually.

A trigger rule defines **when** a workbook should run and **which Power Query queries** should refresh. The app uses a separate background Excel instance, waits for the refresh operation, and saves and closes the workbook when the operation completes successfully.

Typical triggers include:

- scheduled times
- a new file arriving in a folder
- a specific or matching file being changed
- Windows sign-in
- manual execution from the Dashboard

The app normally stays running in the **Windows system tray**, so closing the Dashboard does not stop configured trigger rules.

## 2. Quick start

1. Download and extract the release ZIP.
2. Double-click `Setup.cmd`.
3. Open the app and select **Add**.
4. Choose the workbook and the Power Query queries to refresh.
5. Choose when the rule should run.
6. Save the rule and leave the app running in the system tray.

## 3. How it works

The application uses standard Windows and Microsoft Office technologies and does not require a separate server or cloud service for Excel refreshes.

### PowerShell

**PowerShell** is Microsoft's Windows automation and scripting environment. It provides the application logic, user interface coordination, trigger rules, file monitoring, logging and installation functions.

### Excel COM Automation

**Excel COM Automation** is the Windows interface that allows software to control the desktop version of Microsoft Excel. The app uses it to start a dedicated Excel instance, open the workbook, request the refresh, monitor the operation, save when appropriate and close Excel again.

### Power Query

**Power Query remains part of Excel.** Excel Query Trigger Manager does not replace Power Query or implement its own data-processing engine. It automates when Excel runs the Power Query queries already defined in the workbook.

### File monitoring and schedules

Windows file-system monitoring and timers are used for new-file, changed-file and scheduled triggers. When the configured condition occurs, the corresponding rule is queued for execution.

### Local configuration

Trigger rules, settings, history and logs are stored locally on the Windows PC. JSON is used for structured configuration data.

## 4. Installation

The application is installed for the current Windows user at:

`%LOCALAPPDATA%\ExcelQueryTrigger`

This location allows installation and updates **without administrator rights**. The app does not install a Windows service, does not write to system-wide `Program Files`, and does not install itself for other Windows users.

A Start menu shortcut is created. Automatic startup at Windows sign-in and a desktop shortcut can also be enabled during setup.

## 5. Updates

Excel Query Trigger Manager includes an **in-app update function**.

The application can check the official GitHub Releases for a newer version and install it from within the app. Update packages are checked against the published SHA-256 checksum before Setup starts.

Existing rules, history and logs are retained during a normal update.

## 6. Uninstall

1. Exit Excel Query Trigger Manager from its **system tray icon**.
2. Run `Setup.cmd` again from an extracted package.
3. Select **Uninstall**.
4. Choose whether to keep rules and history for a future reinstall or delete everything.

Uninstall removes the program files, shortcuts and the Windows sign-in startup entry.

## 7. Important — back up your workbooks

**Always keep a backup of important Excel files before using this tool.**

Excel Query Trigger Manager includes safeguards intended to reduce the risk of accidental file damage. For example, it checks for read-only or in-use workbooks, avoids saving when refresh completion cannot be confirmed, and does not forcibly terminate its Excel process while Excel is saving or closing.

These safeguards reduce risk but cannot guarantee that a workbook will never be changed, corrupted, locked, or affected by Excel, Power Query, network, synchronization, macro or external data-source behavior.

**Use this software at your own risk.**

## Documentation

| Document | English | Japanese |
|---|---|---|
| Quick instructions | [`INSTRUCTION_EN.md`](INSTRUCTION_EN.md) | [`INSTRUCTION_JA.md`](INSTRUCTION_JA.md) |
| Full manual | [`docs/MANUAL_EN.md`](docs/MANUAL_EN.md) | [`docs/MANUAL_JA.md`](docs/MANUAL_JA.md) |
| Safety design | [`docs/SAFETY.md`](docs/SAFETY.md) | — |
| Technical reference | [`docs/TECHNICAL-REFERENCE.md`](docs/TECHNICAL-REFERENCE.md) | — |
| Release notes | [`RELEASE-NOTES.md`](RELEASE-NOTES.md) | — |
