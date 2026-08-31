# Excel Query Trigger Manager — quick steps

## 1. Install

**Before first use, back up important workbooks.** Closing the Dashboard does
not stop the program; it remains active in the system tray.

Double-click **`Setup.cmd`** and follow the window. That is the whole
installation.

It installs to `%LOCALAPPDATA%\ExcelQueryTrigger`, adds a Start menu entry, and
offers to start the app when you sign in to Windows. No administrator rights are
needed. Running `Setup.cmd` again updates the app and keeps your rules, history
and logs; it can also remove the app.

## 2. Add your first file

Press **Add** in the Trigger Rules box and answer three questions, one per
screen:

1. Which Excel file.
2. When it should update - at a set time, when a file arrives in a folder, or
   only when you press Update now.
3. Does the summary look right.

That is the whole setup. Everything else takes its default. **Edit** opens the
full rule editor, and **New rule with all options...** on the right-click menu
creates one that way from the start.

**Queries...** decides what gets refreshed in a workbook: everything, or the
ones you pick. The **Where** and **Data updated** columns show whether a
workbook is shared and when its data was last refreshed by anyone.

## 3. Day to day

- Closing the window does not stop the app. It stays in the tray by the clock.
- Click the tray icon once to bring the dashboard back.
- **Hide to Tray** keeps it watching. **Exit** stops it.
- The rule refreshing right now is shaded green; rules waiting behind it, amber.
- **Run Now** becomes **Add to Queue** while a refresh is running.
- Use **Show all / Compact** at the right of the **Trigger Rules** title; the app
  calculates and remembers the appropriate height.
- Drag a Trigger Rules row to change its saved order. The order is restored at
  the next launch.
- Click any Trigger Rules heading to sort ascending; click it again for
  descending. Column sorting changes only the current view.
- Click the version number at the lower left to check for a newer public
  release. No account or token is used.

## 4. Worth knowing before you rely on it

**It only watches while it is running.** A file that arrived overnight, with the
app closed, is not picked up when you start it the next morning.

**Your Excel is not disturbed.** Refreshes run in a separate, hidden Excel. Just
do not open or edit the same workbook while it is being refreshed.

**Warn after is a warning, not a time limit.** The app keeps waiting for Excel
past that point. **Cancel Job** asks Excel to stop; if Excel does not answer,
any dialog it is stuck on is closed and you are then asked whether to end that
one Excel process. Cancel is disabled while Excel is saving or closing so a
workbook cannot be terminated mid-write.

**Right-click a rule → Test data sources** checks each connection in its
workbooks without opening Excel — useful when a refresh fails and you want to
know why.

## 5. More

The **i** button at the top right of the dashboard, or `docs/MANUAL_EN.md`.
