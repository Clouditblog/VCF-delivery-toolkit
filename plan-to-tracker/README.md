# plan-to-tracker

Turns a VCF implementation plan (Word, .docx) into an Excel tracker with three tabs:
a Dashboard for the project manager, a Tracker with Status and RAG dropdowns, and an Issues Log.

There are two versions of the same tool.

| Version | Needs | Files |
|---|---|---|
| PowerShell | Windows PowerShell 5.1 or PowerShell 7. Nothing else: no Python, no modules, no Excel, no Word | `plan_to_tracker.ps1` and the three `.cmd` launchers |
| Python | Python with `python-docx` and `openpyxl` | `plan_to_tracker.py` |

`raw_proto.py` is the Python prototype of the raw-XML writer that the PowerShell script uses.

## Quick start (PowerShell, from Windows)

Keep `plan_to_tracker.ps1` and the `.cmd` files together in one folder.

**One plan:** drag the .docx onto `Make-Tracker.cmd`. The tracker appears next to the Word file.
From a cmd prompt: `Make-Tracker.cmd "C:\Plans\My Plan.docx"`.

**Drop folder:** double-click `Watch-DropFolder.cmd` and leave the window open. It creates a `DropHere` folder next to the scripts.

| Folder | What lands there |
|---|---|
| `DropHere` | Put the Word plan here |
| `DropHere\Trackers` | The finished tracker |
| `DropHere\Processed` | The Word plan, once it has been turned into a tracker |
| `DropHere\Failed` | Anything that could not be read, with a `.txt` note saying why |

To watch another folder: `Watch-DropFolder.cmd "D:\Plans\DropHere"`.
To process what is there and stop (for Task Scheduler): `Process-DropFolder-Once.cmd "D:\Plans\DropHere" nopause`.

An existing tracker is never overwritten, because the team will have been filling it in. A second run on the same plan makes a new file with a timestamp in its name.

## What it takes from the plan

- The Preparation checklist table and the General Requirements tables become Pre-check rows, plus a go or no-go checkpoint
- Each Execution Plan table (Action, Impact to Environment, Required, VM Downtime) becomes a phase with one row per action
- Actions that start with "Validate" become Checkpoints
- Required = No is marked "(optional)"

If the document has no Execution Plan tables (for example a procedure exported from Broadcom's documentation), the script falls back to the numbered steps: one phase per main heading, one row per top-level numbered step, with the number of sub-steps in the Notes column. Bullets are ignored, and a main heading with no numbered steps becomes a single "Complete: ..." row. The console says when this happened. Check the rows over before using them, because the source steps are written for a person following the document, not for tracking.

If neither is found, the error lists the first row of every table in the document, which shows what the script saw.

The Dashboard recalculates from the Tracker and Issues Log tabs. Nothing to type on it.

## Time estimates (PowerShell version)

The tracker has an **Est. (hrs)** column and an **Actual (hrs)** column, and a **Sizing** tab behind them.

1. On the Sizing tab, type the size of the environment: ESX hosts, clusters, vCenter Servers and NSX Edge nodes.
2. Each tracker row has an **Estimate basis** (a dropdown, for example "ESX host upgrade" or "Appliance deploy"). The hours are the fixed hours for that basis plus the hours per unit multiplied by the number of units. A 12-host site gets more ESX upgrade time than a 4-host site without anyone retyping anything.
3. To override one task, type a number over the formula in Est. (hrs). To change what a type of task costs, edit the table on the Sizing tab.
4. Actual (hrs) is worked out from Actual start and Actual end.
5. The Dashboard shows estimated, actual and remaining hours for each phase, and a calibration line once some items are complete with actual times, for example "completed items took 133% of their estimate". Use that to correct the Sizing table before the next environment.

The script picks each row's estimate basis from keywords in the task text (for example "Upgrade ESX hosts" or "Deploy the appliance"), so some rows will be wrong. The hours in the Sizing table are starting values, not measured timings. Replace them with your own.

This is in the PowerShell version only. The Python version does not have it.

## Suggested change window times (PowerShell version)

Three more columns on the tracker (O to Q) give a suggested start, finish and window for each task, and a **Change Windows** tab turns them into windows you can put to the customer.

1. Type **Window start** on the tracker (cell C7), for example Friday 19:00. Until you do, the times stay blank.
2. Tasks are laid out one after another in tracker order, each taking its Est. (hrs).
3. A window holds the hours set on the Sizing tab (default 8, with 20% kept free at the end). A task that does not fit in the hours left starts the next window. Tasks are never split.
4. Next windows fall on the next working day (weekends skipped) at the same time of day. Set Skip weekends to No on the Sizing tab for consecutive days.
5. If you fill in both Window start and Window end on the tracker, the window length is taken from them instead.
6. The Change Windows tab lists each window with its start, end, the time its tasks finish, spare hours, first and last task, and a sentence you can paste into an email. A key on that tab states what the times are based on.

Not allowed for: bank holidays, change freezes, approval waits, parallel work and rollback time. Check those against the customer calendar. The times are a starting suggestion, only as good as the hours on the Sizing tab.


## If it will not run

- **"Running scripts is disabled" or similar.** The launchers already pass `-ExecutionPolicy Bypass`. If a company policy still blocks it, try this from a cmd prompt, which reads the script as text instead of running it as a file:
  `powershell -NoProfile -Command "& ([scriptblock]::Create([IO.File]::ReadAllText('C:\Tools\plan_to_tracker.ps1'))) -Docx 'C:\Plans\My Plan.docx'"`
- **Errors mentioning "language mode" or "Cannot invoke method".** PowerShell is in Constrained Language Mode (AppLocker or similar). The script uses .NET file and zip classes, which that mode blocks. Check with `powershell -Command "$ExecutionContext.SessionState.LanguageMode"`; it should say `FullLanguage`.
- **A file moved to Failed.** Open the `.txt` note next to it. The usual causes are an old `.doc` renamed to `.docx`, or a plan with no Execution Plan table.
- **The Dashboard looks empty when opened from email or a download.** Excel's Protected View shows saved values only. Click Enable Editing and it fills in.
- **Excel offers to repair the file.** Raise an issue with the exact message and the table headings from the Word plan.

## Tested

The PowerShell script was run on PowerShell 7 on Linux against a real plan and a Broadcom procedure export. Its workbook matches the Python version cell for cell, and recalculates with no formula errors. It has not been run on Windows PowerShell 5.1 or opened in desktop Excel yet.
