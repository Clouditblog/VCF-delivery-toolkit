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

The Dashboard recalculates from the Tracker and Issues Log tabs. Nothing to type on it.

## If it will not run

- **"Running scripts is disabled" or similar.** The launchers already pass `-ExecutionPolicy Bypass`. If a company policy still blocks it, try this from a cmd prompt, which reads the script as text instead of running it as a file:
  `powershell -NoProfile -Command "& ([scriptblock]::Create([IO.File]::ReadAllText('C:\Tools\plan_to_tracker.ps1'))) -Docx 'C:\Plans\My Plan.docx'"`
- **Errors mentioning "language mode" or "Cannot invoke method".** PowerShell is in Constrained Language Mode (AppLocker or similar). The script uses .NET file and zip classes, which that mode blocks. Check with `powershell -Command "$ExecutionContext.SessionState.LanguageMode"`; it should say `FullLanguage`.
- **A file moved to Failed.** Open the `.txt` note next to it. The usual causes are an old `.doc` renamed to `.docx`, or a plan with no Execution Plan table.
- **The Dashboard looks empty when opened from email or a download.** Excel's Protected View shows saved values only. Click Enable Editing and it fills in.
- **Excel offers to repair the file.** Raise an issue with the exact message and the table headings from the Word plan.

## Tested

The PowerShell script was run on PowerShell 7 on Linux against a real plan. Its workbook matches the Python version cell for cell, and recalculates with no formula errors. It has not been run on Windows PowerShell 5.1 or opened in desktop Excel yet.
