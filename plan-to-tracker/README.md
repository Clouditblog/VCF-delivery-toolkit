# plan-to-tracker

Turns a VCF implementation plan (.docx) into an Excel tracker with a Dashboard, a step tracker and an Issues Log.

- `plan_to_tracker.py` reads the Word plan and writes the tracker with openpyxl. Has a drop-folder mode (`--watch`, `--once`).
- `raw_proto.py` writes the same workbook as raw XML with no libraries. It is the model for a PowerShell port, for machines where Python is blocked.

Usage: `python plan_to_tracker.py "My Plan.docx"` or `python plan_to_tracker.py --watch <folder>`.

Status: PowerShell version (`plan_to_tracker.ps1`) not written yet.
