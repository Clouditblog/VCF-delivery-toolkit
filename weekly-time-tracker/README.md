# weekly-time-tracker

Excel sheet for tracking hours across projects in a week.

- Columns: Project, Allocated, Remaining, Mon to Fri, Total
- Remaining turns red when a project goes over its allocation
- Total row goes green at 8 hours a day and 40 for the week, amber under, red over
- Yellow cells with blue text are the ones to fill in

`weekly_time_tracker.xlsx` is ready to use. `build_time_tracker.py` regenerates it (needs openpyxl).
