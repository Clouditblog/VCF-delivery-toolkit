# sos-report-to-excel

Turns a VCF SDDC Manager SoS `--health-check` text report into an Excel workbook.

- One worksheet per check table, named from the line above it (NTP, Password Check and so on)
- GREEN / YELLOW / RED cells coloured
- A Summary sheet at the front totals each state

## Use

Needs the ImportExcel module (one-off): `Install-Module ImportExcel -Scope CurrentUser`

```
powershell -NoProfile -ExecutionPolicy Bypass -File .\Sos-Report-To-Excel.ps1 -InputFile .\sos-health-check.log
```

`sample_sos.log` is a made-up report to try it on. The parser expects the `+---+` box-table layout; if your SoS version prints something different, the tables won't be found.
