# RVTools Upgrade Advisor

Reads RVTools exports and builds an upgrade-readiness report for **VCF 9.1** or **VVF 9.1**.
You get one HTML report (a page per environment, with a left-hand menu) and one Excel workbook.

No modules to install. Works on Windows PowerShell 5.1 and PowerShell 7.

## Run it

```powershell
# One or more files, or a folder
.\Invoke-RVToolsAdvisor.ps1 -Path C:\rvtools\*.xlsx -Target 'VCF 9.1' -Features vSAN,NSX -Customer 'Contoso'

# Ask me questions (target, features, customer)
.\Invoke-RVToolsAdvisor.ps1 -Path .\exports -Interactive
```

Output goes to `.\RVTools-Upgrade-Report` unless you set `-OutputPath`.

Features you can pick: vSAN, NSX, VCF Operations, VCF Automation, vSphere Kubernetes Service.

## Several environments

Each vCenter found in the exports is its own environment. Give it several files, or one file with several vCenters in it.
The HTML report has a menu entry per environment. The Excel workbook has an Overview sheet with links to one sheet per environment,
plus combined sheets (Findings, Hosts, VMs and so on) with an Environment column you can filter.

## What it checks

- ESXi and vCenter version against the supported starting point (7.x has to go to 8.0 first)
- Host CPU against Broadcom's discontinued and deprecated lists (KB 318697)
- Server model against HCL data you have loaded (see below)
- Clusters: single-host, mixed versions, DRS off
- VMs: old snapshots, old hardware versions, RDMs, VMware Tools, guest OS past end of support, consolidation needed
- Datastores: space and old VMFS versions
- Licences: expiry, plus the move to subscription licence files
- Sizing: cores, licensed cores (Broadcom's 16-core-per-CPU minimum), vCPU ratio, memory, storage

Each finding is a **Blocker**, **Warning** or **Info**. An environment is Blocked, Ready with actions, or Ready.

## The HCL check

The Broadcom Compatibility Guide is a web app, so the script does not scrape it. The report lists every vendor, model and CPU in the HCL tab
so you can search them there. To have the report mark models as Listed or Not listed:

1. Export the server list for the target release from the Broadcom Compatibility Guide as CSV.
2. Run `.\Invoke-RVToolsAdvisor.ps1 -ImportHcl .\export.csv -HclTarget 'VCF 9.1'`

That adds rows to `data\hcl-servers.csv`. The import looks for a model column, a vendor column and a release column,
and keeps rows that mention the target release. Check the first import by eye, because Broadcom's export headings may differ.
Until you load data, hosts show as Not checked.

## Rules and updating them

All the CPU rules, target notes, feature checklists and source links are in `data\targets.json`.
They were checked on 2026-09-29. Broadcom changes them, so re-check the sources listed in the report before using it with a customer.
Guest OS end-of-support and hardware version limits are advisory rules I set, not Broadcom rules. Change them in the file if you disagree.

## Limits

- RVTools cannot see lifecycle method (baselines vs images), Host Profiles, NSX, firmware or third-party products. The Notes tab lists these as manual checks.
- NSX detection is a guess from network names.
- Pre-Broadwell CPUs are marked Blocker by inference (Broadwell is already discontinued). Confirm on the HCL.
- Tested on PowerShell 7 with generated RVTools-style files. It has not been run against real customer exports or on Windows PowerShell 5.1 yet.
