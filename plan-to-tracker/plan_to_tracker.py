#!/usr/bin/env python3
"""
plan_to_tracker.py

Reads a VCF Implementation Plan (.docx) and builds an Excel execution tracker
with three tabs: Dashboard, Tracker and Issues Log.

What it takes from the Word file
  - The "Preparation" checklist table            -> Pre-check rows
  - "General Requirements" / "Item + Requirements" tables -> Pre-check rows
  - Every "Execution Plan" table (Action, Impact to Environment, Required,
    VM Downtime)                                  -> one phase each, one Task row
                                                     per action. Actions that start
                                                     with "Validate" become Checkpoints.
                                                     Required = No is marked "(optional)".

Use it two ways

  1. One file:
       python plan_to_tracker.py "My Plan.docx"
       python plan_to_tracker.py "My Plan.docx" -o "My Tracker.xlsx"

  2. Drop folder (leave it running, or run it from a schedule with --once):
       python plan_to_tracker.py --watch "C:\\Trackers\\DropHere"
       python plan_to_tracker.py --watch ./DropHere --once

     Put a .docx in the drop folder. The tracker appears in <folder>/Trackers and
     the Word file moves to <folder>/Processed. If a file cannot be read it moves to
     <folder>/Failed with a .txt note saying why. An existing tracker is never
     overwritten, because it will hold the team's progress; a new run gets a
     timestamp in its name instead.

Needs: pip install python-docx openpyxl
"""
import argparse
import datetime as dt
import math
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

try:
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("python-docx is missing. Run: pip install python-docx openpyxl")
try:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("openpyxl is missing. Run: pip install python-docx openpyxl")

# --------------------------------------------------------------------------- reading the plan


def clean(text):
    text = re.sub(r"\s+", " ", text.replace("\t", " ")).strip()
    return text


def sentences_without_see(text):
    """Drop 'See <link text>.' sentences, which mean little once the link is gone."""
    text = re.sub(r"(^|\s)See [^.]*\.", " ", text)
    return clean(text)


def cell_paragraphs(cell):
    out = []
    for p in cell.paragraphs:
        t = clean(p.text)
        if not t or t.lower().startswith("note "):
            continue
        out.append(sentences_without_see(t))
    return [t for t in out if t]


def read_plan(path):
    if not zipfile.is_zipfile(str(path)):
        raise ValueError("This is not a valid .docx file. It may be an old .doc file (save it as .docx) or a damaged download.")
    d = docx.Document(str(path))
    title = ""
    for p in d.paragraphs:
        if clean(p.text):
            title = clean(p.text)
            break

    caption = ""
    heading = ""
    prep_items = []          # (text, note)
    req_items = []
    phases = []              # {"name", "items": [(type, text, note)]}

    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(child, d)
            style = p.style.name if p.style is not None else ""
            text = clean(p.text)
            if not text:
                continue
            if style.lower().startswith("heading"):
                heading = text
            elif style in ("Table Title",):
                caption = re.sub(r"^Table\s*[\d-]*\.?\s*", "", text)
            continue
        if tag != "tbl":
            continue

        t = Table(child, d)
        if not t.rows:
            continue
        head = [clean(c.text) for c in t.rows[0].cells]
        nonblank = [h for h in head if h]

        # Execution plan tables
        if len(head) >= 4 and head[0].lower() == "action" and head[1].lower().startswith("impact"):
            name = re.sub(r"\s*execution plan\s*$", "", caption, flags=re.I).strip() or caption or "Execution"
            items = []
            for row in t.rows[1:]:
                cells = [clean(c.text) for c in row.cells]
                if not cells[0]:
                    continue
                action = re.sub(r"^\d+\.\s*", "", cells[0])
                required = cells[2] if len(cells) > 2 else ""
                downtime = cells[3] if len(cells) > 3 else ""
                impact = cells[1] if len(cells) > 1 else ""
                kind = "Checkpoint" if action.lower().startswith("validate") else "Task"
                if required.lower().startswith("no"):
                    action = re.sub(r"\.?\s*$", " (optional)", action, count=1)
                note = f"{caption}. Impact: {impact} VM downtime: {downtime}".strip()
                items.append((kind, action, note))
            if items:
                phases.append({"name": name, "items": items})
            continue

        # Preparation checklist: blank first column, "Description" second
        if nonblank == ["Description"] and "prepar" in heading.lower():
            for row in t.rows[1:]:
                txt = clean(row.cells[-1].text)
                if txt:
                    prep_items.append(("Pre-check", sentences_without_see(txt), "Convergence Preparation checklist"))
            continue

        # Requirements tables: "Item" + "...Requirement(s)"
        if len(head) >= 2 and head[0].lower() == "item" and "requirement" in head[1].lower():
            for row in t.rows[1:]:
                label = clean(row.cells[0].text)
                paras = cell_paragraphs(row.cells[1])
                if not label or not paras:
                    continue
                src = f"General requirements: {label}" if "general" in heading.lower() else f"Requirements table: {label}"
                if label.lower().startswith("minimum component"):
                    req_items.append(("Pre-check", f"{label}: " + "; ".join(paras), src))
                else:
                    for para in paras:
                        req_items.append(("Pre-check", f"{label}: {para}", src))

    prep = prep_items + req_items
    if prep:
        prep.append(("Checkpoint", "Preparation complete: go or no-go decision recorded", "Added by the script"))
        phases.insert(0, {"name": "Preparation", "items": prep})
    if not any(p["name"] != "Preparation" for p in phases):
        raise ValueError("No execution plan table found (a table headed Action / Impact to Environment / Required / VM Downtime).")
    return title or Path(path).stem, phases


# --------------------------------------------------------------------------- building the workbook

ARIAL = "Arial"
NAVY = "1F3864"
INPUT = "FFF2CC"
STEP_A = "D9E2F3"
STEP_B = "FFE699"
GREY = "BFBFBF"
thin = Side(style="thin", color=GREY)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
STATUSES = "Not started,In progress,Complete,Blocked,Rolled back"
RAGS = "Green,Amber,Red"


def f(size=10, bold=False, italic=False, color=None, name=ARIAL):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def fill(rgb):
    return PatternFill("solid", fgColor=rgb)


def est_height(text, chars_per_line):
    lines = max(1, sum(math.ceil(max(1, len(part)) / chars_per_line) for part in str(text).split("\n")))
    return 15 if lines == 1 else 11.5 * lines + 3.5


def add_rag_rules(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Green"'], fill=fill("00B050"), font=Font(name=ARIAL, color="FFFFFF")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Amber"'], fill=fill("FFC000"), font=Font(name=ARIAL, color="000000")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Red"'], fill=fill("FF0000"), font=Font(name=ARIAL, color="FFFFFF")))


def build_workbook(title, phases, source_name, out_path):
    tracker = "Converge Tracker" if "converge" in title.lower() else ("Upgrade Tracker" if "upgrade" in title.lower() else "Tracker")
    T = f"'{tracker}'"
    FIRST = 12
    n_rows = sum(1 + len(p["items"]) for p in phases)
    LAST = FIRST + n_rows + 24          # leaves 24 spare rows for items added during the work
    ISS_LAST = 204

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    trk = wb.create_sheet(tracker)
    iss = wb.create_sheet("Issues Log")

    # ---- Tracker --------------------------------------------------------
    ws = trk
    ws["A1"] = f"{title}: Execution Tracker"
    ws["A1"].font = f(14, True)
    items_total = sum(len(p["items"]) for p in phases)
    ws["A2"] = (f"Source: {source_name}, generated {dt.date.today():%d/%m/%Y}. {len(phases)} phases, {items_total} items. "
                "Items marked (optional) have Required = No in the plan.")
    ws["A2"].font = f(10, italic=True, color="595959")
    ws["G3"] = "Progress summary"
    ws["G3"].font = f(10, True)

    labels = ["Change reference", "Customer / environment", "Change owner", "Window start", "Window end"]
    for i, lab in enumerate(labels):
        r = 4 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(r, 1, lab).font = f(10, True)
        cell = ws.cell(r, 3)
        cell.border = BOX
        cell.fill = fill(INPUT)
        cell.font = f()
        cell.alignment = Alignment(horizontal="left")
        if lab.startswith("Window"):
            cell.number_format = "dd/mm/yyyy hh:mm"

    summary = [
        (4, "Green", f'=COUNTIF($G${FIRST}:$G${LAST},"Green")', "00B050", "FFFFFF"),
        (5, "Amber", f'=COUNTIF($G${FIRST}:$G${LAST},"Amber")', "FFC000", None),
        (6, "Red", f'=COUNTIF($G${FIRST}:$G${LAST},"Red")', "FF0000", "FFFFFF"),
        (7, "Items complete (excl. step rows)",
         f'=COUNTIF($F${FIRST}:$F${LAST},"Complete")-COUNTIFS($B${FIRST}:$B${LAST},"Step",$F${FIRST}:$F${LAST},"Complete")', None, None),
        (8, "Items total (excl. step rows)", f'=COUNTA($C${FIRST}:$C${LAST})-COUNTIF($B${FIRST}:$B${LAST},"Step")', None, None),
        (9, "Percent complete", "=IF(J8=0,0,J7/J8)", None, None),
        (10, "Open issues", f"=COUNTIF('Issues Log'!$H$5:$H${ISS_LAST},\"Open\")", None, None),
    ]
    for r, lab, formula, bg, fg in summary:
        ws.cell(r, 7, lab)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        ws.cell(r, 7).font = f(10, bg is not None, color=fg)
        if bg:
            for c in (7, 8, 9):
                ws.cell(r, c).fill = fill(bg)
        ws.cell(r, 10, formula).font = f(10, True)
        ws.cell(r, 10).alignment = Alignment(horizontal="left")
        ws.cell(r, 10).border = BOX
    ws["J9"].number_format = "0%"

    ws.merge_cells("A10:F10")
    ws["A10"] = ("Yellow cells are for you to fill in. Pick Status and RAG from the dropdowns. Green = on plan, "
                 "Amber = issue being managed, Red = blocked or needs a decision. Log each issue in the Issues Log tab using the item Ref.")
    ws["A10"].font = f(10, italic=True, color="595959")
    ws["A10"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[10].height = 54

    heads = ["Ref", "Type", "Item", "Owner", "Planned date", "Status", "RAG", "Actual start", "Actual end", "Notes / issues observed", "Plan reference"]
    for c, h in enumerate(heads, 1):
        cell = ws.cell(11, c, h)
        cell.font = f(10, True, color="FFFFFF")
        cell.fill = fill(NAVY)
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[11].height = 24

    r = FIRST
    for pi, ph in enumerate(phases, 1):
        # step row
        vals = {1: f"S{pi}", 2: "Step", 3: f"Phase {pi}: {ph['name']}", 6: "Not started"}
        for c in range(1, 12):
            cell = ws.cell(r, c, vals.get(c))
            cell.border = BOX
            cell.font = f(10, True)
            cell.alignment = Alignment(horizontal="center" if c in (1, 2, 6) else None, vertical="center", wrap_text=True)
            cell.fill = fill(STEP_A if c in (1, 2, 3, 11) else STEP_B)
        ws.row_dimensions[r].height = 15
        r += 1
        for ii, (kind, text, note) in enumerate(ph["items"], 1):
            vals = {1: f"{pi}.{ii}", 2: kind, 3: text, 6: "Not started", 11: note}
            for c in range(1, 12):
                cell = ws.cell(r, c, vals.get(c))
                cell.border = BOX
                cell.font = f()
                cell.alignment = Alignment(horizontal="center" if c in (1, 2, 5, 6, 7, 8, 9) else None, vertical="center", wrap_text=True)
                if c in (4, 5, 6, 7, 8, 9, 10):
                    cell.fill = fill(INPUT)
            ws.cell(r, 5).number_format = "dd/mm/yyyy"
            ws.cell(r, 8).number_format = "dd/mm/yyyy hh:mm"
            ws.cell(r, 9).number_format = "dd/mm/yyyy hh:mm"
            ws.row_dimensions[r].height = max(est_height(text, 75), est_height(note, 36))
            r += 1
    # spare rows for items added on the day
    for rr in range(r, LAST + 1):
        for c in range(1, 12):
            cell = ws.cell(rr, c)
            cell.border = BOX
            cell.font = f()
            cell.alignment = Alignment(horizontal="center" if c in (1, 2, 5, 6, 7, 8, 9) else None, vertical="center", wrap_text=True)
            if c in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11):
                cell.fill = fill(INPUT)
        ws.cell(rr, 5).number_format = "dd/mm/yyyy"
        ws.cell(rr, 8).number_format = "dd/mm/yyyy hh:mm"
        ws.cell(rr, 9).number_format = "dd/mm/yyyy hh:mm"

    for col, w in zip("ABCDEFGHIJK", [7, 16, 70, 14, 13, 14, 10, 15, 15, 55, 36]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"D{FIRST}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "11:11"

    rng_status = f"F{FIRST}:F{LAST}"
    dv = DataValidation(type="list", formula1=f'"{STATUSES}"', allow_blank=True)
    dv.add(rng_status)
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="list", formula1=f'"{RAGS}"', allow_blank=True)
    dv2.add(f"G{FIRST}:G{LAST}")
    ws.add_data_validation(dv2)
    for val, col in (("Complete", "00703C"), ("Blocked", "C00000"), ("Rolled back", "C00000")):
        ws.conditional_formatting.add(rng_status, CellIsRule(operator="equal", formula=[f'"{val}"'], font=Font(name=ARIAL, bold=True, color=col)))
    add_rag_rules(ws, f"G{FIRST}:G{LAST}")

    # ---- Dashboard ------------------------------------------------------
    ws = dash
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{title}: Status at a glance"
    ws["A1"].font = f(16, True)
    ws.row_dimensions[1].height = 20.65
    ws["A2"] = f"Updates itself from the {tracker} and Issues Log tabs. Nothing to type on this page."
    ws["A2"].font = f(10, italic=True, color="595959")

    tiles = [("A", "B", "Overall complete", f"={T}!J9", "0%"),
             ("C", "D", "Items done", f"={T}!J7", None),
             ("E", "F", "Items total", f"={T}!J8", None),
             ("G", "H", "Open issues", f"={T}!J10", None),
             ("I", "J", "Red items", f"={T}!J6", None)]
    for a, b, lab, formula, nf in tiles:
        ws.merge_cells(f"{a}4:{b}4")
        ws.merge_cells(f"{a}5:{b}5")
        ws[f"{a}4"] = lab
        for col in (a, b):
            ws[f"{col}4"].fill = fill(NAVY)
            ws[f"{col}4"].border = BOX
            ws[f"{col}5"].border = BOX
        ws[f"{a}4"].font = f(9, True, color="FFFFFF")
        ws[f"{a}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{a}5"] = formula
        ws[f"{a}5"].font = f(20, True)
        ws[f"{a}5"].alignment = Alignment(horizontal="center", vertical="center")
        if nf:
            ws[f"{a}5"].number_format = nf
    ws.row_dimensions[5].height = 33.75
    ws.conditional_formatting.add("G5", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=ARIAL, bold=True, size=20, color="C55A11")))
    ws.conditional_formatting.add("I5", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=ARIAL, bold=True, size=20, color="C00000")))

    for c, h in enumerate(["#", "Phase", "Status", "RAG", "Done", "Total", "% done", "Progress", "Remaining"], 1):
        cell = ws.cell(8, c, h)
        cell.font = f(10, True, color="FFFFFF")
        cell.fill = fill(NAVY)
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[8].height = 19.5

    A = f"{T}!$A${FIRST}:$A${LAST}"
    Fcol = f"{T}!$F${FIRST}:$F${LAST}"
    Gcol = f"{T}!$G${FIRST}:$G${LAST}"
    Bcol = f"{T}!$B${FIRST}:$B${LAST}"
    for i, ph in enumerate(phases, 1):
        r = 8 + i
        ws.row_dimensions[r].height = 19.5
        ws.cell(r, 1, i)
        ws.cell(r, 2, ph["name"])
        ws.cell(r, 3, (f'=IF(COUNTIFS({A},$A{r}&".*",{Fcol},"Blocked")+COUNTIFS({A},$A{r}&".*",{Fcol},"Rolled back")>0,"Blocked",'
                       f'IF(E{r}=F{r},"Complete",IF(OR(E{r}>0,COUNTIFS({A},$A{r}&".*",{Fcol},"In progress")>0),"In progress","Not started")))'))
        ws.cell(r, 4, (f'=IF(OR(C{r}="Blocked",COUNTIFS({A},$A{r}&".*",{Gcol},"Red")>0),"Red",IF(COUNTIFS({A},$A{r}&".*",{Gcol},"Amber")>0,"Amber",'
                       f'IF(C{r}="Not started","","Green")))'))
        ws.cell(r, 5, f'=COUNTIFS({A},$A{r}&".*",{Fcol},"Complete")')
        ws.cell(r, 6, f'=COUNTIFS({A},$A{r}&".*",{Bcol},"<>Step")')
        ws.cell(r, 7, f"=IF(F{r}=0,0,E{r}/F{r})")
        ws.cell(r, 8, f'=REPT("█",ROUND(G{r}*20,0))&REPT("░",20-ROUND(G{r}*20,0))')
        ws.cell(r, 9, f"=F{r}-E{r}")
        for c in range(1, 10):
            cell = ws.cell(r, c)
            cell.border = BOX
            cell.font = f()
            cell.alignment = Alignment(horizontal="left" if c == 2 else "center", vertical="center")
        ws.cell(r, 7).number_format = "0%"
        ws.cell(r, 8).font = Font(name="Consolas", size=10, color="2E75B6")
        ws.cell(r, 8).alignment = Alignment(horizontal="left", vertical="center")
    last = 8 + len(phases)
    ws.cell(last + 1, 2, "RAG per phase: Red if any item is Red or Blocked, Amber if any is Amber, otherwise Green once the phase has started.")
    ws.cell(last + 1, 2).font = f(9, italic=True, color="595959")
    for val, kw in (("Complete", dict(bold=True, color="00703C")), ("Blocked", dict(bold=True, color="C00000")), ("Not started", dict(color="7F7F7F"))):
        ws.conditional_formatting.add(f"C9:C{last}", CellIsRule(operator="equal", formula=[f'"{val}"'], font=Font(name=ARIAL, **kw)))
    ws.conditional_formatting.add(f"C9:C{last}", CellIsRule(operator="equal", formula=['"In progress"'], fill=fill("DDEBF7"), font=Font(name=ARIAL, color="1F4E79")))
    add_rag_rules(ws, f"D9:D{last}")
    for col, w in zip("ABCDEFGHIJ", [5, 42, 14, 10, 8, 8, 9, 26, 11, 3]):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- Issues Log -----------------------------------------------------
    ws = iss
    ws["A1"] = "Issues Log"
    ws["A1"].font = f(14, True)
    ws["A2"] = ("One row per issue observed during the work. Ref points back to the " + tracker +
                " item. The grey italic first row is an example; overwrite or delete it.")
    ws["A2"].font = f(10, italic=True, color="595959")
    ih = ["No.", "Date / time", "Ref", "Issue observed", "Severity", "Action taken / workaround", "Owner", "Status", "Resolved date / time", "Vendor case / KB"]
    for c, h in enumerate(ih, 1):
        cell = ws.cell(4, c, h)
        cell.font = f(10, True, color="FFFFFF")
        cell.fill = fill(NAVY)
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 24
    first_ref = None
    for ph_i, ph in enumerate(phases, 1):
        if ph["items"] and first_ref is None and ph["name"] != "Preparation":
            first_ref = f"{ph_i}.1"
    example = [1, dt.datetime.combine(dt.date.today(), dt.time(9, 30)), first_ref or "1.1",
               "Example: a precheck reports a stale DNS record for one host.", "Amber",
               "Corrected the record and re-ran the precheck. No impact on the window.", "Engineer", "Resolved",
               dt.datetime.combine(dt.date.today(), dt.time(10, 15)), "n/a"]
    for r in range(5, ISS_LAST + 1):
        for c in range(1, 11):
            cell = ws.cell(r, c)
            cell.border = BOX
            cell.fill = fill(INPUT)
            cell.font = f(10, italic=(r == 5), color="7F7F7F" if r == 5 else None)
            cell.alignment = Alignment(horizontal="center" if c in (1, 2, 3, 5, 7, 8, 9) else None, vertical="center", wrap_text=True)
        ws.cell(r, 2).number_format = "dd/mm/yyyy hh:mm"
        ws.cell(r, 9).number_format = "dd/mm/yyyy hh:mm"
    for c, v in enumerate(example, 1):
        ws.cell(5, c, v)
    for col, w in zip("ABCDEFGHIJ", [6, 17, 8, 60, 11, 50, 14, 12, 17, 20]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    dv = DataValidation(type="list", formula1=f'"{RAGS}"', allow_blank=True)
    dv.add(f"E5:E{ISS_LAST}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"Open,Resolved"', allow_blank=True)
    dv.add(f"H5:H{ISS_LAST}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=f"={T}!$A${FIRST}:$A${LAST}", allow_blank=True)
    dv.add(f"C5:C{ISS_LAST}")
    ws.add_data_validation(dv)
    add_rag_rules(ws, f"E5:E{ISS_LAST}")

    wb.save(out_path)


# --------------------------------------------------------------------------- running it


def unique_path(path):
    path = Path(path)
    if not path.exists():
        return path
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    return path.with_name(f"{path.stem} {stamp}{path.suffix}")


def convert(docx_path, out_path=None):
    docx_path = Path(docx_path)
    title, phases = read_plan(docx_path)
    out = unique_path(out_path or docx_path.with_name(docx_path.stem + " - Tracker.xlsx"))
    build_workbook(title, phases, docx_path.name, out)
    return out, phases


def ready(path):
    """True if the file looks finished being written and is not open in Word."""
    if path.name.startswith("~$"):
        return False
    if (path.parent / f"~${path.name[2:]}").exists() or (path.parent / f"~${path.name}").exists():
        return False
    if time.time() - path.stat().st_mtime < 3:
        return False
    if zipfile.is_zipfile(path):
        return True
    # Not a valid Word file. Give a copy that is still downloading a moment, then fail it properly.
    return time.time() - path.stat().st_mtime > 30


def process_folder(folder, out_dir, done_dir, failed_dir):
    n = 0
    for p in sorted(folder.glob("*.docx")):
        if not ready(p):
            continue
        try:
            out, phases = convert(p, out_dir / (p.stem + " - Tracker.xlsx"))
            shutil.move(str(p), str(unique_path(done_dir / p.name)))
            items = sum(len(x["items"]) for x in phases)
            print(f"{dt.datetime.now():%H:%M:%S}  {p.name} -> {out.name}  ({len(phases)} phases, {items} items)", flush=True)
            n += 1
        except Exception as exc:  # keep the watcher alive whatever is dropped in
            dest = unique_path(failed_dir / p.name)
            shutil.move(str(p), str(dest))
            dest.with_suffix(".txt").write_text(f"Could not build a tracker from {p.name}:\n{exc}\n", encoding="utf8")
            print(f"{dt.datetime.now():%H:%M:%S}  {p.name} FAILED: {exc}", flush=True)
    return n


def main():
    ap = argparse.ArgumentParser(description="Build an Excel execution tracker from a VCF Implementation Plan (.docx).")
    ap.add_argument("docx", nargs="?", help="Word plan to convert")
    ap.add_argument("-o", "--output", help="Output .xlsx (single file mode)")
    ap.add_argument("--watch", metavar="FOLDER", help="Drop folder to watch for .docx files")
    ap.add_argument("--once", action="store_true", help="With --watch: process what is there, then stop")
    ap.add_argument("--interval", type=int, default=5, help="Seconds between checks in watch mode (default 5)")
    a = ap.parse_args()

    if a.watch:
        folder = Path(a.watch).expanduser().resolve()
        out_dir, done_dir, failed_dir = folder / "Trackers", folder / "Processed", folder / "Failed"
        for d in (folder, out_dir, done_dir, failed_dir):
            d.mkdir(parents=True, exist_ok=True)
        print(f"Watching {folder}  (Ctrl+C to stop)", flush=True)
        if a.once:
            process_folder(folder, out_dir, done_dir, failed_dir)
            return
        try:
            while True:
                process_folder(folder, out_dir, done_dir, failed_dir)
                time.sleep(a.interval)
        except KeyboardInterrupt:
            print("Stopped.")
        return

    if not a.docx:
        ap.error("give a .docx file, or use --watch FOLDER")
    out, phases = convert(a.docx, a.output)
    print(f"Built {out}")
    for i, ph in enumerate(phases, 1):
        print(f"  Phase {i}: {ph['name']} ({len(ph['items'])} items)")


if __name__ == "__main__":
    main()
