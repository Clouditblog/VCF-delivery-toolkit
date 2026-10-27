"""Prototype of the raw-OOXML writer that the PowerShell script ports. No openpyxl used for writing."""
import sys, zipfile, datetime as dt, re
sys.path.insert(0, "/home/claude/xl")
from plan_to_tracker import read_plan

NAVY, INPUT, STEP_A, STEP_B, GREY = "1F3864", "FFF2CC", "D9E2F3", "FFE699", "BFBFBF"


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def colname(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Styles:
    def __init__(self):
        self.fonts, self.fills, self.xfs, self.dxfs = [], ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], [], []
        self.fkey, self.flkey, self.xkey = {}, {}, {}
        self.font(10, False, False, None)           # font 0 default
        self.xf(0, 0, 0, 0, "", "", False)         # xf 0 default

    def font(self, sz, b, i, color, name="Arial"):
        key = f"{sz}|{b}|{i}|{color}|{name}"
        if key not in self.fkey:
            x = "<font>" + ("<b/>" if b else "") + ("<i/>" if i else "") + f'<sz val="{sz}"/>' + (f'<color rgb="FF{color}"/>' if color else "") + f'<name val="{name}"/></font>'
            self.fonts.append(x)
            self.fkey[key] = len(self.fonts) - 1
        return self.fkey[key]

    def fill(self, rgb):
        if rgb is None:
            return 0
        if rgb not in self.flkey:
            self.fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/><bgColor indexed="64"/></patternFill></fill>')
            self.flkey[rgb] = len(self.fills) - 1
        return self.flkey[rgb]

    def xf(self, font, fill, border, numfmt, h, v, wrap):
        key = f"{font}|{fill}|{border}|{numfmt}|{h}|{v}|{wrap}"
        if key not in self.xkey:
            al = ""
            if h or v or wrap:
                al = "<alignment" + (f' horizontal="{h}"' if h else "") + (f' vertical="{v}"' if v else "") + (' wrapText="1"' if wrap else "") + "/>"
            x = (f'<xf numFmtId="{numfmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"'
                 + (' applyNumberFormat="1"' if numfmt else "") + ' applyFont="1"' + (' applyFill="1"' if fill else "")
                 + (' applyBorder="1"' if border else "") + (' applyAlignment="1">' + al + "</xf>" if al else "/>"))
            self.xfs.append(x)
            self.xkey[key] = len(self.xfs) - 1
        return self.xkey[key]

    def style(self, sz=10, b=False, i=False, color=None, fill=None, border=False, nf=0, h="", v="", wrap=False, name="Arial"):
        return self.xf(self.font(sz, b, i, color, name), self.fill(fill), 1 if border else 0, nf, h, v, wrap)

    def dxf(self, color=None, bold=False, fill=None, size=None):
        x = "<dxf>"
        if color or bold or size:
            x += "<font>" + ("<b/>" if bold else "") + (f'<sz val="{size}"/>' if size else "") + (f'<color rgb="FF{color}"/>' if color else "") + "</font>"
        if fill:
            x += f'<fill><patternFill><bgColor rgb="FF{fill}"/></patternFill></fill>'
        x += "</dxf>"
        self.dxfs.append(x)
        return len(self.dxfs) - 1

    def xml(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/></numFmts>'
                f'<fonts count="{len(self.fonts)}">{"".join(self.fonts)}</fonts>'
                f'<fills count="{len(self.fills)}">{"".join(self.fills)}</fills>'
                '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
                f'<border><left style="thin"><color rgb="FF{GREY}"/></left><right style="thin"><color rgb="FF{GREY}"/></right>'
                f'<top style="thin"><color rgb="FF{GREY}"/></top><bottom style="thin"><color rgb="FF{GREY}"/></bottom><diagonal/></border></borders>'
                '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                f'<cellXfs count="{len(self.xfs)}">{"".join(self.xfs)}</cellXfs>'
                '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                f'<dxfs count="{len(self.dxfs)}">{"".join(self.dxfs)}</dxfs></styleSheet>')


class Sheet:
    def __init__(self, name):
        self.name, self.cells, self.rowh, self.merges, self.cols, self.cf, self.dv = name, {}, {}, [], [], [], []
        self.freeze, self.grid, self.prio = None, True, 0

    def put(self, r, c, val=None, s=0, kind="s", cached=None, ctype="n"):
        ref = f"{colname(c)}{r}"
        if val is None and kind != "f":
            x = f'<c r="{ref}" s="{s}"/>'
        elif kind == "s":
            x = f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(val)}</t></is></c>'
        elif kind == "n":
            x = f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
        else:  # formula, with a cached value so Protected View is not blank
            t = ' t="str"' if ctype == "s" else ""
            v = f"<v>{esc(cached)}</v>" if cached is not None else ""
            x = f'<c r="{ref}" s="{s}"{t}><f>{esc(val)}</f>{v}</c>'
        self.cells[(r, c)] = x

    def merge(self, ref):
        self.merges.append(ref)

    def add_cf(self, sqref, rules):
        self.cf.append((sqref, rules))

    def xml(self):
        rows = {}
        for (r, c), x in self.cells.items():
            rows.setdefault(r, []).append((c, x))
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>']
        view = f'<sheetView workbookViewId="0"{"" if self.grid else " showGridLines=\"0\""}>'
        if self.freeze:
            col, row = self.freeze
            tl = f"{colname(col + 1)}{row + 1}"
            if col and row:
                view += f'<pane xSplit="{col}" ySplit="{row}" topLeftCell="{tl}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/>'
            else:
                view += f'<pane ySplit="{row}" topLeftCell="{tl}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>'
        view += "</sheetView>"
        out.append(f"<sheetViews>{view}</sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>")
        out.append("<cols>" + "".join(f'<col min="{a}" max="{b}" width="{w}" customWidth="1"/>' for a, b, w in self.cols) + "</cols>")
        out.append("<sheetData>")
        for r in sorted(rows):
            h = f' ht="{self.rowh[r]}" customHeight="1"' if r in self.rowh else ""
            out.append(f'<row r="{r}"{h}>' + "".join(x for _, x in sorted(rows[r])) + "</row>")
        out.append("</sheetData>")
        if self.merges:
            out.append(f'<mergeCells count="{len(self.merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in self.merges) + "</mergeCells>")
        p = 0
        for sq, rules in self.cf:
            out.append(f'<conditionalFormatting sqref="{sq}">')
            for op, val, dxf in rules:
                p += 1
                out.append(f'<cfRule type="cellIs" dxfId="{dxf}" priority="{p}" operator="{op}"><formula>{esc(val)}</formula></cfRule>')
            out.append("</conditionalFormatting>")
        if self.dv:
            out.append(f'<dataValidations count="{len(self.dv)}">' + "".join(
                f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="{sq}"><formula1>{esc(f)}</formula1></dataValidation>' for sq, f in self.dv) + "</dataValidations>")
        out.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>')
        return "".join(out)


def est_height(text, cpl):
    lines = max(1, sum(-(-max(1, len(p)) // cpl) for p in str(text).split("\n")))
    return 15 if lines == 1 else round(11.5 * lines + 3.5, 2)


def build(title, phases, source_name, out_path):
    st = Styles()
    tracker = "Converge Tracker" if "converge" in title.lower() else ("Upgrade Tracker" if "upgrade" in title.lower() else "Tracker")
    T = f"'{tracker}'"
    FIRST = 12
    n_rows = sum(1 + len(p["items"]) for p in phases)
    LAST = FIRST + n_rows + 24
    ISS_LAST = 204
    items_total = sum(len(p["items"]) for p in phases)

    # dxfs
    d_green, d_amber, d_red = st.dxf("FFFFFF", False, "00B050"), st.dxf("000000", False, "FFC000"), st.dxf("FFFFFF", False, "FF0000")
    d_complete, d_blocked = st.dxf("00703C", True), st.dxf("C00000", True)
    d_notstarted, d_inprog = st.dxf("7F7F7F"), st.dxf("1F4E79", False, "DDEBF7")
    d_amber_big, d_red_big = st.dxf("C55A11", True, None, 20), st.dxf("C00000", True, None, 20)
    rag = [("equal", '"Green"', d_green), ("equal", '"Amber"', d_amber), ("equal", '"Red"', d_red)]

    dash, trk, iss = Sheet("Dashboard"), Sheet(tracker), Sheet("Issues Log")
    dash.grid = False

    # ------------------ tracker
    ws = trk
    ws.put(1, 1, f"{title}: Execution Tracker", st.style(14, True))
    ws.put(2, 1, f"Source: {source_name}, generated {dt.date.today():%d/%m/%Y}. {len(phases)} phases, {items_total} items. Items marked (optional) have Required = No in the plan.", st.style(10, i=True, color="595959"))
    ws.put(3, 7, "Progress summary", st.style(10, True))
    for i, lab in enumerate(["Change reference", "Customer / environment", "Change owner", "Window start", "Window end"]):
        r = 4 + i
        ws.put(r, 1, lab, st.style(10, True))
        ws.merge(f"A{r}:B{r}")
        ws.put(r, 3, None, st.style(10, fill=INPUT, border=True, h="left", nf=165 if lab.startswith("Window") else 0))
    summary = [
        (4, "Green", f'COUNTIF($G${FIRST}:$G${LAST},"Green")', "00B050", "FFFFFF", 0),
        (5, "Amber", f'COUNTIF($G${FIRST}:$G${LAST},"Amber")', "FFC000", None, 0),
        (6, "Red", f'COUNTIF($G${FIRST}:$G${LAST},"Red")', "FF0000", "FFFFFF", 0),
        (7, "Items complete (excl. step rows)", f'COUNTIF($F${FIRST}:$F${LAST},"Complete")-COUNTIFS($B${FIRST}:$B${LAST},"Step",$F${FIRST}:$F${LAST},"Complete")', None, None, 0),
        (8, "Items total (excl. step rows)", f'COUNTA($C${FIRST}:$C${LAST})-COUNTIF($B${FIRST}:$B${LAST},"Step")', None, None, items_total),
        (9, "Percent complete", "IF(J8=0,0,J7/J8)", None, None, 0),
        (10, "Open issues", f"COUNTIF('Issues Log'!$H$5:$H${ISS_LAST},\"Open\")", None, None, 0),
    ]
    for r, lab, fml, bg, fg, cached in summary:
        ws.put(r, 7, lab, st.style(10, bg is not None, color=fg, fill=bg))
        ws.put(r, 8, None, st.style(10, fill=bg))
        ws.put(r, 9, None, st.style(10, fill=bg))
        ws.merge(f"G{r}:I{r}")
        ws.put(r, 10, fml, st.style(10, True, border=True, h="left", nf=9 if r == 9 else 0), "f", cached)
    ws.merge(f"A10:F10")
    ws.put(10, 1, "Yellow cells are for you to fill in. Pick Status and RAG from the dropdowns. Green = on plan, Amber = issue being managed, Red = blocked or needs a decision. Log each issue in the Issues Log tab using the item Ref.",
           st.style(10, i=True, color="595959", wrap=True, v="top"))
    ws.rowh[10] = 54
    for c, h in enumerate(["Ref", "Type", "Item", "Owner", "Planned date", "Status", "RAG", "Actual start", "Actual end", "Notes / issues observed", "Plan reference"], 1):
        ws.put(11, c, h, st.style(10, True, color="FFFFFF", fill=NAVY, border=True, h="center", v="center", wrap=True))
    ws.rowh[11] = 24
    ctr = (1, 2, 5, 6, 7, 8, 9)
    r = FIRST

    def item_style(c, step=False):
        nf = 164 if c == 5 else (165 if c in (8, 9) else 0)
        if step:
            return st.style(10, True, fill=STEP_A if c in (1, 2, 3, 11) else STEP_B, border=True, h="center" if c in (1, 2, 6) else "", v="center", wrap=True)
        return st.style(10, fill=INPUT if c in (4, 5, 6, 7, 8, 9, 10) else None, border=True, nf=nf, h="center" if c in ctr else "", v="center", wrap=True)

    for pi, ph in enumerate(phases, 1):
        vals = {1: f"S{pi}", 2: "Step", 3: f"Phase {pi}: {ph['name']}", 6: "Not started"}
        for c in range(1, 12):
            ws.put(r, c, vals.get(c), item_style(c, True))
        ws.rowh[r] = 15
        r += 1
        for ii, (kind, text, note) in enumerate(ph["items"], 1):
            vals = {1: f"{pi}.{ii}", 2: kind, 3: text, 6: "Not started", 11: note}
            for c in range(1, 12):
                ws.put(r, c, vals.get(c), item_style(c))
            ws.rowh[r] = max(est_height(text, 75), est_height(note, 36))
            r += 1
    for rr in range(r, LAST + 1):
        for c in range(1, 12):
            ws.put(rr, c, None, st.style(10, fill=INPUT, border=True, nf=164 if c == 5 else (165 if c in (8, 9) else 0), h="center" if c in ctr else "", v="center", wrap=True))
    ws.cols = [(1, 1, 7), (2, 2, 16), (3, 3, 70), (4, 4, 14), (5, 5, 13), (6, 6, 14), (7, 7, 10), (8, 9, 15), (10, 10, 55), (11, 11, 36)]
    ws.freeze = (3, 11)
    ws.dv += [(f"F{FIRST}:F{LAST}", '"Not started,In progress,Complete,Blocked,Rolled back"'), (f"G{FIRST}:G{LAST}", '"Green,Amber,Red"')]
    ws.add_cf(f"F{FIRST}:F{LAST}", [("equal", '"Complete"', d_complete), ("equal", '"Blocked"', d_blocked), ("equal", '"Rolled back"', d_blocked)])
    ws.add_cf(f"G{FIRST}:G{LAST}", rag)

    # ------------------ dashboard
    ws = dash
    ws.put(1, 1, f"{title}: Status at a glance", st.style(16, True))
    ws.rowh[1] = 20.65
    ws.put(2, 1, f"Updates itself from the {tracker} and Issues Log tabs. Nothing to type on this page.", st.style(10, i=True, color="595959"))
    tiles = [(1, 2, "Overall complete", f"{T}!J9", 9, 0), (3, 4, "Items done", f"{T}!J7", 0, 0), (5, 6, "Items total", f"{T}!J8", 0, items_total),
             (7, 8, "Open issues", f"{T}!J10", 0, 0), (9, 10, "Red items", f"{T}!J6", 0, 0)]
    for a, b, lab, fml, nf, cached in tiles:
        for c in (a, b):
            ws.put(4, c, lab if c == a else None, st.style(9, True, color="FFFFFF", fill=NAVY, border=True, h="center", v="center"))
            ws.put(5, c, fml if c == a else None, st.style(20, True, border=True, h="center", v="center", nf=nf), "f" if c == a else "s", cached)
        ws.merge(f"{colname(a)}4:{colname(b)}4")
        ws.merge(f"{colname(a)}5:{colname(b)}5")
    ws.rowh[5] = 33.75
    ws.add_cf("G5", [("greaterThan", "0", d_amber_big)])
    ws.add_cf("I5", [("greaterThan", "0", d_red_big)])
    for c, h in enumerate(["#", "Phase", "Status", "RAG", "Done", "Total", "% done", "Progress", "Remaining"], 1):
        ws.put(8, c, h, st.style(10, True, color="FFFFFF", fill=NAVY, border=True, h="center", v="center"))
    ws.rowh[8] = 19.5
    A, Fc, Gc, Bc = (f"{T}!$%s${FIRST}:$%s${LAST}" % (x, x) for x in "AFGB")
    for i, ph in enumerate(phases, 1):
        r = 8 + i
        ws.rowh[r] = 19.5
        n_items = len(ph["items"])
        base = st.style(10, border=True, h="center", v="center")
        ws.put(r, 1, i, base, "n")
        ws.put(r, 2, ph["name"], st.style(10, border=True, h="left", v="center"))
        ws.put(r, 3, (f'IF(COUNTIFS({A},$A{r}&".*",{Fc},"Blocked")+COUNTIFS({A},$A{r}&".*",{Fc},"Rolled back")>0,"Blocked",'
                      f'IF(E{r}=F{r},"Complete",IF(OR(E{r}>0,COUNTIFS({A},$A{r}&".*",{Fc},"In progress")>0),"In progress","Not started")))'), base, "f", "Not started", "s")
        ws.put(r, 4, (f'IF(OR(C{r}="Blocked",COUNTIFS({A},$A{r}&".*",{Gc},"Red")>0),"Red",IF(COUNTIFS({A},$A{r}&".*",{Gc},"Amber")>0,"Amber",'
                      f'IF(C{r}="Not started","","Green")))'), base, "f", "", "s")
        ws.put(r, 5, f'COUNTIFS({A},$A{r}&".*",{Fc},"Complete")', base, "f", 0)
        ws.put(r, 6, f'COUNTIFS({A},$A{r}&".*",{Bc},"<>Step")', base, "f", n_items)
        ws.put(r, 7, f"IF(F{r}=0,0,E{r}/F{r})", st.style(10, border=True, h="center", v="center", nf=9), "f", 0)
        ws.put(r, 8, f'REPT("█",ROUND(G{r}*20,0))&REPT("░",20-ROUND(G{r}*20,0))', st.style(10, color="2E75B6", border=True, h="left", v="center", name="Consolas"), "f", "░" * 20, "s")
        ws.put(r, 9, f"F{r}-E{r}", base, "f", n_items)
    last = 8 + len(phases)
    ws.put(last + 1, 2, "RAG per phase: Red if any item is Red or Blocked, Amber if any is Amber, otherwise Green once the phase has started.", st.style(9, i=True, color="595959"))
    ws.add_cf(f"C9:C{last}", [("equal", '"Complete"', d_complete), ("equal", '"In progress"', d_inprog), ("equal", '"Blocked"', d_blocked), ("equal", '"Not started"', d_notstarted)])
    ws.add_cf(f"D9:D{last}", rag)
    ws.cols = [(1, 1, 5), (2, 2, 42), (3, 3, 14), (4, 4, 10), (5, 6, 8), (7, 7, 9), (8, 8, 26), (9, 9, 11), (10, 10, 3)]

    # ------------------ issues log
    ws = iss
    ws.put(1, 1, "Issues Log", st.style(14, True))
    ws.put(2, 1, f"One row per issue observed during the work. Ref points back to the {tracker} item. The grey italic first row is an example; overwrite or delete it.", st.style(10, i=True, color="595959"))
    for c, h in enumerate(["No.", "Date / time", "Ref", "Issue observed", "Severity", "Action taken / workaround", "Owner", "Status", "Resolved date / time", "Vendor case / KB"], 1):
        ws.put(4, c, h, st.style(10, True, color="FFFFFF", fill=NAVY, border=True, h="center", v="center", wrap=True))
    ws.rowh[4] = 24
    first_ref = "1.1"
    for pi, ph in enumerate(phases, 1):
        if ph["name"] != "Preparation":
            first_ref = f"{pi}.1"
            break
    today = dt.date.today()
    serial = (today - dt.date(1899, 12, 30)).days
    example = {1: (1, "n"), 2: (serial + 9.5 / 24, "n"), 3: (first_ref, "s"), 4: ("Example: a precheck reports a stale DNS record for one host.", "s"), 5: ("Amber", "s"),
               6: ("Corrected the record and re-ran the precheck. No impact on the window.", "s"), 7: ("Engineer", "s"), 8: ("Resolved", "s"), 9: (serial + 10.25 / 24, "n"), 10: ("n/a", "s")}
    for r in range(5, ISS_LAST + 1):
        for c in range(1, 11):
            s = st.style(10, i=(r == 5), color="7F7F7F" if r == 5 else None, fill=INPUT, border=True, nf=165 if c in (2, 9) else 0,
                         h="center" if c in (1, 2, 3, 5, 7, 8, 9) else "", v="center", wrap=True)
            if r == 5:
                v, k = example[c]
                ws.put(r, c, v, s, k)
            else:
                ws.put(r, c, None, s)
    ws.cols = [(1, 1, 6), (2, 2, 17), (3, 3, 8), (4, 4, 60), (5, 5, 11), (6, 6, 50), (7, 7, 14), (8, 8, 12), (9, 9, 17), (10, 10, 20)]
    ws.freeze = (0, 4)
    ws.dv += [(f"E5:E{ISS_LAST}", '"Green,Amber,Red"'), (f"H5:H{ISS_LAST}", '"Open,Resolved"'), (f"C5:C{ISS_LAST}", f"{T}!$A${FIRST}:$A${LAST}")]
    ws.add_cf(f"E5:E{ISS_LAST}", rag)

    sheets = [dash, trk, iss]
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in (1, 2, 3))
          + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wbx = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<bookViews><workbookView activeTab="0"/></bookViews><sheets>'
           + "".join(f'<sheet name="{esc(s.name)}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(sheets, 1))
           + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in (1, 2, 3))
             + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wbx)
        z.writestr("xl/_rels/workbook.xml.rels", wrels)
        z.writestr("xl/styles.xml", st.xml())
        for i, s in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", s.xml())


if __name__ == "__main__":
    title, phases = read_plan(sys.argv[1])
    build(title, phases, sys.argv[1].split("/")[-1], sys.argv[2])
    print("built", sys.argv[2])
