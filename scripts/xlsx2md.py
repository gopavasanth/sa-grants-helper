"""Convert proposal budget .xlsx files into Markdown reference files.

Usage: python xlsx2md.py <budgets_dir> <proposals_dir>
Writes <budgets_dir>/<slug>.md for every <budgets_dir>/raw/<slug>.xlsx.
"""
import datetime
import glob
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as L

budgets_dir, proposals_dir = sys.argv[1], sys.argv[2]
raw_dir = os.path.join(budgets_dir, "raw")
TODAY = "2026-11-22"

CURRENCY = re.compile(r"^[A-Z]{3}$")
# Drop-down helper lists that the Rapid Fund template keeps beside the budget.
EXPENSE_TYPES = {
    "logistics and space", "administration and fees", "equipment and tools",
    "merchendize", "merchandise", "prizes and awards", "staff related expenses",
    "communications and materials", "food and beverages", "travel and accommodation",
    "internet and services", "technical support", "other", "accessibility",
    "safety and security", "others",
}


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        v = v.isoformat()[:10] if isinstance(v, datetime.datetime) and v.time() == datetime.time() else v.isoformat()
    s = str(v).strip()
    s = re.sub(r"\s*\n\s*", " ", s)
    return s.replace("|", "\\|")


def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def is_helper_list_column(vals):
    vals = [str(v).strip() for v in vals if not is_empty(v)]
    if len(vals) < 4:
        return False
    cur = sum(1 for v in vals if CURRENCY.match(v))
    exp = sum(1 for v in vals if v.lower() in EXPENSE_TYPES)
    return cur / len(vals) > 0.9 or exp / len(vals) > 0.9


def proposal_meta(slug):
    path = os.path.join(proposals_dir, slug + ".md")
    text = open(path, encoding="utf-8").read()
    title = text.splitlines()[0].lstrip("# ").strip()
    src = re.search(r"^Source page: (\S+)", text, re.M).group(1)
    links = sorted(set(re.findall(r"https://docs\.google\.com/spreadsheets/[^\s)\]|]+", text)))
    return title, src, links


report = []
for xlsx in sorted(glob.glob(os.path.join(raw_dir, "*.xlsx"))):
    slug = os.path.basename(xlsx)[:-5]
    title, src, links = proposal_meta(slug)
    wb_v = openpyxl.load_workbook(xlsx, data_only=True)
    wb_f = openpyxl.load_workbook(xlsx, data_only=False)

    out = []
    out.append(f"# Budget: {title}\n")
    out.append(f"This file belongs ONLY to the proposal \"{title}\".")
    out.append(f"Proposal file: references/proposals/{slug}.md")
    out.append(f"Proposal page: {src}")
    out.append("Budget spreadsheet link as given in the proposal: " + ", ".join(links))
    out.append(f"Original download kept at: references/budgets/raw/{slug}.xlsx")
    out.append(f"Downloaded and converted: {TODAY}\n")
    out.append(
        "How this file was made: every non-empty cell of every sheet was copied from the .xlsx. "
        "Values are the stored cell values, so some numbers show more decimals than the spreadsheet displays. "
        "Removed only: fully empty rows, rows where every cell is 0 (unused formula rows), and the template's "
        "drop-down helper lists (currency codes, expense type lists). Nothing was added, totalled or converted. "
        "\"Row\" is the row number in the spreadsheet. Separate tables come from separate blocks of columns "
        "in the same sheet; a value in one table is not related to the row with the same number in another table "
        "unless the spreadsheet says so.\n"
    )
    out.append("---\n")

    missing_cache = 0
    for ws in wb_v.worksheets:
        wf = wb_f[ws.title]
        out.append(f"## Sheet: {ws.title}\n")
        if ws.sheet_state != "visible":
            out.append(f"(This sheet is {ws.sheet_state} in the spreadsheet.)\n")
        maxr, maxc = ws.max_row, ws.max_column
        grid = {}
        for r in range(1, maxr + 1):
            for c in range(1, maxc + 1):
                v = ws.cell(r, c).value
                if is_empty(v) and isinstance(wf.cell(r, c).value, str) and wf.cell(r, c).value.startswith("="):
                    missing_cache += 1
                if not is_empty(v):
                    grid[(r, c)] = v
        cols = sorted({c for _, c in grid})
        cols = [c for c in cols if not is_helper_list_column([grid[k] for k in grid if k[1] == c])]
        # rows that only contain zeros are unused formula rows
        def row_keep(r, cs):
            vals = [grid[(r, c)] for c in cs if (r, c) in grid]
            return vals and not all(v == 0 for v in vals)
        # split columns into blocks separated by empty columns
        blocks, cur = [], []
        for c in cols:
            if cur and c != cur[-1] + 1:
                blocks.append(cur)
                cur = []
            cur.append(c)
        if cur:
            blocks.append(cur)
        blocks = [b for b in blocks if any(row_keep(r, b) for r in range(1, maxr + 1))]
        if not blocks:
            out.append("(This sheet is empty.)\n")
            continue
        for b in blocks:
            rows = [r for r in range(1, maxr + 1) if row_keep(r, b)]
            if len(blocks) > 1:
                out.append(f"### Columns {L(b[0])}–{L(b[-1])}\n")
            out.append("| Row | " + " | ".join(L(c) for c in b) + " |")
            out.append("|---|" + "---|" * len(b))
            for r in rows:
                out.append(f"| {r} | " + " | ".join(fmt(grid.get((r, c))) for c in b) + " |")
            out.append("")
    dst = os.path.join(budgets_dir, slug + ".md")
    open(dst, "w", encoding="utf-8").write("\n".join(out).rstrip() + "\n")
    report.append((slug, len(wb_v.worksheets), missing_cache))

for r in report:
    print(r)
