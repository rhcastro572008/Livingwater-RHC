"""Build the daily Ice Monitoring Report.

Layout:
  Row 1  "Ice Monitoring Report" (bold 16, merged A1:C1)
  Row 2  report date as text (merged A2:C2)
  Row 4  header Rider | Customer | Qty (navy fill, white bold)
  Then one block per rider (config order): client rows, then "<Rider> Subtotal" (=SUM).
  Riders with no entries get a "No sales" row with Qty 0.
  Last row: TOTAL = sum of the subtotal cells.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .tracker import Sale, fmt_long


@dataclass
class RiderBlock:
    rider: str
    lines: list[tuple[str, float]] = field(default_factory=list)

    @property
    def subtotal(self) -> float:
        return sum(q for _, q in self.lines)


@dataclass
class DailyReport:
    report_date: date
    blocks: list[RiderBlock]

    @property
    def total(self) -> float:
        return sum(b.subtotal for b in self.blocks)

    @property
    def title_date(self) -> str:
        return fmt_long(self.report_date)

    @property
    def filename(self) -> str:
        return f"Ice Monitoring Report - {self.title_date}.xlsx"


def build_report(report_date: date, sales: list[Sale], riders: list[str]) -> DailyReport:
    order = list(riders) + sorted({s.rider for s in sales} - set(riders))
    blocks = []
    for rider in order:
        block = RiderBlock(rider)
        merged: dict[str, float] = {}
        for s in sales:
            if s.rider == rider and s.date == report_date:
                merged[s.client] = merged.get(s.client, 0) + s.count
        block.lines = list(merged.items())
        blocks.append(block)
    return DailyReport(report_date, blocks)


def report_workbook(report: DailyReport) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    navy = PatternFill("solid", fgColor="1F3864")
    light_blue = PatternFill("solid", fgColor="DDEBF7")
    yellow = PatternFill("solid", fgColor="FFFF00")
    grey = Side(style="thin", color="BFBFBF")
    border = Border(left=grey, right=grey, top=grey, bottom=grey)
    center = Alignment(horizontal="center")

    ws["A1"] = "Ice Monitoring Report"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A1"].alignment = center
    ws.merge_cells("A1:C1")
    ws["A2"] = report.title_date  # stored as text, not a date
    ws["A2"].data_type = "s"
    ws["A2"].alignment = center
    ws.merge_cells("A2:C2")

    for col, text in enumerate(["Rider", "Customer", "Qty"], start=1):
        c = ws.cell(row=4, column=col, value=text)
        c.fill = navy
        c.font = Font(bold=True, color="FFFFFF")
    ws.freeze_panes = "A5"

    r = 5
    subtotal_cells = []
    for block in report.blocks:
        first = r
        lines = block.lines or [("No sales", 0)]
        for client, qty in lines:
            ws.cell(row=r, column=1, value=block.rider)
            ws.cell(row=r, column=2, value=client)
            ws.cell(row=r, column=3, value=qty)
            r += 1
        ws.cell(row=r, column=1, value=f"{block.rider} Subtotal")
        ws.cell(row=r, column=3, value=f"=SUM(C{first}:C{r - 1})")
        for col in range(1, 4):
            ws.cell(row=r, column=col).font = Font(bold=True)
            ws.cell(row=r, column=col).fill = light_blue
        subtotal_cells.append(f"C{r}")
        r += 1

    ws.cell(row=r, column=1, value="TOTAL")
    ws.cell(row=r, column=3, value="=" + "+".join(subtotal_cells) if subtotal_cells else 0)
    for col in range(1, 4):
        ws.cell(row=r, column=col).font = Font(bold=True)
        ws.cell(row=r, column=col).fill = yellow

    for row in ws.iter_rows(min_row=4, max_row=r, max_col=3):
        for c in row:
            c.border = border
    for row in ws.iter_rows(min_row=5, max_row=r, min_col=3, max_col=3):
        row[0].number_format = "0.0"

    # pixel widths 150 / 200 / 80 -> Excel character widths
    ws.column_dimensions["A"].width = 21
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 11
    return wb


def save_report(report: DailyReport, reports_dir: Path) -> Path:
    reports_dir.mkdir(parents=True, exist_ok=True)
    path = reports_dir / report.filename
    report_workbook(report).save(path)
    return path


def report_bytes(report: DailyReport) -> bytes:
    buf = BytesIO()
    report_workbook(report).save(buf)
    return buf.getvalue()
