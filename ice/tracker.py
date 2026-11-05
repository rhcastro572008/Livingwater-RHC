"""Read and write Tracker.xlsx.

Workbook layout (kept compatible with the existing tracker):
  * "Data Source"  columns: Rider | Date | Client | Count
  * "Client List"  columns: Client | STATUS   (ACTIVE / INACTIVE)

Rows are appended with openpyxl so existing formatting, formulas and any other
sheets in the workbook are preserved. A backup copy is written before every save.
"""
from __future__ import annotations

import math
import re
import shutil
import threading
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.datetime import from_excel

DATA_SHEET = "Data Source"
CLIENT_SHEET = "Client List"
DATA_HEADERS = ["Rider", "Date", "Client", "Count"]
CLIENT_HEADERS = ["Client", "STATUS"]
DATE_FORMAT = "m/d/yyyy"

_lock = threading.Lock()


class TrackerError(Exception):
    pass


class DateAlreadyLoaded(TrackerError):
    def __init__(self, report_date: date):
        self.report_date = report_date
        super().__init__(f"{fmt_long(report_date)} is already in the tracker")


@dataclass(frozen=True)
class Sale:
    rider: str
    date: date
    client: str
    count: float


@dataclass(frozen=True)
class Client:
    name: str
    status: str  # ACTIVE / INACTIVE / "" (blank)

    @property
    def active(self) -> bool:
        return self.status.upper() != "INACTIVE"


def fmt_long(d: date) -> str:
    """November 3, 2026"""
    return f"{d:%B} {d.day}, {d.year}"


def parse_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial number (a date cell formatted as General/Number)
        try:
            return from_excel(value).date() if 20000 < value < 80000 else None
        except (ValueError, OverflowError):
            return None
    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_qty(value) -> float | None:
    """Parse a quantity: 2, 2.5, "1/2", "1 1/2". Returns None if it is not a finite number."""
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        num = float(value)
    else:
        text = str(value).strip().replace(",", "")
        m = re.fullmatch(r"(?:(\d+)\s+)?(\d+)\s*/\s*(\d+)", text)
        if m:
            whole, top, bottom = m.groups()
            if int(bottom) == 0:
                return None
            num = int(whole or 0) + int(top) / int(bottom)
        else:
            try:
                num = float(text)
            except ValueError:
                return None
    return num if math.isfinite(num) else None


_num = parse_qty


def _get_sheet(wb, name: str):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    return None


def _style_header(ws) -> None:
    fill = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill


def create_tracker(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = DATA_SHEET
    ws.append(DATA_HEADERS)
    _style_header(ws)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 10
    ws.freeze_panes = "A2"
    cl = wb.create_sheet(CLIENT_SHEET)
    cl.append(CLIENT_HEADERS)
    _style_header(cl)
    cl.column_dimensions["A"].width = 28
    cl.column_dimensions["B"].width = 12
    cl.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def ensure_tracker(path: Path) -> None:
    if not path.exists():
        create_tracker(path)


def _open(path: Path, read_only: bool = False):
    ensure_tracker(path)
    wb = load_workbook(path, read_only=read_only, data_only=read_only)
    if _get_sheet(wb, DATA_SHEET) is None:
        raise TrackerError(f'Tracker.xlsx has no "{DATA_SHEET}" worksheet')
    return wb


def validate_workbook(path: Path) -> None:
    """Raise TrackerError if the file is not a usable tracker."""
    try:
        wb = load_workbook(path, read_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl raises many types
        raise TrackerError(f"Not a valid .xlsx file: {exc}") from exc
    if _get_sheet(wb, DATA_SHEET) is None:
        raise TrackerError(f'The workbook has no "{DATA_SHEET}" worksheet')
    wb.close()


def load_sales(path: Path) -> list[Sale]:
    wb = _open(path, read_only=True)
    ws = _get_sheet(wb, DATA_SHEET)
    sales: list[Sale] = []
    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        if not row or all(v in (None, "") for v in row):
            continue
        rider, d, client, count = (list(row) + [None] * 4)[:4]
        d = parse_date(d)
        count = _num(count)
        if not rider or d is None or not client or count is None:
            continue
        sales.append(Sale(str(rider).strip(), d, str(client).strip(), count))
    wb.close()
    return sales


def load_clients(path: Path) -> list[Client]:
    wb = _open(path, read_only=True)
    ws = _get_sheet(wb, CLIENT_SHEET)
    clients: list[Client] = []
    if ws is not None:
        seen = set()
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            name = (row[0] if row else None) or ""
            name = str(name).strip()
            if not name or name.upper() in seen:
                continue
            seen.add(name.upper())
            status = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ""
            clients.append(Client(name, status))
    wb.close()
    return clients


def loaded_dates(path: Path) -> set[date]:
    return {s.date for s in load_sales(path)}


def _last_data_row(ws, max_col: int) -> int:
    last = 1
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_col=max_col, values_only=True), start=1):
        if any(v not in (None, "") for v in row):
            last = idx
    return last


def backup(path: Path, backups_dir: Path, reason: str = "before load", keep: int = 60) -> Path | None:
    if not path.exists():
        return None
    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    dest = backups_dir / f"Tracker (backup {reason} {stamp}).xlsx"
    n = 1
    while dest.exists():  # two saves in the same second
        n += 1
        dest = backups_dir / f"Tracker (backup {reason} {stamp}-{n}).xlsx"
    shutil.copy2(path, dest)
    prune_backups(backups_dir, keep)
    return dest


def prune_backups(backups_dir: Path, keep: int) -> None:
    if keep <= 0:
        return
    files = sorted(backups_dir.glob("Tracker (backup *.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    for old in files[keep:]:
        old.unlink(missing_ok=True)


def append_sales(path: Path, sales: list[Sale], backups_dir: Path) -> Path | None:
    """Append rows for one or more *new* dates. Refuses dates already present."""
    if not sales:
        raise TrackerError("No rows to save")
    for s in sales:
        if parse_qty(s.count) is None or s.count <= 0:
            raise TrackerError(f"Invalid quantity {s.count!r} for {s.client}")
    with _lock:
        existing = loaded_dates(path)
        for d in sorted({s.date for s in sales}):
            if d in existing:
                raise DateAlreadyLoaded(d)
        backup_path = backup(path, backups_dir)
        wb = _open(path)
        ws = _get_sheet(wb, DATA_SHEET)
        row_idx = _last_data_row(ws, 4)
        date_fmt = ws.cell(row=row_idx, column=2).number_format if row_idx > 1 else DATE_FORMAT
        if date_fmt in (None, "General"):
            date_fmt = DATE_FORMAT
        for s in sorted(sales, key=lambda s: s.date):  # stable: keeps report order within a date
            row_idx += 1
            ws.cell(row=row_idx, column=1, value=s.rider)
            c = ws.cell(row=row_idx, column=2, value=s.date)
            c.number_format = date_fmt
            ws.cell(row=row_idx, column=3, value=s.client)
            ws.cell(row=row_idx, column=4, value=int(s.count) if float(s.count).is_integer() else s.count)
        wb.save(path)
        return backup_path


def delete_date(path: Path, report_date: date, backups_dir: Path) -> int:
    """Remove every Data Source row for a date (used to redo a day). Returns rows removed."""
    with _lock:
        wb = _open(path)
        ws = _get_sheet(wb, DATA_SHEET)
        to_delete = [
            idx
            for idx, row in enumerate(ws.iter_rows(min_row=2, max_col=4, values_only=True), start=2)
            if parse_date(row[1]) == report_date
        ]
        if not to_delete:
            return 0
        backup(path, backups_dir, reason="before delete")
        for idx in reversed(to_delete):
            ws.delete_rows(idx)
        wb.save(path)
        return len(to_delete)


def replace_tracker(path: Path, new_file: Path, backups_dir: Path) -> None:
    """Swap in an uploaded tracker (already validated), keeping a backup of the old one."""
    with _lock:
        backup(path, backups_dir, reason="before replace")
        shutil.move(str(new_file), str(path))


def upsert_client(path: Path, name: str, status: str, backups_dir: Path) -> None:
    name = name.strip().upper()
    status = status.strip().upper()
    if not name:
        raise TrackerError("Client name is required")
    if status not in ("ACTIVE", "INACTIVE"):
        raise TrackerError("Status must be ACTIVE or INACTIVE")
    with _lock:
        backup(path, backups_dir, reason="before client edit")
        wb = _open(path)
        ws = _get_sheet(wb, CLIENT_SHEET)
        if ws is None:
            ws = wb.create_sheet(CLIENT_SHEET)
            ws.append(CLIENT_HEADERS)
            _style_header(ws)
        for row in ws.iter_rows(min_row=2, max_col=2):
            if row[0].value and str(row[0].value).strip().upper() == name:
                row[1].value = status
                break
        else:
            r = _last_data_row(ws, 2) + 1
            ws.cell(row=r, column=1, value=name)
            ws.cell(row=r, column=2, value=status)
        wb.save(path)


# ---------------------------------------------------------------- name matching

def _key(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip().upper())


def match_client(raw: str, clients: list[Client], aliases: dict[str, str]) -> tuple[str, list[str]]:
    """Map a name read off the board to the Client List spelling.

    Returns (name_to_store, flags). Flags are short human-readable warnings.
    """
    key = _key(raw)
    by_key = {_key(c.name): c for c in clients}
    flags: list[str] = []

    alias = aliases.get(key)
    if alias and _key(alias) in by_key:
        client = by_key[_key(alias)]
    elif key in by_key:
        client = by_key[key]
    else:
        compact = {re.sub(r"[^A-Z0-9]", "", k): c for k, c in by_key.items()}
        client = compact.get(re.sub(r"[^A-Z0-9]", "", key))

    if client is None:
        if clients:
            flags.append("not on the Client List (new client?)")
        return key, flags

    # Ambiguity: other clients that start with the same word(s), e.g. ROSE vs ROSE LOTTO.
    others = [
        c.name for k, c in by_key.items()
        if k != _key(client.name) and re.match(re.escape(_key(client.name)) + r"[\s(]", k)
    ]
    if others:
        flags.append("could also be " + ", ".join(sorted(others)))
    if client.status == "INACTIVE":
        flags.append("marked INACTIVE on the Client List")
    elif client.status == "":
        flags.append("has a blank status on the Client List")
    return client.name, flags
