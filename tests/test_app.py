from datetime import date
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook
from PIL import Image

from app import create_app
from ice import dashboard, tracker
from ice.extract import BoardEntry, BoardReading, to_draft
from ice.report import build_report, report_workbook
from ice.tracker import Client, Sale

RIDERS = ["Amped", "Khurtz", "Ram", "Mark", "Mike", "Jomar", "Pickup"]


@pytest.fixture
def tracker_path(tmp_path):
    """A tracker shaped like the real one, with some history and a Client List."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Source"
    ws.append(["Rider", "Date", "Client", "Count"])
    ws.append(["Ram", date(2026, 11, 1), "LISA", 2])
    ws.cell(row=2, column=2).number_format = "m/d/yyyy"
    ws.append(["Mike", date(2026, 11, 1), "JOY STORE", 0.5])
    ws.cell(row=3, column=2).number_format = "m/d/yyyy"
    cl = wb.create_sheet("Client List")
    cl.append(["Client", "STATUS"])
    for row in [("LISA", "ACTIVE"), ("JOY STORE", "ACTIVE"), ("ROSE", "ACTIVE"), ("ROSE LOTTO", "ACTIVE"),
                ("LITO FOODHAUSE", "ACTIVE"), ("OLD SHOP", "INACTIVE"), ("TAIPEI", "")]:
        cl.append(row)
    wb.create_sheet("Pivot")["A1"] = "keep me"
    path = tmp_path / "Tracker.xlsx"
    wb.save(path)
    return path


def fake_reader(image, media_type, riders, clients, model):
    return BoardReading(
        entries=[
            BoardEntry(rider="Ram", client="Lisa", quantity=2, unclear=False, note=""),
            BoardEntry(rider="Ram", client="lisa", quantity=1, unclear=False, note=""),
            BoardEntry(rider="Jomor", client="Lito's", quantity=0.5, unclear=True, note="5 or 1/2?"),
            BoardEntry(rider="Mike", client="Rose", quantity=3, unclear=False, note=""),
            BoardEntry(rider="Mike", client="Newbie", quantity=1, unclear=False, note=""),
        ],
        riders_off=["Mark"], left_out=["~ Bella"], notes=[],
    )


@pytest.fixture
def client(tmp_path, tracker_path):
    app = create_app({
        "TESTING": True,
        "TRACKER_PATH": tracker_path,
        "REPORTS_DIR": tmp_path / "reports",
        "RAW_DIR": tmp_path / "raw",
        "UPLOADS_DIR": tmp_path / "uploads",
        "BACKUPS_DIR": tmp_path / "backups",
        "RIDERS": RIDERS,
        "READER": fake_reader,
        "APP_PASSWORD": "",
        "CSRF_ENABLED": False,
    })
    return app.test_client()


def jpeg() -> BytesIO:
    buf = BytesIO()
    Image.new("RGB", (40, 30), "white").save(buf, "JPEG")
    buf.seek(0)
    return buf


# ------------------------------------------------------------------ units

def test_match_client_rules():
    clients = [Client("LISA", "ACTIVE"), Client("ROSE", "ACTIVE"), Client("ROSE LOTTO", "ACTIVE"),
               Client("CATH", "ACTIVE"), Client("CATH(JOHN)", "ACTIVE"), Client("OLD SHOP", "INACTIVE"),
               Client("JOY STORE", "ACTIVE"), Client("TAIPEI", "")]
    aliases = {"JOY": "JOY STORE"}
    assert tracker.match_client("lisa", clients, aliases) == ("LISA", [])
    assert tracker.match_client("Joy", clients, aliases) == ("JOY STORE", [])
    name, flags = tracker.match_client("Rose", clients, aliases)
    assert name == "ROSE" and "ROSE LOTTO" in flags[0]
    assert "CATH(JOHN)" in tracker.match_client("cath", clients, aliases)[1][0]
    assert "INACTIVE" in tracker.match_client("old shop", clients, aliases)[1][0]
    assert "blank status" in tracker.match_client("Taipei", clients, aliases)[1][0]
    assert tracker.match_client("Zed", clients, aliases) == ("ZED", ["not on the Client List (new client?)"])


def test_to_draft_merges_and_flags():
    clients = [Client("LISA", "ACTIVE"), Client("LITO FOODHAUSE", "ACTIVE")]
    draft = to_draft(fake_reader(None, None, RIDERS, clients, ""), RIDERS, clients, {"LITO'S": "LITO FOODHAUSE"})
    rows = {(r.rider, r.client): r for r in draft.rows}
    assert rows[("Ram", "LISA")].qty == 3
    assert rows[("Jomar", "LITO FOODHAUSE")].qty == 0.5
    assert any("unclear" in f for f in rows[("Jomar", "LITO FOODHAUSE")].flags)
    assert [r.rider for r in draft.rows] == sorted([r.rider for r in draft.rows], key=RIDERS.index)
    assert any("Mark" in n for n in draft.notes) and any("Bella" in n for n in draft.notes)


def test_report_layout_and_totals():
    d = date(2026, 11, 3)
    sales = [Sale("Ram", d, "LISA", 2), Sale("Ram", d, "JOY STORE", 1.5), Sale("Amped", d, "ROSE", 4)]
    rep = build_report(d, sales, RIDERS)
    assert rep.total == 7.5
    ws = report_workbook(rep)["Report"]
    assert ws["A1"].value == "Ice Monitoring Report"
    assert ws["A2"].value == "November 3, 2026"
    assert [c.value for c in ws[4]] == ["Rider", "Customer", "Qty"]
    assert ws["A5"].value == "Amped" and ws["A6"].value == "Amped Subtotal"
    assert ws["C6"].value == "=SUM(C5:C5)"
    assert ws["B7"].value == "No sales"  # Khurtz
    last = ws.max_row
    assert ws.cell(row=last, column=1).value == "TOTAL"
    assert ws.cell(row=last, column=3).value.count("+") == len(RIDERS) - 1


def test_append_preserves_other_sheets_and_refuses_duplicates(tracker_path, tmp_path):
    d = date(2026, 11, 3)
    tracker.append_sales(tracker_path, [Sale("Ram", d, "LISA", 2)], tmp_path / "b")
    wb = load_workbook(tracker_path)
    assert wb["Pivot"]["A1"].value == "keep me"
    ws = wb["Data Source"]
    assert ws.max_row == 4 and ws["B4"].number_format == "m/d/yyyy"
    assert list(( tmp_path / "b").iterdir())
    with pytest.raises(tracker.DateAlreadyLoaded):
        tracker.append_sales(tracker_path, [Sale("Ram", d, "LISA", 1)], tmp_path / "b")
    assert tracker.delete_date(tracker_path, d, tmp_path / "b") == 1
    assert d not in tracker.loaded_dates(tracker_path)


def test_dashboard_numbers():
    sales = [
        Sale("Ram", date(2026, 11, 1), "LISA", 2), Sale("Ram", date(2026, 11, 2), "LISA", 3),
        Sale("Mike", date(2026, 11, 2), "ROSE", 1), Sale("Mike", date(2026, 10, 30), "OLD", 9),
    ]
    perf = {p.rider: p for p in dashboard.rider_performance(sales, 2026, 11, RIDERS)}
    assert perf["Ram"].total == 5 and perf["Ram"].days_with_sales == 2 and perf["Mike"].total == 1
    daily = dashboard.daily_sales(sales, 2026, 11)
    assert len(daily) == 30 and daily[1] == (date(2026, 11, 2), 4)

    clients = [Client("LISA", "ACTIVE"), Client("ROSE", "ACTIVE"), Client("NEVER", "ACTIVE"),
               Client("GONE", "INACTIVE")]
    gaps = dashboard.clients_not_ordering(sales, clients, as_of=date(2026, 11, 7), days=5)
    names = [g.client for g in gaps]
    # window is Nov 3-7: nobody ordered in it
    assert names == ["LISA", "ROSE", "OLD", "NEVER"]
    assert gaps[0].days_since == 5 and gaps[-1].last_order is None
    gaps = dashboard.clients_not_ordering(sales, clients, as_of=date(2026, 11, 6), days=5)
    assert "LISA" not in [g.client for g in gaps]  # Nov 2 is inside Nov 2-6


# ------------------------------------------------------------------ end to end

def test_upload_review_save_flow(client, tracker_path, tmp_path):
    r = client.post("/upload", data={"report_date": "2026-11-03", "photo": (jpeg(), "board.jpg")},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    assert "LITO FOODHAUSE" in html and "Jomar" in html and "not on the Client List" in html
    token = html.split('name="token" value="')[1].split('"')[0]

    r = client.post("/save", data={
        "report_date": "2026-11-03", "token": token,
        "rider": ["Ram", "Jomar", "Mike", "Mike", ""], "client": ["LISA", "LITO FOODHAUSE", "ROSE", "Newbie", ""],
        "qty": ["3", "1/2", "3", "1", ""],
    })
    assert r.status_code == 302 and r.headers["Location"].endswith("/reports/2026-11-03")
    sales = [s for s in tracker.load_sales(tracker_path) if s.date == date(2026, 11, 3)]
    assert sum(s.count for s in sales) == 7.5 and {s.client for s in sales} >= {"NEWBIE"}
    assert (tmp_path / "reports" / "Ice Monitoring Report - November 3, 2026.xlsx").exists()
    assert (tmp_path / "raw" / "2026-11-03.jpg").exists()

    assert client.get("/reports/2026-11-03").status_code == 200
    dl = client.get("/reports/2026-11-03/download")
    assert dl.status_code == 200 and dl.data[:2] == b"PK"

    # same date again is refused
    r = client.post("/upload", data={"report_date": "2026-11-03", "photo": (jpeg(), "b.jpg")},
                    content_type="multipart/form-data", follow_redirects=True)
    assert "already in Tracker.xlsx" in r.get_data(as_text=True)


def test_save_validation(client):
    r = client.post("/save", data={"report_date": "2026-11-04", "rider": ["Nobody"], "client": ["LISA"], "qty": ["x"]})
    assert r.status_code == 400
    html = r.get_data(as_text=True)
    assert "choose a rider" in html and "quantity must be a number" in html


def test_pages_render(client):
    for url in ["/", "/reports", "/dashboard", "/dashboard?month=2026-11&as_of=2026-11-07", "/clients", "/tracker"]:
        assert client.get(url).status_code == 200, url
    html = client.get("/dashboard?month=2026-11&as_of=2026-11-07").get_data(as_text=True)
    assert "ROSE LOTTO" in html and "OLD SHOP" not in html


def test_manual_entry_without_photo(client):
    r = client.post("/upload", data={"report_date": "2026-11-04", "manual": "1"})
    assert r.status_code == 200 and "Check the rows" in r.get_data(as_text=True)


def test_clients_and_password(client, tracker_path, tmp_path):
    client.post("/clients", data={"name": "newbie", "status": "ACTIVE"})
    assert "NEWBIE" in [c.name for c in tracker.load_clients(tracker_path)]
    app = create_app({"TRACKER_PATH": tracker_path, "REPORTS_DIR": tmp_path / "r", "RAW_DIR": tmp_path / "w",
                      "UPLOADS_DIR": tmp_path / "u", "BACKUPS_DIR": tmp_path / "k", "APP_PASSWORD": "pw",
                      "CSRF_ENABLED": False})
    c = app.test_client()
    assert c.get("/").status_code == 401
    import base64
    assert c.get("/", headers={"Authorization": "Basic " + base64.b64encode(b"x:pw").decode()}).status_code == 200


# ------------------------------------------------------------------ regressions

def test_blank_rows_with_rider_selected_are_ignored(client, tracker_path):
    # the review form always sends a rider for empty lines (it's a dropdown)
    r = client.post("/save", data={
        "report_date": "2026-11-05",
        "rider": ["Ram", "Mike", "Amped", "Amped"], "client": ["LISA", "ROSE", "", ""], "qty": ["2", "1", "", ""],
    })
    assert r.status_code == 302
    assert len([s for s in tracker.load_sales(tracker_path) if s.date == date(2026, 11, 5)]) == 2


@pytest.mark.parametrize("bad", ["nan", "inf", "-1", "0", "abc", "1/0"])
def test_invalid_quantities_rejected(client, tracker_path, bad):
    r = client.post("/save", data={"report_date": "2026-11-06", "rider": ["Ram"], "client": ["LISA"], "qty": [bad]})
    assert r.status_code == 400
    assert f'value="{bad}"' in r.get_data(as_text=True) or bad in ("-1", "0")  # typed value is kept
    assert date(2026, 11, 6) not in tracker.loaded_dates(tracker_path)


def test_quantity_formats_and_duplicates_merged(client, tracker_path):
    r = client.post("/save", data={
        "report_date": "2026-11-07",
        "rider": ["Ram", "Ram", "Mike"], "client": ["LISA", "lisa", "ROSE"], "qty": ["1 1/2", "1,000", "0.25"],
    })
    assert r.status_code == 302
    sales = [s for s in tracker.load_sales(tracker_path) if s.date == date(2026, 11, 7)]
    assert [(s.rider, s.client, s.count) for s in sales] == [("Ram", "LISA", 1001.5), ("Mike", "ROSE", 0.25)]


def test_review_inputs_keep_exact_quantity(client):
    def reader(*a):
        return BoardReading(entries=[BoardEntry(rider="Ram", client="LISA", quantity=0.25, unclear=False, note=""),
                                     BoardEntry(rider="Mike", client="ROSE", quantity=1500, unclear=False, note="")],
                            riders_off=[], left_out=[], notes=[])
    client.application.config["READER"] = reader
    r = client.post("/upload", data={"report_date": "2026-11-08", "photo": (jpeg(), "b.jpg")},
                    content_type="multipart/form-data")
    html = r.get_data(as_text=True)
    assert 'value="0.25"' in html and 'value="1500"' in html


def test_excel_serial_dates_are_read(tracker_path):
    wb = load_workbook(tracker_path)
    wb["Data Source"].append(["Ram", 46330, "LISA", 1])  # 46330 = 11/4/2026 stored as a plain number
    wb.save(tracker_path)
    assert date(2026, 11, 4) in tracker.loaded_dates(tracker_path)


def test_report_merges_client_case():
    d = date(2026, 11, 3)
    rep = build_report(d, [Sale("Ram", d, "LISA", 1), Sale("Ram", d, "Lisa", 1)], ["Ram"])
    assert rep.blocks[0].lines == [("LISA", 2)]


def test_backups_are_pruned(tracker_path, tmp_path):
    for _ in range(5):
        tracker.backup(tracker_path, tmp_path / "b", keep=3)
    assert len(list((tmp_path / "b").iterdir())) == 3


def test_csrf_token_required(tmp_path, tracker_path):
    app = create_app({"TRACKER_PATH": tracker_path, "REPORTS_DIR": tmp_path / "r", "RAW_DIR": tmp_path / "w",
                      "UPLOADS_DIR": tmp_path / "u", "BACKUPS_DIR": tmp_path / "k", "APP_PASSWORD": ""})
    c = app.test_client()
    # a forged post (no token) is refused and changes nothing
    c.post("/reports/2026-11-01/delete")
    assert date(2026, 11, 1) in tracker.loaded_dates(tracker_path)
    html = c.get("/reports/2026-11-01").get_data(as_text=True)
    token = html.split('name="csrf_token" value="')[1].split('"')[0]
    c.post("/reports/2026-11-01/delete", data={"csrf_token": token})
    assert date(2026, 11, 1) not in tracker.loaded_dates(tracker_path)


def test_non_ascii_password_does_not_crash(tmp_path, tracker_path):
    import base64
    app = create_app({"TRACKER_PATH": tracker_path, "REPORTS_DIR": tmp_path / "r", "RAW_DIR": tmp_path / "w",
                      "UPLOADS_DIR": tmp_path / "u", "BACKUPS_DIR": tmp_path / "k", "APP_PASSWORD": "señor"})
    c = app.test_client()
    bad = base64.b64encode("x:mañana".encode()).decode()
    good = base64.b64encode("x:señor".encode()).decode()
    assert c.get("/", headers={"Authorization": "Basic " + bad}).status_code == 401
    assert c.get("/", headers={"Authorization": "Basic " + good}).status_code == 200


def test_today_uses_business_timezone(client):
    from datetime import datetime
    from zoneinfo import ZoneInfo
    html = client.get("/").get_data(as_text=True)
    assert f'max="{datetime.now(ZoneInfo("Asia/Manila")).date().isoformat()}"' in html
