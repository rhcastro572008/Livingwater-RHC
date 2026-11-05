# Livingwater-RHC · Ice Monitoring

A web app for the Living Water ice business. It turns a photo of the daily whiteboard into rows in **Tracker.xlsx**, makes the daily **Ice Monitoring Report**, and shows a monthly **dashboard**.

## What it does

1. **Upload a photo.** Take a photo of the board on a phone and pick the report date. Claude reads the board column by column: the rider name is at the top, and each entry under it is `quantity client` ("1/2" = 0.5).
2. **Check the rows.** The photo and the rows it found are shown side by side. Each row has a rider, client and quantity, and you can edit any of them. Rows are flagged when the handwriting was unclear, when the client is not on the Client List, is INACTIVE, or could be another client (ROSE vs ROSE LOTTO), or when the same client was listed twice under one rider (those are combined). Nothing is saved until you press **Save**.
3. **Save.** The rows are added to the `Data Source` sheet of `Tracker.xlsx` (Rider | Date | Client | Count). Other sheets, formatting and formulas are kept, and a backup copy is made before every write. A date that is already in the tracker is refused, so a day can't be loaded twice.
4. **Daily report.** `Ice Monitoring Report - <Month D, YYYY>.xlsx` is created with one block per rider (Amped, Khurtz, Ram, Mark, Mike, Jomar, Pickup), `=SUM` subtotals and a TOTAL row. You can view it, print it or download it from the **Reports** page.
5. **Dashboard** (pick a month):
   - **Rider performance:** total, share, days worked, average per day, deliveries, clients served and best day for each rider.
   - **Daily sales:** total per day for the month, all riders combined. Hover over a bar to see the split by rider, or open the table view.
   - **Clients who did not order in the past 5 days:** active clients with no order in the 5 days up to the "as of" date, with their last order date, last rider and quantity this month.
6. **Clients:** add clients, or mark them ACTIVE or INACTIVE (this edits the `Client List` sheet). Names that appear in sales but are not on the list are shown here so you can add them with one click.
7. **Tracker:** download the current `Tracker.xlsx`, or upload your existing one to start from your history.

## Setup

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # then put your ANTHROPIC_API_KEY in .env
flask --app app run --host 0.0.0.0 --port 5000
```

Open http://localhost:5000. Phones on the same Wi-Fi can use `http://<computer-ip>:5000`.

- Go to **Tracker → Use your existing Tracker.xlsx** and upload your current file. It needs a `Data Source` sheet and, ideally, a `Client List` sheet.
- Without an `ANTHROPIC_API_KEY`, everything still works except automatic reading: you upload the photo and type the rows next to it.
- Set `APP_PASSWORD` if the site can be reached by anyone outside the shop, and set a fixed `SECRET_KEY`.
- "Today" and the 5-day window use `APP_TIMEZONE` (default `Asia/Manila`), not the server clock.
- Files are kept under `data/`: `Tracker.xlsx`, `reports/`, `raw/` (one photo per date), `backups/`. Back this folder up. It is not committed to git.

To run it in production, use a real WSGI server, e.g. `pip install waitress && waitress-serve --port 5000 --call app:create_app`.

## Fixing a mistake

On a report's page, use **Remove this day from the tracker** (a backup is kept), then upload the photo again.

## Project layout

```
app.py              Flask routes
ice/config.py       settings (riders, paths, model) from environment variables
ice/extract.py      reads the whiteboard photo with Claude; normalizes and flags rows
ice/tracker.py      reads/writes Tracker.xlsx; client name matching
ice/report.py       builds the daily Ice Monitoring Report workbook
ice/dashboard.py    rider performance, daily sales, clients not ordering
templates/, static/ pages and styles (Chart.js is bundled in static/vendor)
tests/              pytest suite: `python -m pytest`
```
