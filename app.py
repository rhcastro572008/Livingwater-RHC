"""Living Water RHC - Ice Monitoring web app.

Run:  flask --app app run --debug     (or: python app.py)
"""
from __future__ import annotations

import secrets
import shutil
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from flask import (
    Flask, Response, abort, flash, redirect, render_template, request, send_file, session, url_for,
)

from ice import config, dashboard, extract, tracker
from ice.report import build_report, report_bytes, save_report
from ice.tracker import Sale, fmt_long


def create_app(overrides: dict | None = None) -> Flask:
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=config.SECRET_KEY,
        MAX_CONTENT_LENGTH=config.MAX_UPLOAD_MB * 1024 * 1024,
        TRACKER_PATH=config.TRACKER_PATH,
        REPORTS_DIR=config.REPORTS_DIR,
        RAW_DIR=config.RAW_DIR,
        UPLOADS_DIR=config.UPLOADS_DIR,
        BACKUPS_DIR=config.BACKUPS_DIR,
        RIDERS=config.RIDERS,
        CLIENT_ALIASES=config.CLIENT_ALIASES,
        INACTIVE_DAYS=config.INACTIVE_DAYS,
        ANTHROPIC_MODEL=config.ANTHROPIC_MODEL,
        APP_PASSWORD=config.APP_PASSWORD,
        APP_TIMEZONE=config.APP_TIMEZONE,
        CSRF_ENABLED=True,
        READER=None,  # tests can inject a fake board reader
    )
    if overrides:
        app.config.update(overrides)
    for key in ("REPORTS_DIR", "RAW_DIR", "UPLOADS_DIR", "BACKUPS_DIR"):
        Path(app.config[key]).mkdir(parents=True, exist_ok=True)
    tracker.ensure_tracker(Path(app.config["TRACKER_PATH"]))

    def tp() -> Path:
        return Path(app.config["TRACKER_PATH"])

    def riders() -> list[str]:
        return app.config["RIDERS"]

    def today() -> date:
        """Today's date where the business is, even if the server runs on UTC."""
        return datetime.now(ZoneInfo(app.config["APP_TIMEZONE"])).date()

    # ------------------------------------------------------------- auth (optional)
    @app.before_request
    def require_password():
        pw = app.config["APP_PASSWORD"]
        if not pw or request.endpoint == "static":
            return None
        auth = request.authorization
        if auth and secrets.compare_digest((auth.password or "").encode(), pw.encode()):
            return None
        return Response("Password required", 401, {"WWW-Authenticate": 'Basic realm="Ice Monitoring"'})

    def csrf_token() -> str:
        if "csrf" not in session:
            session["csrf"] = secrets.token_hex(16)
        return session["csrf"]

    @app.before_request
    def check_csrf():
        """Every form carries a per-session token, so another website can't post forms
        (e.g. a hidden "delete this day") using the browser's saved password."""
        if request.method != "POST" or not app.config["CSRF_ENABLED"]:
            return None
        sent = request.form.get("csrf_token", "")
        if not sent or not secrets.compare_digest(sent, session.get("csrf", "")):
            flash("The page had expired. Please try again.", "error")
            return redirect(request.referrer or url_for("upload"))
        return None

    @app.template_filter("qty")
    def qty_filter(v) -> str:
        v = float(v or 0)
        return f"{v:,.0f}" if v.is_integer() else f"{v:,.1f}"

    @app.template_filter("qty_input")
    def qty_input_filter(v) -> str:
        """Exact value for an editable input: no thousands separator, no rounding."""
        if v in (None, ""):
            return ""
        try:
            return f"{float(v):.10g}"
        except (TypeError, ValueError):
            return str(v)  # keep what the user typed so they can correct it

    @app.template_filter("longdate")
    def longdate_filter(d) -> str:
        return fmt_long(d) if d else "—"

    @app.context_processor
    def inject():
        return {"nav": request.endpoint, "today": today(), "csrf_token": csrf_token}

    # ------------------------------------------------------------------ upload
    @app.get("/")
    def upload():
        return render_template(
            "upload.html",
            reader_ready=bool(app.config["READER"]) or extract.is_configured(),
            loaded=sorted(tracker.loaded_dates(tp()), reverse=True)[:7],
        )

    @app.post("/upload")
    def upload_post():
        report_date = tracker.parse_date(request.form.get("report_date"))
        if report_date is None:
            flash("Choose the report date (the date written on the board, not today's date).", "error")
            return redirect(url_for("upload"))
        if report_date in tracker.loaded_dates(tp()):
            flash(f"{fmt_long(report_date)} is already in Tracker.xlsx. "
                  "Delete it on the Reports page first if you want to redo it.", "error")
            return redirect(url_for("upload"))

        clean_old_uploads()
        photo = request.files.get("photo")
        manual = request.form.get("manual") == "1"
        token = None
        draft = extract.Draft([], [])
        if photo and photo.filename:
            try:
                image, media_type = extract.prepare_image(photo.read())
            except extract.ExtractionError as exc:
                flash(str(exc), "error")
                return redirect(url_for("upload"))
            token = secrets.token_hex(8)
            (Path(app.config["UPLOADS_DIR"]) / f"{token}.jpg").write_bytes(image)

            if not manual:
                clients = tracker.load_clients(tp())
                reader = app.config["READER"] or extract.read_board
                if app.config["READER"] or extract.is_configured():
                    try:
                        reading = reader(image, media_type, riders(), clients, app.config["ANTHROPIC_MODEL"])
                        draft = extract.to_draft(reading, riders(), clients, app.config["CLIENT_ALIASES"])
                        if not draft.rows:
                            flash("No entries were found on the photo. Add the rows by hand.", "warning")
                    except extract.ExtractionError as exc:
                        flash(f"{exc}", "warning")
                else:
                    flash("Automatic photo reading is off (no ANTHROPIC_API_KEY). "
                          "Type the rows from the photo below.", "warning")
        elif not manual:
            flash("Choose a photo of the board, or pick “Enter rows by hand”.", "error")
            return redirect(url_for("upload"))

        return render_review(report_date, token, draft.rows, draft.notes)

    def clean_old_uploads(max_age_hours: int = 48) -> None:
        """Delete photos from reviews that were never saved."""
        cutoff = time.time() - max_age_hours * 3600
        for f in Path(app.config["UPLOADS_DIR"]).glob("*"):
            try:
                if f.is_file() and f.stat().st_mtime < cutoff:
                    f.unlink()
            except OSError:
                pass

    def render_review(report_date, token, rows, notes, status=200):
        return render_template(
            "review.html",
            report_date=report_date,
            token=token,
            rows=rows,
            notes=notes,
            riders=riders(),
            clients=tracker.load_clients(tp()),
        ), status

    @app.get("/uploads/<token>.jpg")
    def upload_image(token):
        if not token.isalnum():
            abort(404)
        path = Path(app.config["UPLOADS_DIR"]) / f"{token}.jpg"
        if not path.exists():
            abort(404)
        return send_file(path, mimetype="image/jpeg")

    # ------------------------------------------------------------------ save
    @app.post("/save")
    def save():
        report_date = tracker.parse_date(request.form.get("report_date"))
        token = request.form.get("token") or None
        if token and not token.isalnum():
            token = None
        if report_date is None:
            abort(400)

        riders_in = request.form.getlist("rider")
        clients_in = request.form.getlist("client")
        qty_in = request.form.getlist("qty")
        clients = tracker.load_clients(tp())

        rows, errors = [], []
        merged: dict[tuple[str, str], extract.DraftRow] = {}
        for i, (r, c, q) in enumerate(zip(riders_in, clients_in, qty_in), start=1):
            r, c, q = r.strip(), c.strip(), q.strip()
            if not (c or q):
                continue  # empty line (the rider dropdown always has a value, so ignore it)
            if r not in riders():
                errors.append(f"Row {i}: choose a rider from the list ({', '.join(riders())}).")
            if not c:
                errors.append(f"Row {i}: client is missing.")
            qty = tracker.parse_qty(q)
            if qty is None or qty <= 0:
                errors.append(f"Row {i}: quantity must be a number greater than 0 (e.g. 2, 0.5 or 1/2).")
            name, flags = tracker.match_client(c, clients, app.config["CLIENT_ALIASES"]) if c else ("", [])
            name = name or c
            if qty is not None and qty > 0 and (r, name) in merged:
                merged[(r, name)].qty += qty  # same client twice under one rider -> one line
                continue
            row = extract.DraftRow(r, name, qty if qty is not None else q, flags)
            rows.append(row)
            if qty is not None and qty > 0:
                merged[(r, name)] = row

        if not rows:
            errors.append("Add at least one row.")
        if errors:
            for e in errors:
                flash(e, "error")
            return render_review(report_date, token, rows, [], status=400)

        sales = [Sale(r.rider, report_date, r.client, r.qty) for r in rows]
        try:
            tracker.append_sales(tp(), sales, Path(app.config["BACKUPS_DIR"]))
        except tracker.TrackerError as exc:
            flash(str(exc), "error")
            return render_review(report_date, token, rows, [], status=409)

        report = build_report(report_date, sales, riders())
        save_report(report, Path(app.config["REPORTS_DIR"]))
        if token:
            src = Path(app.config["UPLOADS_DIR"]) / f"{token}.jpg"
            if src.exists():
                shutil.move(src, Path(app.config["RAW_DIR"]) / f"{report_date.isoformat()}.jpg")

        new_clients = sorted({r.client for r in rows if any("not on the Client List" in f for f in r.flags)})
        if new_clients:
            flash("Not on the Client List yet: " + ", ".join(new_clients)
                  + ". Add them on the Clients page so they show up in the dashboard.", "warning")
        flash(f"Saved {len(sales)} row{'s' if len(sales) != 1 else ''} for {fmt_long(report_date)} to Tracker.xlsx.", "success")
        return redirect(url_for("report_view", day=report_date.isoformat()))

    # ------------------------------------------------------------------ reports
    def report_for(day: str):
        d = tracker.parse_date(day)
        if d is None:
            abort(404)
        sales = [s for s in tracker.load_sales(tp()) if s.date == d]
        if not sales:
            abort(404)
        return build_report(d, sales, riders())

    @app.get("/reports")
    def reports():
        sales = tracker.load_sales(tp())
        by_day: dict[date, list[Sale]] = {}
        for s in sales:
            by_day.setdefault(s.date, []).append(s)
        days = [
            {"date": d, "rows": len(v), "total": sum(s.count for s in v),
             "photo": (Path(app.config["RAW_DIR"]) / f"{d.isoformat()}.jpg").exists()}
            for d, v in sorted(by_day.items(), reverse=True)
        ]
        return render_template("reports.html", days=days)

    @app.get("/reports/<day>")
    def report_view(day):
        report = report_for(day)
        photo = (Path(app.config["RAW_DIR"]) / f"{report.report_date.isoformat()}.jpg").exists()
        return render_template("report.html", report=report, photo=photo)

    @app.get("/reports/<day>/download")
    def report_download(day):
        report = report_for(day)
        return Response(
            report_bytes(report),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{report.filename}"'},
        )

    @app.get("/reports/<day>/photo")
    def report_photo(day):
        d = tracker.parse_date(day)
        path = Path(app.config["RAW_DIR"]) / f"{d.isoformat()}.jpg" if d else None
        if not path or not path.exists():
            abort(404)
        return send_file(path, mimetype="image/jpeg")

    @app.post("/reports/<day>/delete")
    def report_delete(day):
        d = tracker.parse_date(day)
        if d is None:
            abort(404)
        n = tracker.delete_date(tp(), d, Path(app.config["BACKUPS_DIR"]))
        flash(f"Removed {n} rows for {fmt_long(d)} from Tracker.xlsx (a backup was kept).", "success")
        return redirect(url_for("reports"))

    # ------------------------------------------------------------------ dashboard
    @app.get("/dashboard")
    def dashboard_view():
        sales = tracker.load_sales(tp())
        today_ = today()
        month_param = request.args.get("month")  # YYYY-MM
        try:
            year, month = (int(x) for x in month_param.split("-")) if month_param else (None, None)
            date(year, month, 1)
        except (TypeError, ValueError):
            latest = max((s.date for s in sales), default=today_)
            year, month = latest.year, latest.month
        as_of = tracker.parse_date(request.args.get("as_of")) or today_
        days_n = app.config["INACTIVE_DAYS"]

        perf = dashboard.rider_performance(sales, year, month, riders())
        daily = dashboard.daily_sales(sales, year, month)
        matrix = dashboard.rider_daily_matrix(sales, year, month, riders())
        gaps = dashboard.clients_not_ordering(sales, tracker.load_clients(tp()), as_of, days_n)

        month_total = sum(q for _, q in daily)
        selling_days = [q for _, q in daily if q > 0]
        best = max(daily, key=lambda dq: dq[1]) if selling_days else None
        months = sorted({(s.date.year, s.date.month) for s in sales} | {(year, month)}, reverse=True)
        return render_template(
            "dashboard.html",
            year=year, month=month,
            month_label=date(year, month, 1).strftime("%B %Y"),
            months=[(f"{y}-{m:02d}", date(y, m, 1).strftime("%B %Y")) for y, m in months],
            perf=perf, daily=daily, matrix=matrix, gaps=gaps,
            as_of=as_of, days_n=days_n,
            window_start=as_of - timedelta(days=days_n - 1),
            month_total=month_total,
            selling_days=len(selling_days),
            avg_day=(month_total / len(selling_days)) if selling_days else 0,
            best=best,
            chart={
                "labels": [d.day for d, _ in daily],
                "totals": [q for _, q in daily],
                "riders": [p.rider for p in perf],
                "riderTotals": [p.total for p in perf],
                "matrix": matrix,
            },
        )

    # ------------------------------------------------------------------ clients
    @app.get("/clients")
    def clients_view():
        sales = tracker.load_sales(tp())
        last: dict[str, date] = {}
        for s in sales:
            k = s.client.upper()
            if k not in last or s.date > last[k]:
                last[k] = s.date
        clients = tracker.load_clients(tp())
        listed = {c.name.upper() for c in clients}
        unlisted = sorted(k for k in last if k not in listed)
        return render_template("clients.html", clients=clients, last=last, unlisted=unlisted)

    @app.post("/clients")
    def clients_post():
        try:
            tracker.upsert_client(tp(), request.form.get("name", ""), request.form.get("status", "ACTIVE"),
                                  Path(app.config["BACKUPS_DIR"]))
            flash(f"Saved client {request.form.get('name', '').strip().upper()}.", "success")
        except tracker.TrackerError as exc:
            flash(str(exc), "error")
        return redirect(url_for("clients_view"))

    # ------------------------------------------------------------------ tracker file
    @app.get("/tracker")
    def tracker_view():
        sales = tracker.load_sales(tp())
        return render_template(
            "tracker.html",
            rows=len(sales),
            days=len({s.date for s in sales}),
            first=min((s.date for s in sales), default=None),
            last=max((s.date for s in sales), default=None),
            clients=len(tracker.load_clients(tp())),
            path=tp(),
            modified=datetime.fromtimestamp(tp().stat().st_mtime),
        )

    @app.get("/tracker/download")
    def tracker_download():
        return send_file(tp(), as_attachment=True, download_name="Tracker.xlsx")

    @app.post("/tracker/upload")
    def tracker_upload():
        f = request.files.get("tracker")
        if not f or not f.filename.lower().endswith(".xlsx"):
            flash("Choose your Tracker.xlsx file.", "error")
            return redirect(url_for("tracker_view"))
        tmp = Path(app.config["UPLOADS_DIR"]) / f"tracker-{secrets.token_hex(6)}.xlsx"
        f.save(tmp)
        try:
            tracker.validate_workbook(tmp)
        except tracker.TrackerError as exc:
            tmp.unlink(missing_ok=True)
            flash(str(exc), "error")
            return redirect(url_for("tracker_view"))
        tracker.replace_tracker(tp(), tmp, Path(app.config["BACKUPS_DIR"]))
        flash("Tracker.xlsx replaced (the previous file was backed up).", "success")
        return redirect(url_for("tracker_view"))

    @app.errorhandler(413)
    def too_large(_):
        flash(f"That file is too big (limit {config.MAX_UPLOAD_MB} MB).", "error")
        return redirect(url_for("upload"))

    return app


if __name__ == "__main__":
    config.ensure_dirs()
    create_app().run(debug=True, host="0.0.0.0", port=5000)
