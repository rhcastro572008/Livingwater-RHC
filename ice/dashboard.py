"""Numbers for the dashboard."""
from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, timedelta

from .tracker import Client, Sale


@dataclass
class RiderStat:
    rider: str
    total: float
    days_with_sales: int
    deliveries: int
    clients: int
    best_day: date | None
    best_day_qty: float
    share: float  # 0..1 of the month's total

    @property
    def avg_per_day(self) -> float:
        return self.total / self.days_with_sales if self.days_with_sales else 0.0


@dataclass
class ClientGap:
    client: str
    status: str
    last_order: date | None
    days_since: int | None
    last_rider: str | None
    total_month: float


def month_bounds(year: int, month: int) -> tuple[date, date]:
    last = calendar.monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last)


def rider_performance(sales: list[Sale], year: int, month: int, riders: list[str]) -> list[RiderStat]:
    start, end = month_bounds(year, month)
    month_sales = [s for s in sales if start <= s.date <= end]
    grand = sum(s.count for s in month_sales)
    names = list(riders) + sorted({s.rider for s in month_sales} - set(riders))
    stats = []
    for rider in names:
        rs = [s for s in month_sales if s.rider == rider]
        per_day: dict[date, float] = {}
        for s in rs:
            per_day[s.date] = per_day.get(s.date, 0) + s.count
        best = max(per_day.items(), key=lambda kv: kv[1]) if per_day else (None, 0.0)
        total = sum(s.count for s in rs)
        stats.append(RiderStat(
            rider=rider,
            total=total,
            days_with_sales=sum(1 for q in per_day.values() if q > 0),
            deliveries=len(rs),
            clients=len({s.client.upper() for s in rs}),
            best_day=best[0],
            best_day_qty=best[1],
            share=(total / grand) if grand else 0.0,
        ))
    stats.sort(key=lambda s: s.total, reverse=True)
    return stats


def daily_sales(sales: list[Sale], year: int, month: int) -> list[tuple[date, float]]:
    """Total quantity per calendar day of the month (all riders), 0 for days without sales."""
    start, end = month_bounds(year, month)
    totals: dict[date, float] = {}
    for s in sales:
        if start <= s.date <= end:
            totals[s.date] = totals.get(s.date, 0) + s.count
    days = []
    d = start
    while d <= end:
        days.append((d, totals.get(d, 0.0)))
        d += timedelta(days=1)
    return days


def rider_daily_matrix(sales: list[Sale], year: int, month: int, riders: list[str]) -> dict[str, list[float]]:
    """rider -> list of daily quantities, aligned with daily_sales()."""
    start, end = month_bounds(year, month)
    n = (end - start).days + 1
    names = list(riders) + sorted({s.rider for s in sales if start <= s.date <= end} - set(riders))
    matrix = {r: [0.0] * n for r in names}
    for s in sales:
        if start <= s.date <= end:
            matrix[s.rider][(s.date - start).days] += s.count
    return matrix


def clients_not_ordering(
    sales: list[Sale],
    clients: list[Client],
    as_of: date,
    days: int = 5,
    include_inactive: bool = False,
) -> list[ClientGap]:
    """Clients with no order in the `days` days up to and including `as_of`.

    Uses the Client List when there is one (ACTIVE and blank-status clients; INACTIVE
    ones only if include_inactive), and also any client that appears in the sales data.
    """
    window_start = as_of - timedelta(days=days - 1)
    month_start = as_of.replace(day=1)

    last: dict[str, Sale] = {}
    month_total: dict[str, float] = {}
    display: dict[str, str] = {}
    for s in sales:
        if s.date > as_of:
            continue
        k = s.client.upper()
        display.setdefault(k, s.client)
        if k not in last or s.date >= last[k].date:
            last[k] = s
        if s.date >= month_start:
            month_total[k] = month_total.get(k, 0) + s.count

    status = {c.name.upper(): c.status for c in clients}
    for c in clients:
        display[c.name.upper()] = c.name

    gaps = []
    for k, name in display.items():
        st = status.get(k, "")
        if st == "INACTIVE" and not include_inactive:
            continue
        sale = last.get(k)
        if sale and sale.date >= window_start:
            continue  # ordered recently
        gaps.append(ClientGap(
            client=name,
            status=st or ("" if k in status else "not on list"),
            last_order=sale.date if sale else None,
            days_since=(as_of - sale.date).days if sale else None,
            last_rider=sale.rider if sale else None,
            total_month=month_total.get(k, 0.0),
        ))
    # Most recently lost clients first (they are the easiest to win back); never-ordered last.
    gaps.sort(key=lambda g: (g.days_since is None, g.days_since or 0, g.client))
    return gaps
