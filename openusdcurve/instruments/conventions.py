"""Dependency-free day counts, US business-day calendar, tenor arithmetic, and schedules.

Only stdlib + numpy are used here (no pandas / dateutil), per docs/PLAN.md §4.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

# ---------------------------------------------------------------------------
# Day counts
# ---------------------------------------------------------------------------


def act360(d0: date, d1: date) -> float:
    """Actual/360 year fraction."""
    return (d1 - d0).days / 360.0


def act365f(d0: date, d1: date) -> float:
    """Actual/365 fixed year fraction."""
    return (d1 - d0).days / 365.0


def thirty360(d0: date, d1: date) -> float:
    """30/360 (US / Bond Basis) year fraction.

    Matches QuantLib ``Thirty360(BondBasis)`` — the standard USD fixed-leg convention: apply
    only the 31->30 rules, with NO end-of-February adjustment (that belongs to the separate
    ``30/360 US with EOM`` variant, not to Bond Basis).
    """
    y0, m0, day0 = d0.year, d0.month, d0.day
    y1, m1, day1 = d1.year, d1.month, d1.day

    if day0 == 31:
        day0 = 30
    if day1 == 31 and day0 == 30:
        day1 = 30

    return ((y1 - y0) * 360 + (m1 - m0) * 30 + (day1 - day0)) / 360.0


DAY_COUNTS = {
    "act360": act360,
    "act365f": act365f,
    "thirty360": thirty360,
}


def year_fraction(d0: date, d1: date, convention: str = "act365f") -> float:
    """Dispatch to a named day-count convention."""
    try:
        fn = DAY_COUNTS[convention]
    except KeyError as exc:
        raise ValueError(f"Unknown day-count convention: {convention!r}") from exc
    return fn(d0, d1)


# ---------------------------------------------------------------------------
# US business-day calendar
# ---------------------------------------------------------------------------


def _nth_weekday_of_month(year: int, month: int, weekday: int, n: int) -> date:
    """The n-th (1-indexed) occurrence of ``weekday`` (Mon=0..Sun=6) in ``month``/``year``."""
    d = date(year, month, 1)
    offset = (weekday - d.weekday()) % 7
    d = d + timedelta(days=offset + 7 * (n - 1))
    return d


def _last_weekday_of_month(year: int, month: int, weekday: int) -> date:
    if month == 12:
        d = date(year, 12, 31)
    else:
        d = date(year, month + 1, 1) - timedelta(days=1)
    offset = (d.weekday() - weekday) % 7
    return d - timedelta(days=offset)


def _observed(d: date) -> date:
    """Federal 'in lieu of' observance: Sat -> Fri, Sun -> Mon."""
    if d.weekday() == 5:
        return d - timedelta(days=1)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def us_federal_holidays(year: int) -> set[date]:
    """Approximate set of US federal holidays observed in ``year`` (for bank-holiday rolls)."""
    holidays = {
        _observed(date(year, 1, 1)),  # New Year's Day
        _nth_weekday_of_month(year, 1, 0, 3),  # MLK Day: 3rd Monday of Jan
        _nth_weekday_of_month(year, 2, 0, 3),  # Presidents' Day: 3rd Monday of Feb
        _last_weekday_of_month(year, 5, 0),  # Memorial Day: last Monday of May
        _observed(date(year, 6, 19)),  # Juneteenth
        _observed(date(year, 7, 4)),  # Independence Day
        _nth_weekday_of_month(year, 9, 0, 1),  # Labor Day: 1st Monday of Sep
        _nth_weekday_of_month(year, 10, 0, 2),  # Columbus Day: 2nd Monday of Oct
        _observed(date(year, 11, 11)),  # Veterans Day
        _nth_weekday_of_month(year, 11, 3, 4),  # Thanksgiving: 4th Thursday of Nov
        _observed(date(year, 12, 25)),  # Christmas
    }
    # Cross-year spillover: if next year's New Year (Jan 1) falls on a Saturday, it is observed
    # on Friday Dec 31 of THIS year, so it must appear in this year's holiday set.
    if date(year + 1, 1, 1).weekday() == 5:
        holidays.add(date(year, 12, 31))
    return holidays


class USBusinessCalendar:
    """A simple US business-day calendar: weekends + approximate federal holidays."""

    def __init__(self) -> None:
        self._holiday_cache: dict[int, set[date]] = {}

    def _holidays_for_year(self, year: int) -> set[date]:
        if year not in self._holiday_cache:
            self._holiday_cache[year] = us_federal_holidays(year)
        return self._holiday_cache[year]

    def is_business_day(self, d: date) -> bool:
        if d.weekday() >= 5:
            return False
        return d not in self._holidays_for_year(d.year)

    def next_business_day(self, d: date) -> date:
        nxt = d + timedelta(days=1)
        while not self.is_business_day(nxt):
            nxt += timedelta(days=1)
        return nxt

    def previous_business_day(self, d: date) -> date:
        prv = d - timedelta(days=1)
        while not self.is_business_day(prv):
            prv -= timedelta(days=1)
        return prv

    def add_business_days(self, d: date, n: int) -> date:
        cur = d
        step = 1 if n >= 0 else -1
        for _ in range(abs(n)):
            cur = cur + timedelta(days=step)
            while not self.is_business_day(cur):
                cur = cur + timedelta(days=step)
        return cur

    def following(self, d: date) -> date:
        if self.is_business_day(d):
            return d
        return self.next_business_day(d)

    def modified_following(self, d: date) -> date:
        rolled = self.following(d)
        if rolled.month != d.month:
            return self.previous_business_day(d)
        return rolled

    def preceding(self, d: date) -> date:
        if self.is_business_day(d):
            return d
        return self.previous_business_day(d)

    def roll(self, d: date, convention: str = "modified_following") -> date:
        if convention == "following":
            return self.following(d)
        if convention == "modified_following":
            return self.modified_following(d)
        if convention == "preceding":
            return self.preceding(d)
        if convention == "unadjusted":
            return d
        raise ValueError(f"Unknown roll convention: {convention!r}")


DEFAULT_CALENDAR = USBusinessCalendar()


# ---------------------------------------------------------------------------
# Tenor arithmetic
# ---------------------------------------------------------------------------

_TENOR_RE = re.compile(r"^(\d+)\s*([DdWwMmYy])$")


def _add_months(d: date, months: int) -> date:
    total_month_index = (d.month - 1) + months
    year = d.year + total_month_index // 12
    month = total_month_index % 12 + 1
    # clamp day to last day of the resulting month
    if month == 12:
        next_month_first = date(year + 1, 1, 1)
    else:
        next_month_first = date(year, month + 1, 1)
    last_day = (next_month_first - timedelta(days=1)).day
    day = min(d.day, last_day)
    return date(year, month, day)


def shift_months(d: date, months: int) -> date:
    """Add (or subtract, if negative) a whole number of calendar months to ``d``, unadjusted."""
    return _add_months(d, months)


def parse_tenor(tenor: str) -> tuple[int, str]:
    """Parse a tenor string like '3M', '1Y', '2W', '1D' into (count, unit)."""
    m = _TENOR_RE.match(tenor.strip())
    if not m:
        raise ValueError(f"Invalid tenor string: {tenor!r}")
    count = int(m.group(1))
    unit = m.group(2).upper()
    return count, unit


def add_tenor(
    d: date,
    tenor: str,
    calendar: USBusinessCalendar | None = None,
    roll: str = "modified_following",
) -> date:
    """Add a tenor (e.g. '3M', '1Y', '1W', '2D') to ``d`` and roll to a business day.

    Day and week tenors add calendar days (then roll); month and year tenors add
    calendar months (then roll), matching standard swap-schedule conventions.
    """
    count, unit = parse_tenor(tenor)
    if unit == "D":
        raw = d + timedelta(days=count)
    elif unit == "W":
        raw = d + timedelta(weeks=count)
    elif unit == "M":
        raw = _add_months(d, count)
    elif unit == "Y":
        raw = _add_months(d, count * 12)
    else:  # pragma: no cover - guarded by parse_tenor
        raise ValueError(f"Unknown tenor unit: {unit!r}")

    cal = calendar or DEFAULT_CALENDAR
    return cal.roll(raw, roll)


# ---------------------------------------------------------------------------
# Schedule generation
# ---------------------------------------------------------------------------


def generate_schedule(
    start: date,
    end: date,
    frequency: str,
    calendar: USBusinessCalendar | None = None,
    roll: str = "modified_following",
) -> list[date]:
    """Generate a list of period-end dates from ``start`` to ``end`` stepping by ``frequency``.

    Dates are generated FORWARD from ``start`` on the unadjusted grid (``start + k*frequency``)
    and each is rolled to a business day. This anchors the schedule to ``start`` so a clean
    N-period swap yields exactly N period-ends with no spurious front stub (the failure mode of
    generating backward from an already-rolled ``end``). The returned list does NOT include
    ``start`` but always ends exactly at ``end``.
    """
    cal = calendar or DEFAULT_CALENDAR
    count, unit = parse_tenor(frequency)

    def _step(k: int) -> date:
        if unit == "D":
            return start + timedelta(days=count * k)
        if unit == "W":
            return start + timedelta(weeks=count * k)
        if unit == "M":
            return _add_months(start, count * k)
        if unit == "Y":
            return _add_months(start, count * 12 * k)
        raise ValueError(f"Unknown tenor unit: {unit!r}")  # pragma: no cover

    dates: list[date] = []
    k = 1
    while True:
        rolled = cal.roll(_step(k), roll)
        if rolled >= end:
            break
        dates.append(rolled)
        k += 1
    dates.append(end)
    return dates
