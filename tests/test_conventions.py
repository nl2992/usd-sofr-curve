"""Unit tests for day counts, the US business calendar, tenor arithmetic, and schedules."""

from __future__ import annotations

from datetime import date

import pytest

from openusdcurve.instruments.conventions import (
    USBusinessCalendar,
    act360,
    act365f,
    add_tenor,
    generate_schedule,
    parse_tenor,
    shift_months,
    thirty360,
    us_federal_holidays,
    year_fraction,
)


# ---------------------------------------------------------------------------
# Day counts
# ---------------------------------------------------------------------------


def test_act360_half_year():
    d0, d1 = date(2025, 1, 1), date(2025, 7, 1)
    assert act360(d0, d1) == pytest.approx((d1 - d0).days / 360.0)
    # 181 days / 360
    assert act360(d0, d1) == pytest.approx(181 / 360.0)


def test_act365f_full_year():
    assert act365f(date(2025, 1, 1), date(2026, 1, 1)) == pytest.approx(365 / 365.0)


def test_thirty360_simple_half_year():
    # Jan 1 -> Jul 1 is exactly 6 months of 30 days each.
    assert thirty360(date(2025, 1, 1), date(2025, 7, 1)) == pytest.approx(0.5)


def test_thirty360_full_year():
    assert thirty360(date(2025, 1, 15), date(2026, 1, 15)) == pytest.approx(1.0)


def test_thirty360_end_of_month_adjustment():
    # day0 = 31 -> treated as 30; day1 = 31 with day0=30 -> treated as 30.
    assert thirty360(date(2025, 1, 31), date(2025, 3, 31)) == pytest.approx(60 / 360.0)


def test_thirty360_february_end():
    # Last day of Feb -> day0 becomes 30, so Feb28 -> Aug 28 counts as 6*30.
    assert thirty360(date(2025, 2, 28), date(2025, 8, 28)) == pytest.approx(180 / 360.0)


def test_year_fraction_dispatch():
    d0, d1 = date(2025, 1, 1), date(2025, 7, 1)
    assert year_fraction(d0, d1, "act360") == act360(d0, d1)
    assert year_fraction(d0, d1, "act365f") == act365f(d0, d1)
    assert year_fraction(d0, d1, "thirty360") == thirty360(d0, d1)
    with pytest.raises(ValueError):
        year_fraction(d0, d1, "nonsense")


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------


def test_weekend_is_not_business_day():
    cal = USBusinessCalendar()
    assert not cal.is_business_day(date(2025, 5, 31))  # Saturday
    assert not cal.is_business_day(date(2025, 6, 1))  # Sunday
    assert cal.is_business_day(date(2025, 6, 2))  # Monday


def test_independence_day_observed():
    # July 4 2025 is a Friday -> holiday.
    assert date(2025, 7, 4) in us_federal_holidays(2025)
    cal = USBusinessCalendar()
    assert not cal.is_business_day(date(2025, 7, 4))


def test_new_year_observed_on_monday_when_saturday():
    # Jan 1 2022 was a Saturday -> observed Friday Dec 31 2021.
    hols_2021 = us_federal_holidays(2021)
    assert date(2021, 12, 31) in hols_2021


def test_following_rolls_forward():
    cal = USBusinessCalendar()
    # 2025-07-05 is Saturday -> following is Monday 2025-07-07.
    assert cal.following(date(2025, 7, 5)) == date(2025, 7, 7)


def test_modified_following_stays_in_month():
    cal = USBusinessCalendar()
    # 2025-05-31 Saturday -> following would be Mon Jun 2 (next month) -> mod-following = Fri May 30.
    assert cal.modified_following(date(2025, 5, 31)) == date(2025, 5, 30)


def test_add_business_days():
    cal = USBusinessCalendar()
    # From Fri 2025-07-03? 2025-07-03 is Thu. +1 bd skips July 4 holiday (Fri) and weekend.
    assert cal.add_business_days(date(2025, 7, 3), 1) == date(2025, 7, 7)


# ---------------------------------------------------------------------------
# Tenor arithmetic
# ---------------------------------------------------------------------------


def test_parse_tenor():
    assert parse_tenor("3M") == (3, "M")
    assert parse_tenor("1y") == (1, "Y")
    assert parse_tenor("2W") == (2, "W")
    assert parse_tenor("10D") == (10, "D")
    with pytest.raises(ValueError):
        parse_tenor("3Q")


def test_add_tenor_month_end_clamp():
    # Jan 31 + 1M -> Feb 28 2025 (Friday, business day, no roll needed).
    assert add_tenor(date(2025, 1, 31), "1M") == date(2025, 2, 28)


def test_add_tenor_one_year():
    # 2025-01-02 + 1Y -> 2026-01-02 (Friday).
    assert add_tenor(date(2025, 1, 2), "1Y") == date(2026, 1, 2)


def test_add_tenor_rolls_to_business_day():
    # Weekend landing NOT at a month boundary, so modified-following == following here.
    # 2025-05-30 Fri + 2D = 2025-06-01 Sun -> roll forward to Mon 2025-06-02 (still in June).
    result = add_tenor(date(2025, 5, 30), "2D")
    assert result == date(2025, 6, 2)
    # And an explicit 'following' roll over a month boundary rolls into the next month.
    assert add_tenor(date(2025, 5, 30), "1D", roll="following") == date(2025, 6, 2)


def test_shift_months():
    assert shift_months(date(2025, 6, 15), -3) == date(2025, 3, 15)
    assert shift_months(date(2025, 12, 31), 2) == date(2026, 2, 28)


# ---------------------------------------------------------------------------
# Schedules
# ---------------------------------------------------------------------------


def test_generate_schedule_semiannual():
    start = date(2025, 1, 2)
    end = add_tenor(start, "2Y")
    sched = generate_schedule(start, end, "6M")
    assert len(sched) == 4  # 6M, 12M, 18M, 24M
    assert sched[-1] == end
    # Strictly increasing.
    assert all(sched[i] < sched[i + 1] for i in range(len(sched) - 1))
    # Every date is a business day.
    cal = USBusinessCalendar()
    assert all(cal.is_business_day(d) for d in sched)


def test_generate_schedule_annual():
    start = date(2025, 1, 2)
    end = add_tenor(start, "3Y")
    sched = generate_schedule(start, end, "1Y")
    assert len(sched) == 3
    assert sched[-1] == end
