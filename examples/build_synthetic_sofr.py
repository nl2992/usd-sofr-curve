#!/usr/bin/env python
"""Build a curve end-to-end from the synthetic source and print zero/forward/par rates.

Runs fully offline (no network, no bundled sample files) — ``SyntheticSource`` generates a
self-consistent quote set from a closed-form Nelson-Siegel reference curve (docs/PLAN.md §6,
dataset OS-1). Run with:

    python examples/build_synthetic_sofr.py
"""

from __future__ import annotations

from datetime import date

from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.data.synthetic import SyntheticSource
from openusdcurve.instruments.factory import build_instruments


def _dedupe_by_pillar_date(instruments: list) -> list:
    """Keep only the first instrument seen at each distinct pillar_date.

    ``SyntheticSource``'s default tenor grids deliberately overlap at a couple of points (e.g.
    the 12M deposit and the OIS 1Y both land on the same date) since each instrument type is
    generated independently. ``DiscountCurve`` itself only keeps one node per distinct pillar
    time, so a second instrument pinned to an already-solved date has no effect on the trial
    curve and the bootstrap cannot bracket a root for it. Dropping duplicates up front (keeping
    the first-generated instrument per date, matching DiscountCurve's own dedup rule) keeps this
    end-to-end demo well-posed.
    """
    seen: set = set()
    out = []
    for inst in instruments:
        if inst.pillar_date in seen:
            continue
        seen.add(inst.pillar_date)
        out.append(inst)
    return out


def main() -> None:
    valuation_date = date(2026, 7, 20)

    source = SyntheticSource()
    quotes = source.get_quotes(valuation_date)
    print(f"Generated {len(quotes)} synthetic quotes for {valuation_date}.")

    instruments = _dedupe_by_pillar_date(build_instruments(quotes, valuation_date))
    curve = bootstrap(instruments, valuation_date)
    print(f"Bootstrapped {len(instruments)} calibration instruments into a discount curve.\n")

    print(f"{'pillar_date':12s} {'instrument_id':12s} {'discount':>12s} {'zero_rate':>11s}")
    ordered = sorted(instruments, key=lambda i: i.pillar_date)
    for inst in ordered:
        df = curve.discount(inst.pillar_date)
        z = curve.zero_rate(inst.pillar_date)
        print(f"{inst.pillar_date.isoformat():12s} {inst.instrument_id:12s} {df:12.8f} {z:10.4%}")

    print("\nForward rates between consecutive pillars:")
    for prev, nxt in zip(ordered, ordered[1:], strict=False):
        fwd = curve.forward_rate(prev.pillar_date, nxt.pillar_date)
        print(f"  {prev.pillar_date} -> {nxt.pillar_date}: {fwd:.4%}")

    print("\nPar rates (target quotes the curve was calibrated to reprice):")
    for inst in ordered:
        print(f"  {inst.instrument_id:12s} target_quote={inst.target_quote:.6f}")


if __name__ == "__main__":
    main()
