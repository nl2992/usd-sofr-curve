"""Reproducibility helpers (PLAN §8).

Persists raw provider responses under ``data/raw/<source>/<retrieval_ts>/`` (one SHA-256 per raw
input, recorded in a small JSON manifest alongside the raw payload) and normalized quote sets
under ``data/normalized/<valuation_date>/``. Nothing here performs network I/O — callers pass in
already-fetched bytes/text and already-built ``MarketQuote`` lists.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from datetime import date, datetime, timezone
from pathlib import Path

from openusdcurve.data.base import MarketQuote

__all__ = [
    "compute_sha256",
    "save_raw",
    "save_normalized",
    "load_normalized",
]


def compute_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _retrieval_timestamp(now: datetime | None = None) -> str:
    now = now or datetime.now(timezone.utc)
    return now.strftime("%Y%m%dT%H%M%SZ")


def save_raw(
    source: str,
    raw: bytes | str,
    *,
    base_dir: str | Path = "data/raw",
    filename: str = "raw.txt",
    retrieval_ts: str | None = None,
    extra_metadata: dict | None = None,
) -> dict:
    """Write a raw provider response to ``<base_dir>/<source>/<retrieval_ts>/<filename>`` plus a
    ``manifest.json`` with its SHA-256. Returns a dict with ``path``, ``sha256``, ``retrieval_ts``.
    """
    if isinstance(raw, str):
        raw_bytes = raw.encode("utf-8")
    else:
        raw_bytes = raw

    ts = retrieval_ts or _retrieval_timestamp()
    out_dir = Path(base_dir) / source / ts
    out_dir.mkdir(parents=True, exist_ok=True)

    raw_path = out_dir / filename
    raw_path.write_bytes(raw_bytes)

    digest = compute_sha256(raw_bytes)
    manifest = {
        "source": source,
        "retrieval_ts": ts,
        "filename": filename,
        "sha256": digest,
        "size_bytes": len(raw_bytes),
        **(extra_metadata or {}),
    }
    (out_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    return {"path": str(raw_path), "sha256": digest, "retrieval_ts": ts}


def _quote_to_jsonable(q: MarketQuote) -> dict:
    d = asdict(q)
    for key in ("valuation_date", "maturity_date", "start_date"):
        if d.get(key) is not None:
            d[key] = d[key].isoformat() if hasattr(d[key], "isoformat") else d[key]
    if d.get("observed_at") is not None:
        d["observed_at"] = d["observed_at"].isoformat()
    return d


def save_normalized(
    quotes: list[MarketQuote],
    valuation_date: date,
    *,
    base_dir: str | Path = "data/normalized",
    filename: str = "quotes.json",
) -> dict:
    """Write a normalized quote set to ``<base_dir>/<valuation_date>/<filename>``. Returns a dict
    with ``path`` and ``sha256`` of the serialized content."""
    out_dir = Path(base_dir) / valuation_date.isoformat()
    out_dir.mkdir(parents=True, exist_ok=True)

    payload = [_quote_to_jsonable(q) for q in quotes]
    content = json.dumps(payload, indent=2, sort_keys=True)
    content_bytes = content.encode("utf-8")

    out_path = out_dir / filename
    out_path.write_bytes(content_bytes)

    digest = compute_sha256(content_bytes)
    (out_dir / f"{filename}.sha256").write_text(digest, encoding="utf-8")

    return {"path": str(out_path), "sha256": digest}


def load_normalized(path: str | Path) -> list[dict]:
    """Read back a normalized quote JSON file as a list of plain dicts (not MarketQuote objects —
    callers reconstruct MarketQuote themselves if needed, keeping this module dependency-light)."""
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)
