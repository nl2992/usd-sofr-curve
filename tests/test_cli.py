"""Smoke tests for openusdcurve/cli.py (no pricing/validation dependency).

These exercise ``main()`` directly (not a subprocess) against configs that are fully offline
(synthetic/CSV sources), per the task spec: do not depend on the pricing/validation layers here.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from openusdcurve.cli import main

_REPO_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture(autouse=True)
def _run_from_repo_root(monkeypatch):
    monkeypatch.chdir(_REPO_ROOT)


def test_build_sofr_futures_public_runs_without_error(capsys):
    exit_code = main(
        [
            "build",
            "--config",
            "configs/sofr_futures_public.yaml",
            "--date",
            "2026-07-20",
            "--no-save",
        ]
    )
    assert exit_code == 0
    out = capsys.readouterr().out
    assert "USD-SOFR-FUTURES-PUBLIC" in out
    assert "pillar_date" in out


def test_build_lehman_public_2002_uses_config_default_date(capsys):
    exit_code = main(
        ["build", "--config", "configs/lehman_public_2002.yaml", "--no-save"]
    )
    assert exit_code == 0
    out = capsys.readouterr().out
    assert "2002-08-26" in out


def test_build_saves_curve_json(tmp_path, capsys):
    out_dir = tmp_path / "curves"
    exit_code = main(
        [
            "build",
            "--config",
            "configs/sofr_futures_public.yaml",
            "--date",
            "2026-07-20",
            "--out-dir",
            str(out_dir),
        ]
    )
    assert exit_code == 0
    saved = out_dir / "USD-SOFR-FUTURES-PUBLIC" / "2026-07-20" / "curve.json"
    assert saved.is_file()
    payload = json.loads(saved.read_text())
    assert payload["curve_id"] == "USD-SOFR-FUTURES-PUBLIC"
    assert payload["pillars"]
    for pillar in payload["pillars"]:
        assert pillar["discount_factor"] > 0.0


def test_build_missing_config_exits_nonzero(capsys):
    exit_code = main(["build", "--config", "configs/does_not_exist.yaml", "--date", "2026-07-20"])
    assert exit_code != 0
    err = capsys.readouterr().err
    assert "not found" in err


def test_validate_reports_layer_unavailable_or_runs(capsys):
    """The validation package may or may not exist yet (built concurrently); either way this
    must exit 0 and print something meaningful."""
    exit_code = main(
        ["validate", "--curve", "USD-SOFR-FUTURES-PUBLIC", "--date", "2026-07-20", "--offline"]
    )
    assert exit_code == 0
    out = capsys.readouterr().out
    assert out.strip()


def test_data_pull_offline_new_york_fed(tmp_path, capsys):
    exit_code = main(
        [
            "data",
            "pull",
            "--source",
            "new-york-fed",
            "--date",
            "2026-07-20",
            "--offline",
            "--raw-dir",
            str(tmp_path / "raw"),
            "--normalized-dir",
            str(tmp_path / "normalized"),
        ]
    )
    assert exit_code == 0
    out = capsys.readouterr().out
    assert "used bundled sample: True" in out


def test_data_pull_offline_treasury(tmp_path, capsys):
    exit_code = main(
        [
            "data",
            "pull",
            "--source",
            "treasury",
            "--date",
            "2026-07-20",
            "--offline",
            "--raw-dir",
            str(tmp_path / "raw"),
            "--normalized-dir",
            str(tmp_path / "normalized"),
        ]
    )
    assert exit_code == 0


def test_data_pull_offline_fred(tmp_path, capsys):
    exit_code = main(
        [
            "data",
            "pull",
            "--source",
            "fred",
            "--date",
            "2026-07-20",
            "--offline",
            "--raw-dir",
            str(tmp_path / "raw"),
            "--normalized-dir",
            str(tmp_path / "normalized"),
        ]
    )
    assert exit_code == 0
