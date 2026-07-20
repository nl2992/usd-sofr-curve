"""Common ``ValidationReport`` dataclass + pretty-printer shared by the validation ladder.

None of the ladder functions in ``openusdcurve/validation/*.py`` raise on a failing check — they
always return a report; callers decide what a ``fail`` means for them (e.g. the CLI can exit
non-zero). This mirrors the "report, don't raise" philosophy already used by
``openusdcurve.data.quality.QualityReport``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

Status = Literal["pass", "warn", "fail", "skipped"]

__all__ = ["Status", "ValidationItem", "ValidationReport"]


@dataclass
class ValidationItem:
    """A single named check result within a :class:`ValidationReport`."""

    name: str
    status: Status
    message: str
    details: dict = field(default_factory=dict)


@dataclass
class ValidationReport:
    """An ordered collection of :class:`ValidationItem` plus an overall status."""

    label: str
    items: list[ValidationItem] = field(default_factory=list)
    skipped: bool = False

    @property
    def status(self) -> Status:
        if self.skipped:
            return "skipped"
        statuses = {i.status for i in self.items}
        if "fail" in statuses:
            return "fail"
        if "warn" in statuses:
            return "warn"
        if not statuses:
            return "pass"
        return "pass"

    def failing(self) -> list[ValidationItem]:
        return [i for i in self.items if i.status == "fail"]

    def warnings(self) -> list[ValidationItem]:
        return [i for i in self.items if i.status == "warn"]

    def to_text(self) -> str:
        lines = [f"=== {self.label} ===", f"status: {self.status}"]
        if self.skipped:
            lines.append("(skipped)")
        for i in self.items:
            lines.append(f"  [{i.status.upper():7s}] {i.name}: {i.message}")
        return "\n".join(lines)
