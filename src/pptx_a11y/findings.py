"""Finding data structures, sorting, and summary metrics.

Delegates core finding structures to engine_a11y.findings.
"""
from engine_a11y.findings import (
    Finding,
    Severity,
    SEVERITY_ORDER,
    findings_sorted,
    summarize,
)

__all__ = [
    "Finding",
    "Severity",
    "SEVERITY_ORDER",
    "findings_sorted",
    "summarize",
]
