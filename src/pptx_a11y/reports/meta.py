"""Canonical W3C WCAG 2.2 metadata, Understanding URLs, and POUR taxonomy.

Delegates to engine_a11y.reports.meta.
"""
from engine_a11y.reports.meta import (
    POUR_INTROS,
    PRINCIPLES,
    SC_META,
    W3C_QUICKREF,
    W3C_UNDERSTANDING_BASE,
)


def sc_understanding_url(sc: str) -> str:
    item = SC_META.get(sc)
    return item[3] if item else W3C_QUICKREF


def sc_label(sc: str) -> str:
    item = SC_META.get(sc)
    return item[0] if item else sc


def sc_level(sc: str) -> str:
    item = SC_META.get(sc)
    return item[1] if item else "AA"


def sc_principle(sc: str) -> str:
    item = SC_META.get(sc)
    return item[2] if item else "Unknown"


__all__ = [
    "W3C_UNDERSTANDING_BASE",
    "W3C_QUICKREF",
    "SC_META",
    "PRINCIPLES",
    "POUR_INTROS",
    "sc_understanding_url",
    "sc_label",
    "sc_level",
    "sc_principle",
]
