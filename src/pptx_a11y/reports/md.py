"""Canonical Markdown report renderer.

Delegates to engine_a11y.reports.md with pptx document profile.
"""
from typing import Any, Dict, Optional, Set
from engine_a11y.profile import get_pptx_profile
from engine_a11y.reports.md import render_md as _engine_render_md


def render_md(
    result: Dict[str, Any],
    after_result: Optional[Dict[str, Any]] = None,
    source_path: Optional[str] = None,
    profile: Optional[Any] = None,
    excluded_sc: Optional[Set[str]] = None,
) -> str:
    """Render canonical Markdown report using engine-a11y with pptx profile."""
    active_profile = profile or get_pptx_profile()
    md = _engine_render_md(
        result=result,
        after_result=after_result,
        source_path=source_path,
        profile=active_profile,
        excluded_sc=excluded_sc,
    )
    return md.replace("WCAG21/Understanding", "WCAG22/Understanding")


__all__ = ["render_md"]
