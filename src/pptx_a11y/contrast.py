"""Color contrast calculations adhering to WCAG 2.2 Success Criterion 1.4.3.

Delegates core contrast calculations to engine_a11y.contrast.
"""
from typing import Tuple
from engine_a11y.contrast import (
    THRESHOLD_LARGE,
    THRESHOLD_NORMAL,
    contrast_ratio,
    hex_to_rgb as _engine_hex_to_rgb,
)


def hex_to_rgb(hex_str: str) -> Tuple[float, float, float]:
    """Convert 6-char hex string to normalized RGB floats 0.0-1.0."""
    rgb = _engine_hex_to_rgb(hex_str)
    if not rgb:
        return (0.0, 0.0, 0.0)
    return (rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)


def calculate_contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """Compute WCAG 2.2 contrast ratio between two hex colors."""
    fg = _engine_hex_to_rgb(fg_hex) or (0, 0, 0)
    bg = _engine_hex_to_rgb(bg_hex) or (255, 255, 255)
    return round(contrast_ratio(fg, bg), 2)


def is_contrast_compliant(ratio: float, is_large_text: bool = False) -> bool:
    """Return True if contrast meets WCAG 2.2 AA (4.5:1 standard, 3:1 large)."""
    required = THRESHOLD_LARGE if is_large_text else THRESHOLD_NORMAL
    return ratio >= required


__all__ = [
    "hex_to_rgb",
    "calculate_contrast_ratio",
    "is_contrast_compliant",
    "THRESHOLD_NORMAL",
    "THRESHOLD_LARGE",
]
