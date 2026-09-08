"""Color contrast calculations adhering to WCAG 2.2 Success Criterion 1.4.3."""
from typing import Tuple
import wcag_contrast_ratio as contrast


def hex_to_rgb(hex_str: str) -> Tuple[float, float, float]:
    """Convert 6-char hex string to normalized RGB floats 0.0-1.0."""
    clean_hex = hex_str.lstrip("#")
    if len(clean_hex) != 6:
        return (0.0, 0.0, 0.0)
    try:
        r = int(clean_hex[0:2], 16) / 255.0
        g = int(clean_hex[2:4], 16) / 255.0
        b = int(clean_hex[4:6], 16) / 255.0
        return (r, g, b)
    except ValueError:
        return (0.0, 0.0, 0.0)


def calculate_contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """Compute WCAG 2.2 contrast ratio between two hex colors."""
    fg_rgb = hex_to_rgb(fg_hex)
    bg_rgb = hex_to_rgb(bg_hex)
    return round(contrast.rgb(fg_rgb, bg_rgb), 2)


def is_contrast_compliant(ratio: float, is_large_text: bool = False) -> bool:
    """Return True if contrast meets WCAG 2.2 AA (4.5:1 standard, 3:1 large)."""
    required = 3.0 if is_large_text else 4.5
    return ratio >= required
