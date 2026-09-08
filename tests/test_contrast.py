"""Tests for WCAG 2.2 contrast ratio calculations."""
from pptx_a11y.contrast import calculate_contrast_ratio, is_contrast_compliant


def test_contrast_ratios():
    # Black on White: 21:1
    assert calculate_contrast_ratio("000000", "FFFFFF") == 21.0
    # White on White: 1:1
    assert calculate_contrast_ratio("FFFFFF", "FFFFFF") == 1.0

    # Low contrast gray on white (#888888 on #FFFFFF ~ 3.54:1)
    ratio = calculate_contrast_ratio("888888", "FFFFFF")
    assert ratio < 4.5
    assert not is_contrast_compliant(ratio, is_large_text=False)
    assert is_contrast_compliant(ratio, is_large_text=True)
