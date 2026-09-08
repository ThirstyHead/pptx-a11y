"""Reading order analysis and spatial sorting for PowerPoint slides (WCAG 1.3.2)."""
from typing import List, Tuple
from pptx.slide import Slide
from pptx.shapes.base import BaseShape
from pptx.util import Inches

_VERTICAL_TOLERANCE = Inches(0.25)


def get_positionable_shapes(slide: Slide) -> List[BaseShape]:
    """Return all content shapes with valid spatial coordinates."""
    shapes: List[BaseShape] = []
    for s in slide.shapes:
        if hasattr(s, "top") and hasattr(s, "left") and s.top is not None and s.left is not None:
            shapes.append(s)
    return shapes


def check_slide_reading_order(slide: Slide) -> bool:
    """Return True if shapes in the shape tree are spatially inverted (out of reading order)."""
    shapes = get_positionable_shapes(slide)
    if len(shapes) < 2:
        return False

    for i in range(len(shapes) - 1):
        s_current = shapes[i]
        s_next = shapes[i + 1]

        # If s_next is visually significantly above s_current, order is inverted
        if (s_current.top - s_next.top) > _VERTICAL_TOLERANCE:
            return True
    return False


def reorder_slide_shapes(slide: Slide) -> int:
    """Sort shapes in p:spTree by spatial coordinates (top, left). Returns number of shapes reordered."""
    sp_tree = slide.shapes._spTree
    shapes = get_positionable_shapes(slide)
    if len(shapes) < 2:
        return 0

    # Sort shapes primarily by top (within vertical band of 0.25 in), then by left
    def sort_key(s: BaseShape) -> Tuple[int, int]:
        band = int(s.top // _VERTICAL_TOLERANCE)
        return (band, int(s.left))

    sorted_shapes = sorted(shapes, key=sort_key)

    # Re-append shape XML elements in sorted order to the spTree container
    for shape in sorted_shapes:
        elem = shape._element
        sp_tree.append(elem)

    return len(sorted_shapes)
