"""pptx-a11y reporting modules."""
from .md import render_md
from .meta import POUR_INTROS, PRINCIPLES, SC_META, W3C_QUICKREF, W3C_UNDERSTANDING_BASE
from .stats import compute_progress_stats
from .tone import BANNED_PHRASES, RULE_BARRIER_EXPLANATIONS, WHO_MAP, assert_social_model_language

__all__ = [
    "render_md",
    "SC_META",
    "PRINCIPLES",
    "POUR_INTROS",
    "W3C_QUICKREF",
    "W3C_UNDERSTANDING_BASE",
    "WHO_MAP",
    "RULE_BARRIER_EXPLANATIONS",
    "BANNED_PHRASES",
    "assert_social_model_language",
    "compute_progress_stats",
]
