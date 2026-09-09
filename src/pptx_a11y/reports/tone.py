"""Strict tone and language guards implementing the Social Model of Disability.

Delegates to engine_a11y.reports.tone.
"""
from engine_a11y.reports.tone import (
    BANNED_PHRASES,
    RULE_BARRIER_EXPLANATIONS,
    WHO_MAP,
    assert_social_model_language,
)


def who_benefits_for_rule(rule_id: str) -> str:
    return WHO_MAP.get(
        rule_id,
        "All users, particularly people navigating with assistive technologies.",
    )


def barrier_for_rule(rule_id: str) -> str:
    return RULE_BARRIER_EXPLANATIONS.get(
        rule_id,
        "Inaccessible content prevents users from reading or navigating.",
    )


__all__ = [
    "BANNED_PHRASES",
    "RULE_BARRIER_EXPLANATIONS",
    "WHO_MAP",
    "assert_social_model_language",
    "barrier_for_rule",
    "who_benefits_for_rule",
]
