"""Fuzzy name scoring for placement-time alias proposals (pure, no Revit).

When placement matching (strictly exact since the b1eefcf fix) leaves
target family names unmatched, the placement window uses this module to
find the profile each near-miss name most plausibly belongs to and
offers to record it as a ``merged_aliases`` entry.

Scoring is the max of two views of the pair, on a 0-100 scale:

* character-level ``SequenceMatcher`` ratio — catches suffix/typo
  near-misses like ``"LTG_2x4 Troffer_CED_2"`` vs
  ``"LTG_2x4 Troffer_CED"``;
* token-set ratio (adapted from the Audit Linked Model 2.0 tool) —
  catches reordered or re-delimited names. Family names are
  underscore/space/dash-delimited, so tokens split on any
  non-alphanumeric run.
"""

import re

from difflib import SequenceMatcher

DEFAULT_THRESHOLD = 80.0

_TOKEN_RE = re.compile(r"[^0-9a-z]+")


def _tokens(value):
    return set(t for t in _TOKEN_RE.split((value or "").lower()) if t)


def token_set_ratio(a, b):
    """0-100 similarity of the two names' token sets."""
    a_tokens = _tokens(a)
    b_tokens = _tokens(b)
    if not a_tokens or not b_tokens:
        return 0.0
    common = " ".join(sorted(a_tokens & b_tokens))
    a_diff = " ".join(sorted(a_tokens - b_tokens))
    b_diff = " ".join(sorted(b_tokens - a_tokens))
    return max(
        SequenceMatcher(None, common, (common + " " + a_diff).strip()).ratio(),
        SequenceMatcher(None, common, (common + " " + b_diff).strip()).ratio(),
    ) * 100.0


def similarity(a, b):
    """0-100: max of character-level ratio and token-set ratio."""
    a_norm = (a or "").strip().lower()
    b_norm = (b or "").strip().lower()
    if not a_norm or not b_norm:
        return 0.0
    char_ratio = SequenceMatcher(None, a_norm, b_norm).ratio() * 100.0
    return max(char_ratio, token_set_ratio(a_norm, b_norm))


def propose_aliases(unmatched_names, profile_keys, threshold=DEFAULT_THRESHOLD):
    """Best profile per unmatched name, above ``threshold``.

    ``profile_keys`` is ``[(profile_index, iterable_of_name_keys)]`` —
    each profile's full answer set (family pattern, profile name,
    existing aliases). Returns ``[(name, profile_index, best_key,
    score)]`` in the input order of ``unmatched_names``, at most one
    entry per name; names scoring below threshold are omitted.

    Ties are broken deterministically: higher score wins, then lower
    profile index (stable across runs so the prompt doesn't flicker
    between equally-scored profiles).
    """
    proposals = []
    for name in unmatched_names:
        best = None  # (score, profile_index, key)
        for profile_index, keys in profile_keys:
            for key in keys or ():
                score = similarity(name, key)
                if score < threshold:
                    continue
                candidate = (score, -int(profile_index), key)
                if best is None or candidate[:2] > best[:2]:
                    best = candidate
        if best is not None:
            score, neg_index, key = best
            proposals.append((name, -neg_index, key, score))
    return proposals
