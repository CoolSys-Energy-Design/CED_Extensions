# -*- coding: utf-8 -*-
"""Small, runtime-independent helpers for documentation search highlighting."""

from __future__ import print_function

import re

try:
    _text_type = unicode
except NameError:
    _text_type = str


def _text(value):
    if value is None:
        return _text_type("")
    if isinstance(value, _text_type):
        return value
    return _text_type(value)


def query_terms(query):
    """Return unique, non-empty query terms in display order."""
    terms = []
    seen = set()
    for value in re.split(r"\s+", _text(query).strip(), flags=re.UNICODE):
        if not value:
            continue
        key = value.lower()
        if key in seen:
            continue
        seen.add(key)
        terms.append(value)
    return terms


def highlight_segments(value, query):
    """Split text into ``(text, is_match)`` tuples, case-insensitively."""
    text = _text(value)
    terms = query_terms(query) if isinstance(query, (_text_type, bytes)) else list(query or [])
    terms = sorted([_text(term) for term in terms if _text(term)], key=len, reverse=True)
    if not text or not terms:
        return [(text, False)] if text else []
    pattern = re.compile("|".join(re.escape(term) for term in terms), re.IGNORECASE | re.UNICODE)
    segments = []
    position = 0
    for match in pattern.finditer(text):
        if match.start() > position:
            segments.append((text[position:match.start()], False))
        segments.append((match.group(0), True))
        position = match.end()
    if position < len(text):
        segments.append((text[position:], False))
    return segments or [(text, False)]
