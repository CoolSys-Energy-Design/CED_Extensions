# -*- coding: utf-8 -*-
"""Unicode-preserving text boundary for Revit API values.

Revit exposes names and parameter text as CLR ``System.String`` values.  In
IronPython, calling ``str()`` on those values can ask the runtime to encode
them through the active Windows code page.  Parameter Monitor must never do
that: it converts CLR strings directly to Python Unicode and rejects raw byte
strings rather than guessing an encoding for them.
"""

from __future__ import print_function


class TextConversionError(ValueError):
    """Raised when an unexpected byte string reaches monitor data."""


try:
    _UNICODE_TYPE = unicode
    _BYTE_STRING_TYPE = str
    _CHARACTER_FROM_ORDINAL = unichr
    _IS_PYTHON_2 = True
except NameError:
    _UNICODE_TYPE = str
    _BYTE_STRING_TYPE = bytes
    _CHARACTER_FROM_ORDINAL = chr
    _IS_PYTHON_2 = False


def _is_system_string(value):
    """Identify a CLR string without invoking Python ``str(value)``."""
    try:
        return value.GetType().FullName == "System.String"
    except Exception:
        return False


def is_text_value(value):
    """Return whether ``value`` must be normalized before JSON storage."""
    if value is None:
        return False
    if _is_system_string(value) or isinstance(value, _UNICODE_TYPE):
        return True
    if isinstance(value, (bytes, bytearray)):
        return True
    return bool(_IS_PYTHON_2 and isinstance(value, _BYTE_STRING_TYPE))


def _failure(context, detail):
    location = " at {}".format(context) if context else ""
    return TextConversionError(
        "Unexpected byte text{} ({}). Revit text must remain a CLR/System.String "
        "or Python Unicode value; no encoding guess was attempted.".format(
            location, detail
        )
    )


def to_text(value, fallback=u"", context=None):
    """Return a Unicode value without ever decoding arbitrary bytes.

    ``unicode(System.String)`` is the IronPython-safe conversion prescribed by
    the CLR type boundary.  The character loop is only a Unicode fallback for
    an unusual CLR proxy; it never interprets bytes under a code page.
    """
    if value is None:
        return fallback
    if _is_system_string(value):
        try:
            return _UNICODE_TYPE(value)
        except Exception:
            try:
                return u"".join([
                    _CHARACTER_FROM_ORDINAL(ord(character))
                    for character in value
                ])
            except Exception as ex:
                raise TextConversionError(
                    "Could not convert Revit System.String{}: {}".format(
                        " at {}".format(context) if context else "", ex
                    )
                )
    if isinstance(value, _UNICODE_TYPE):
        return value
    if isinstance(value, (bytes, bytearray)):
        raise _failure(context, "bytes")
    if _IS_PYTHON_2 and isinstance(value, _BYTE_STRING_TYPE):
        raise _failure(context, "Python str")
    try:
        return _UNICODE_TYPE(value)
    except Exception as ex:
        raise TextConversionError(
            "Could not convert text{}: {}".format(
                " at {}".format(context) if context else "", ex
            )
        )


def diagnostic_text(value, fallback=u"<unavailable>"):
    """Best-effort Unicode text for a diagnostic without hiding the failure."""
    try:
        return to_text(value, fallback=fallback)
    except Exception:
        return fallback
