# -*- coding: utf-8 -*-
"""
Modeless ExternalEvent gateway for ``Place Space Elements``.

Why this module exists
----------------------
For ``OneLevelBased`` / workplane-based families (most of the CED MEP
catalog), ``Document.Create.NewFamilyInstance(point, symbol, level,
NonStructural)`` only binds the new instance to ``level`` when it runs in
a valid Revit API context with the active view already a plan of that
level. Two contexts are NOT valid for this:

  1. A **modal** dialog's event handler — ``ShowDialog`` blocks Revit's
     API context, so ``NewFamilyInstance`` ignores the ``level`` arg and
     the instance lands at ``LevelId = -1`` (no adjustable Level param).
  2. The pushbutton ``main()`` body — also insufficient in practice
     (observed: the level still fails to bind there).

The only reliable context is Revit's main thread serviced via
``ExternalEvent``. This gateway runs ``space_apply.apply_plans`` from
there. ``apply_plans`` switches the active view to each level's plan
(legal here) before placing that level's instances.

This mirrors ``placement_apply`` (the equipment side). The handler class
registers a CLR type, so ``_dev_reload`` intentionally omits this module
from its purge list (re-importing would raise "Duplicate type name within
an assembly"); the gateway singleton survives across runs by design.
"""

import clr  # noqa: F401

from Autodesk.Revit.UI import (  # noqa: E402
    ExternalEvent,
    IExternalEventHandler,
)

import space_apply


class _SpaceApplyExternalEventHandler(IExternalEventHandler):
    """Internal handler. Real work lives on ``SpaceApplyGateway``.

    ``__namespace__`` is required so pythonnet 3 registers this Python
    class with the CLR type system. The fully-qualified name must be
    unique across the session, hence a Space-specific namespace distinct
    from ``placement_apply``'s.
    """

    __namespace__ = "MEPRFP.Automation.SpacePlacement"

    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, uiapp):
        try:
            self._gateway._execute_pending(uiapp)
        except Exception:
            # Never raise into Revit's external-event loop.
            pass

    def GetName(self):
        return "MEPRFP Space Apply"


class SpaceApplyGateway(object):
    """Modeless-safe wrapper around ``space_apply.apply_plans``.

    Usage::

        gateway = get_or_create_gateway()
        gateway.request_apply(doc, uidoc, plans, on_complete=callback)

    The window stays open while the gateway hands the work to Revit's
    main thread via ``ExternalEvent``. ``on_complete`` is invoked with the
    ``_ApplyResult`` after the placement transaction(s) commit.

    Construct via ``get_or_create_gateway()`` — the gateway is a
    per-Revit-session singleton so re-running the pushbutton doesn't try
    to register a second ``IExternalEventHandler`` of the same
    fully-qualified name (which pythonnet 3 + the CLR refuse).
    """

    def __init__(self):
        self._handler = _SpaceApplyExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)
        self._pending = None

    def request_apply(self, doc, uidoc, plans,
                      action=None, on_complete=None):
        """Queue a single placement run. A second call before the previous
        Execute fires REPLACES the queued payload — single-slot by design
        so a double-click can't stack two identical runs."""
        self._pending = {
            "doc": doc,
            "uidoc": uidoc,
            "plans": list(plans or []),
            "action": action or "Place Space Elements (MEPRFP 2.0)",
            "on_complete": on_complete,
        }
        self._event.Raise()

    # ----- internal -------------------------------------------------

    def _execute_pending(self, uiapp):
        payload = self._pending
        if not payload:
            return
        self._pending = None
        callback = payload["on_complete"]
        try:
            result = space_apply.apply_plans(
                payload["doc"],
                payload["plans"],
                action=payload["action"],
                uidoc=payload["uidoc"],
            )
        except Exception as exc:
            result = space_apply._ApplyResult()
            result.warnings.append("Space placement run failed: {}".format(exc))
        if callback is not None:
            try:
                callback(result)
            except Exception:
                # Callback failures shouldn't crash the external event.
                pass


# Module-level singleton — survives between pushbutton runs so we don't
# re-call ``ExternalEvent.Create`` (valid only in a Revit API context)
# and don't trigger the pythonnet "Duplicate type name" error.
_GATEWAY_SINGLETON = None


def get_or_create_gateway():
    """Return the per-Revit-session ``SpaceApplyGateway``.

    The FIRST call must happen inside Revit's API context (i.e. during a
    pushbutton ``main()``); later calls — including from WPF event
    handlers — return the same instance.
    """
    global _GATEWAY_SINGLETON
    if _GATEWAY_SINGLETON is None:
        _GATEWAY_SINGLETON = SpaceApplyGateway()
    return _GATEWAY_SINGLETON
