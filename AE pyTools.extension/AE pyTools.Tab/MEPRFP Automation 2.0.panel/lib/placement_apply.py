# -*- coding: utf-8 -*-
"""
Modeless ExternalEvent gateway for ``Place from CAD or Linked Model``.

Why this module exists
----------------------
For workplane-based families (most of the CED MEP catalog),
``Document.Create.NewFamilyInstance(point, symbol, level, NonStructural)``
silently ignores the ``level`` argument and binds the new instance to the
**active view's** ``GenLevel``. If the active view isn't a plan whose
level matches, the instance lands at ``LevelId = -1`` even though a valid
Level was passed. So the load-bearing step is switching the active view
to the target level's plan BEFORE ``NewFamilyInstance`` runs — exactly
what the working MCP probe does.

Two Revit constraints make that switch impossible from the old modal
dialog:

  1. ``UIDocument.ActiveView = view`` throws ``InvalidOperationException``
     while a Transaction is open.
  2. It also throws from inside a **modal** dialog's event handler
     (``ShowDialog`` blocks Revit's API context).

The old dialog tried to pre-switch the view in its Place handler and
swallowed the resulting exception, so the view never changed and every
placement bound to ``LevelId = -1``.

This gateway fixes it by running the whole placement on Revit's main
thread via ``ExternalEvent``. From there the API context is valid and no
modal dialog blocks it, so the view switch succeeds. We switch once per
target level (outside any transaction), then open a transaction for that
level and place only that level's fixtures — reproducing the MCP
sequence for every level in the run.

The implementation mirrors ``circuit_apply``'s gateway. The handler class
must NOT be re-imported between script runs (it registers a .NET type);
``_dev_reload`` intentionally omits ``placement_apply`` from its purge
list for the same reason it omits ``circuit_apply``.
"""

import clr  # noqa: F401

from Autodesk.Revit.UI import (  # noqa: E402
    ExternalEvent,
    IExternalEventHandler,
)

import placement


# ---------------------------------------------------------------------
# Result aggregation
# ---------------------------------------------------------------------

def _merge_result(agg, res):
    """Fold one pass's ``PlacementResult`` into the aggregate."""
    if res is None:
        return
    agg.placed_fixture_count += res.placed_fixture_count
    agg.placed_annotation_count += res.placed_annotation_count
    agg.element_linker_writes += res.element_linker_writes
    agg.static_param_writes += res.static_param_writes
    agg.parent_directive_writes += res.parent_directive_writes
    agg.byparent_group_count += getattr(res, "byparent_group_count", 0)
    agg.byparent_suffix_writes += getattr(res, "byparent_suffix_writes", 0)
    agg.skipped_already_placed += res.skipped_already_placed
    agg.normalized_match_count += res.normalized_match_count
    agg.substituted_type_count += res.substituted_type_count
    agg.warnings.extend(res.warnings)
    agg.errors.extend(res.errors)


def _pass_options(base, only_level_id):
    """Clone ``base`` PlacementOptions, overriding ``only_level_id`` so a
    single pass places just one level's LEDs."""
    return placement.PlacementOptions(
        skip_already_placed=base.skip_already_placed,
        default_level_id=base.default_level_id,
        transaction_action=base.transaction_action,
        allow_type_substitution=base.allow_type_substitution,
        category_filter=base.category_filter,
        uidoc=base.uidoc,
        only_level_id=only_level_id,
        host_relative_height_inches=base.host_relative_height_inches,
    )


def _run_placement(doc, uidoc, matches, options, transaction_name):
    """Place every match, switching the active view to each target
    level's plan (outside any transaction) before placing that level's
    fixtures. Returns an aggregated ``PlacementResult``.

    MUST run in a valid Revit API context (i.e. from inside an
    ``ExternalEvent`` Execute), never from a modal dialog handler — that
    is the whole reason the active-view switch works here and didn't in
    the old modal flow.
    """
    from pyrevit import revit

    aggregate = placement.PlacementResult()

    # Read-only pre-pass (no transaction open) — which levels will we bind?
    levels = placement.collect_target_levels(doc, matches, options)
    if not levels:
        # No in-scope LEDs resolved a level (e.g. picked-point with no
        # level). Place once under the current view.
        levels = [placement._ALL_LEVELS]

    original_active_view = None
    if uidoc is not None:
        try:
            original_active_view = uidoc.ActiveView
        except Exception:
            original_active_view = None

    try:
        for level_id in levels:
            # Switch the active view to this level's plan BEFORE opening
            # the transaction. This is legal here (valid API context, no
            # txn open, no modal dialog) and is what binds the Level on
            # workplane-based families.
            #
            # Only switch when the user's CURRENT view can't already bind
            # this level — i.e. it isn't a plan of this level. This keeps
            # the tool from jumping to a view the user didn't have open
            # whenever they're already working in the right level's plan.
            if (level_id is not None
                    and level_id is not placement._ALL_LEVELS
                    and uidoc is not None
                    and not placement.active_view_is_level_plan(doc, level_id)):
                plan_view = placement._find_plan_view_for_level(doc, level_id)
                if plan_view is not None:
                    try:
                        uidoc.ActiveView = plan_view
                    except Exception as exc:
                        aggregate.warnings.append(
                            "Could not activate plan view for level {} "
                            "({}); fixtures on that level may not bind."
                            .format(level_id, type(exc).__name__)
                        )
            pass_options = _pass_options(options, level_id)
            try:
                with revit.Transaction(transaction_name, doc=doc):
                    res = placement.execute_placement(doc, matches, pass_options)
                _merge_result(aggregate, res)
            except Exception as exc:
                aggregate.errors.append(
                    "Placement transaction failed for level {}: {}".format(
                        level_id, exc
                    )
                )
    finally:
        if uidoc is not None and original_active_view is not None:
            try:
                uidoc.ActiveView = original_active_view
            except Exception:
                pass

    return aggregate


# ---------------------------------------------------------------------
# Modeless ExternalEvent gateway
# ---------------------------------------------------------------------

class _PlacementExternalEventHandler(IExternalEventHandler):
    """Internal handler. Real work lives on ``PlacementGateway``.

    ``__namespace__`` is required so pythonnet 3 registers this Python
    class with the CLR type system (see ``circuit_apply`` for the full
    explanation). The fully-qualified name must be unique across the
    session, hence a Placement-specific namespace.
    """

    __namespace__ = "MEPRFP.Automation.Placement"

    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, uiapp):
        try:
            self._gateway._execute_pending(uiapp)
        except Exception:
            # Never raise into Revit's external-event loop.
            pass

    def GetName(self):
        return "MEPRFP Placement Apply"


class PlacementGateway(object):
    """Modeless-safe wrapper around the per-level placement run.

    Usage::

        gateway = get_or_create_gateway()
        gateway.request_placement(doc, uidoc, matches, options,
                                  on_complete=callback)

    The window stays open while the gateway hands the work to Revit's
    main thread via ``ExternalEvent``. ``on_complete`` is invoked with the
    aggregated ``PlacementResult`` after every level's transaction commits.

    Construct via ``get_or_create_gateway()`` — the gateway is a
    per-Revit-session singleton so re-running the pushbutton doesn't try
    to register a second ``IExternalEventHandler`` of the same
    fully-qualified name (which pythonnet 3 + the CLR refuse).
    """

    def __init__(self):
        self._handler = _PlacementExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)
        self._pending = None

    def request_placement(self, doc, uidoc, matches, options,
                          transaction_name=None, on_complete=None):
        """Queue a single placement run. A second call before the
        previous Execute fires REPLACES the queued payload — single-slot
        by design so a double-click can't stack two identical runs."""
        self._pending = {
            "doc": doc,
            "uidoc": uidoc,
            "matches": list(matches or []),
            "options": options,
            "transaction_name": (
                transaction_name or "Place from CAD or Linked Model (MEPRFP 2.0)"
            ),
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
            result = _run_placement(
                payload["doc"],
                payload["uidoc"],
                payload["matches"],
                payload["options"],
                payload["transaction_name"],
            )
        except Exception as exc:
            result = placement.PlacementResult()
            result.errors.append("Placement run failed: {}".format(exc))
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
    """Return the per-Revit-session ``PlacementGateway``.

    The FIRST call must happen inside Revit's API context (i.e. during a
    pushbutton ``main()``); later calls — including from WPF event
    handlers — return the same instance.
    """
    global _GATEWAY_SINGLETON
    if _GATEWAY_SINGLETON is None:
        _GATEWAY_SINGLETON = PlacementGateway()
    return _GATEWAY_SINGLETON


# ---------------------------------------------------------------------
# Profile-save gateway (fuzzy-alias persistence)
# ---------------------------------------------------------------------

class _ProfileSaveExternalEventHandler(IExternalEventHandler):
    """Internal handler for ``ProfileSaveGateway``. Separate event from
    the placement handler so a queued alias-save can never clobber (or
    be clobbered by) a queued placement in the single-slot gateways.

    Distinct ``__namespace__`` — the CLR requires each registered
    handler type's fully-qualified name to be session-unique.
    """

    __namespace__ = "MEPRFP.Automation.ProfileSave"

    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, uiapp):
        try:
            self._gateway._execute_pending(uiapp)
        except Exception:
            # Never raise into Revit's external-event loop.
            pass

    def GetName(self):
        return "MEPRFP Profile Save"


class ProfileSaveGateway(object):
    """Persist the in-memory profile store from a modeless window.

    The placement window mutates ``profile_data`` in place (e.g. when
    the user accepts a fuzzy-match alias) but, being modeless, has no
    valid API context to open a transaction in. This gateway runs
    ``active_yaml.save_active_data`` on Revit's main thread. Slot
    replacement is harmless here: the payload references the live
    ``profile_data`` dict, so the last Execute serializes the superset
    of all accepted edits.
    """

    def __init__(self):
        self._handler = _ProfileSaveExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)
        self._pending = None

    def request_save(self, doc, profile_data, action="MEPRFP 2.0 edit",
                     on_complete=None):
        self._pending = {
            "doc": doc,
            "profile_data": profile_data,
            "action": action,
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
        error = None
        try:
            from pyrevit import revit
            import active_yaml
            with revit.Transaction(payload["action"], doc=payload["doc"]):
                active_yaml.save_active_data(
                    payload["doc"], payload["profile_data"],
                    action=payload["action"],
                )
        except Exception as exc:
            error = "Profile save failed: {}".format(exc)
        if callback is not None:
            try:
                callback(error)
            except Exception:
                pass


_SAVE_GATEWAY_SINGLETON = None


def get_or_create_save_gateway():
    """Return the per-Revit-session ``ProfileSaveGateway``. Same
    first-call-needs-API-context rule as ``get_or_create_gateway``."""
    global _SAVE_GATEWAY_SINGLETON
    if _SAVE_GATEWAY_SINGLETON is None:
        _SAVE_GATEWAY_SINGLETON = ProfileSaveGateway()
    return _SAVE_GATEWAY_SINGLETON
