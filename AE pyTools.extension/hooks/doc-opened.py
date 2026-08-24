# -*- coding: utf-8 -*-
"""doc-opened hook :: CPython stdlib path guard (Revit 2027+).

Companion to the CPython engine primer in ``startup.py``. The primer
initializes pythonnet's CPython engine at session load (pre-document)
and seeds the stdlib paths — but something in the Revit 2027 process
(leading suspect: another add-in's embedded Python, e.g. Dynamo,
initializing at document open) can wipe the interpreter's ``sys.path``
afterward. pyRevit then snapshots that empty list per command
(``CPythonEngine.StoreSearchPaths``), leaving every ``#! python3``
button without a standard library ("No module named 'configparser'").

This hook runs after every document open and repairs both places the
paths live:

1. the live interpreter ``sys.path`` (re-seed zip + engine dir), and
2. the ``_sysPaths`` snapshot of any already-cached CPython engine in
   the ``PYREVITCachedEngines`` AppDomain slot.

Verified interactively on Revit 2027.0.20 + pyRevit 6.5.4 (2026-08-24):
patching exactly these two spots took the buttons from dead to working.
Every step is fail-safe and breadcrumbed to
``%APPDATA%\\pyRevit\\CED_cpython_primer.log``.
"""

import os
import re
import time


_PRIMER_LOG = os.path.join(
    os.environ.get("APPDATA", ""), "pyRevit", "CED_cpython_primer.log")


def _note(msg):
    try:
        with open(_PRIMER_LOG, "a") as f:
            f.write("%s [doc-opened] %s\n" % (
                time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def _stdlib_paths():
    """Return (zip, engine_dir) for the newest bundled CPython engine."""
    import pyrevit as _pyrevit
    cengines = os.path.join(_pyrevit.HOME_DIR, "bin", "cengines")
    if not os.path.isdir(cengines):
        return None, None
    engine_dirs = sorted(
        d for d in os.listdir(cengines)
        if re.match(r"^CPY\d+$", d)
        and os.path.isdir(os.path.join(cengines, d))
    )
    if not engine_dirs:
        return None, None
    engine_dir = os.path.join(cengines, engine_dirs[-1])
    dlls = sorted(
        f for f in os.listdir(engine_dir)
        if re.match(r"^python3\d+\.dll$", f.lower())
    )
    if not dlls:
        return None, None
    stdlib_zip = os.path.splitext(os.path.join(engine_dir, dlls[-1]))[0] + ".zip"
    return stdlib_zip, engine_dir


def _guard_cpython_paths():
    from pyrevit import HOST_APP
    from System import AppDomain, Array, Object
    from System.Reflection import BindingFlags

    try:
        if int(str(HOST_APP.version)) < 2027:
            return
    except Exception:
        return

    pn = None
    for asm in AppDomain.CurrentDomain.GetAssemblies():
        try:
            if asm.GetName().Name == "pyRevitLabs.PythonNet":
                pn = asm
                break
        except Exception:
            continue
    if pn is None:
        return  # engine never touched this session; primer handles init
    engine_t = pn.GetType("Python.Runtime.PythonEngine")
    py_t = pn.GetType("Python.Runtime.Py")
    if engine_t is None or py_t is None:
        return
    if not engine_t.GetProperty("IsInitialized").GetValue(None, None):
        return  # nothing to repair yet

    stdlib_zip, engine_dir = _stdlib_paths()
    if not stdlib_zip:
        _note("ABORT: could not locate bundled CPython engine")
        return

    # 1. live interpreter sys.path
    snippet = (
        "import sys\n"
        "for _p in (r'{dir}', r'{zip}'):\n"
        "    if _p not in sys.path:\n"
        "        sys.path.insert(0, _p)\n"
    ).format(zip=stdlib_zip, dir=engine_dir)
    gil = py_t.GetMethod("GIL").Invoke(None, None)
    try:
        rc = engine_t.GetMethod("RunSimpleString").Invoke(
            None, Array[Object]([snippet]))
    finally:
        gil.Dispose()
    if rc != 0:
        _note("sys.path re-seed FAILED (rc=%s)" % rc)
        return

    # 2. _sysPaths snapshots of cached CPython engines
    patched = 0
    try:
        engines = AppDomain.CurrentDomain.GetData("PYREVITCachedEngines")
        if engines is not None:
            for kv in engines:
                eng = kv.Value
                et = eng.GetType()
                if "CPythonEngine" not in et.FullName:
                    continue
                fld = et.GetField(
                    "_sysPaths",
                    BindingFlags.NonPublic | BindingFlags.Instance)
                if fld is None:
                    continue
                paths = fld.GetValue(eng)
                existing = [p for p in paths]
                for want in (stdlib_zip, engine_dir):
                    if want not in existing:
                        paths.Insert(0, want)
                        patched += 1
    except Exception as exc:
        _note("cached-engine patch failed: %s" % exc)
    _note("re-seeded sys.path; patched %d cached engine path entries" % patched)


try:
    _guard_cpython_paths()
except Exception as _exc:
    _note("EXCEPTION: %s" % _exc)
