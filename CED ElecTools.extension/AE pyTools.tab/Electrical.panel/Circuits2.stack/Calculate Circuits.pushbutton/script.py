# -*- coding: utf-8 -*-

import os

from System.Collections.ObjectModel import ObservableCollection
from pyrevit import DB, forms, revit, script

from UIClasses import pathing as ui_pathing

_THIS_DIR = os.path.abspath(os.path.dirname(__file__))
_LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(_THIS_DIR)
if not _LIB_ROOT:
    forms.alert("Could not locate CEDLib.lib.", title="Calculate Circuits", exitscript=True)

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner
from Snippets import _elecutils as eu
from Snippets import revit_helpers
from UIClasses import Resources as UIResources
from UIClasses import load_theme_state_from_config
from UIClasses import resource_loader


doc = revit.doc
logger = script.get_logger()
TITLE = "Calculate Circuits"
ALERT_DATA_PARAM = "Circuit Data_CED"
CIRCUIT_MANAGER_CONFIG_SECTION = "AE-pyTools-CircuitManager"
CALC_PREVIEW_CONFIG_KEY = "calc_preview_enabled"


def _idval(item):
    return revit_helpers.get_elementid_value(item)


def _bool_from_config(value, fallback=False):
    if isinstance(value, bool):
        return bool(value)
    text = str(value or "").strip().lower()
    if text in ("1", "true", "yes", "on"):
        return True
    if text in ("0", "false", "no", "off"):
        return False
    return bool(fallback)


def _load_calc_preview_enabled(default_value=False):
    try:
        cfg = script.get_config(CIRCUIT_MANAGER_CONFIG_SECTION)
        if cfg is None:
            return bool(default_value)
        return _bool_from_config(cfg.get_option(CALC_PREVIEW_CONFIG_KEY, default_value), default_value)
    except Exception:
        return bool(default_value)


def _electrical_panel_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        if os.path.basename(current) == "Electrical.panel":
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def _preview_xaml_path():
    panel_root = _electrical_panel_root(_THIS_DIR)
    if not panel_root:
        return None
    return os.path.abspath(
        os.path.join(
            panel_root,
            "Circuit Manager.pushbutton",
            "CircuitCalculationPreviewWindow.xaml",
        )
    )


def _apply_theme(window):
    resources_root = (
        UIResources.get_resources_root()
        or ui_pathing.resolve_ui_resources_root(_LIB_ROOT)
        or os.path.abspath(os.path.join(_LIB_ROOT, "UIClasses", "Resources"))
    )
    theme_mode, accent_mode = load_theme_state_from_config(
        default_theme="light",
        default_accent="blue",
    )
    resource_loader.apply_theme(
        window,
        resources_root=resources_root,
        theme_mode=theme_mode,
        accent_mode=accent_mode,
    )


class CalculationPreviewRow(object):
    def __init__(self, row):
        data = dict(row or {})
        self.circuit_id = data.get("circuit_id")
        self.circuit = str(data.get("circuit") or "-")
        self.load_name = str(data.get("load_name") or "-")
        self.previous_size = str(data.get("previous_size") or "-")
        self.new_size = str(data.get("new_size") or "-")


class CalculationPreviewWindow(forms.WPFWindow):
    def __init__(self, preview_rows):
        xaml = _preview_xaml_path()
        if not xaml or not os.path.exists(xaml):
            forms.alert("Preview XAML not found.\n\n{}".format(xaml or "<missing>"), title=TITLE, exitscript=True)
        self.decision = None
        forms.WPFWindow.__init__(self, xaml)
        _apply_theme(self)
        rows = [CalculationPreviewRow(x) for x in list(preview_rows or [])]
        preview_list = self.FindName("PreviewList")
        if preview_list is not None:
            preview_list.ItemsSource = ObservableCollection[CalculationPreviewRow](rows)

    def keep_new_clicked(self, sender, args):
        self.decision = "keep_new"
        self.Close()

    def keep_existing_clicked(self, sender, args):
        self.decision = "keep_existing"
        self.Close()

    def skip_clicked(self, sender, args):
        choice = forms.alert(
            "Are you sure you want to skip calculations?\n\nResulting data will be inaccurate.",
            title="Skip Calculations",
            options=["No", "Yes, Skip"],
        )
        if choice != "Yes, Skip":
            return
        self.decision = "skip"
        self.Close()


def _collect_target_circuit_ids(doc):
    selection = list(revit.get_selection() or [])
    if selection:
        selected = []
        for el in selection:
            if isinstance(el, DB.Electrical.ElectricalSystem):
                selected.append(el)
        if not selected:
            selected = eu.get_circuits_from_selection(selection)
    else:
        selected = eu.pick_circuits_from_list(doc, select_multiple=True)

    return [_idval(c.Id) for c in selected if isinstance(c, DB.Electrical.ElectricalSystem)]


def main():
    circuit_ids = _collect_target_circuit_ids(doc)
    if not circuit_ids:
        forms.alert("No circuits selected.", exitscript=True)

    options = {
        "show_output": True,
        "calc_preview_enabled": _load_calc_preview_enabled(False),
    }
    runner = build_default_runner(alert_parameter_name=ALERT_DATA_PARAM)

    while True:
        request = OperationRequest(
            operation_key="calculate_circuits",
            circuit_ids=circuit_ids,
            source="ribbon",
            options=dict(options),
        )
        result = runner.run(request, doc)
        if not result:
            return
        if result.get("status") == "preview_required":
            window = CalculationPreviewWindow(result.get("preview_rows") or [])
            window.ShowDialog()
            decision = getattr(window, "decision", None)
            if decision in ("keep_new", "keep_existing"):
                options["calc_preview_decision"] = decision
                continue
            if decision == "skip":
                return
            logger.info("Calculate circuits request ended: %s", result)
            return
        if result.get("status") != "ok":
            logger.info("Calculate circuits request ended: %s", result)
        return


main()
