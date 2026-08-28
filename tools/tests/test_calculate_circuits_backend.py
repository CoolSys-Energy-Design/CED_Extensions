# -*- coding: utf-8 -*-
"""Regression tests for the Calculate Circuits staged backend contract."""

from pathlib import Path
import importlib.util
import sys
import types


ROOT = Path(__file__).resolve().parents[2]


class _ElementId(object):
    def __init__(self, value):
        self.Value = int(value)
        self.IntegerValue = int(value)

    def __eq__(self, other):
        return isinstance(other, _ElementId) and self.Value == other.Value


_ElementId.InvalidElementId = _ElementId(-1)


class _StorageType(object):
    String = "String"
    Integer = "Integer"
    Double = "Double"
    ElementId = "ElementId"


class _Transaction(object):
    def __init__(self, unused_doc, unused_name):
        self.started = False

    def Start(self):
        self.started = True

    def Commit(self):
        self.started = False

    def Assimilate(self):
        self.started = False

    def RollBack(self):
        self.started = False


class _Logger(object):
    def info(self, unused_message):
        pass

    def warning(self, unused_message):
        pass

    def error(self, unused_message):
        pass


def _module(name):
    value = types.ModuleType(name)
    sys.modules[name] = value
    return value


def _load_backend_modules():
    db = _module("Autodesk.Revit.DB")
    db.ElementId = _ElementId
    db.StorageType = _StorageType
    db.Transaction = _Transaction
    db.TransactionGroup = _Transaction
    db.BuiltInParameter = types.SimpleNamespace(
        RBS_ELEC_VOLTAGE="voltage",
        RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM="distribution_system",
    )
    db.UnitTypeId = types.SimpleNamespace(Volts="volts")
    db.UnitUtils = types.SimpleNamespace(ConvertFromInternalUnits=lambda value, unused: value)
    db.CheckoutStatus = types.SimpleNamespace(OwnedByOtherUser="owned")
    db.WorksharingUtils = types.SimpleNamespace(GetCheckoutStatus=lambda unused_doc, unused_id: "free")

    autodesk = _module("Autodesk")
    revit_package = _module("Autodesk.Revit")
    autodesk.Revit = revit_package
    revit_package.DB = db

    electrical = _module("Autodesk.Revit.DB.Electrical")

    class _Descriptor(object):
        def __init__(self, name):
            self.name = name

        def __get__(self, instance):
            return getattr(instance, self.name, None)

    electrical.ElectricalSystem = types.SimpleNamespace(
        ApparentLoad=_Descriptor("ApparentLoad"),
        ApparentCurrent=_Descriptor("ApparentCurrent"),
        PowerFactor=_Descriptor("PowerFactor"),
        PolesNumber=_Descriptor("PolesNumber"),
    )

    system = _module("System")
    system.Int64 = int

    pyrevit = _module("pyrevit")
    pyrevit.DB = db
    pyrevit.forms = types.SimpleNamespace(alert=lambda *args, **kwargs: None)
    pyrevit.script = types.SimpleNamespace(
        get_logger=lambda: _Logger(),
        get_output=lambda: types.SimpleNamespace(),
    )

    snippets = _module("Snippets")
    eu = _module("Snippets._elecutils")
    eu.is_circuit_eligible = lambda unused: True
    categories = _module("Snippets.categories")
    categories._revit_major_version = lambda doc=None: getattr(doc, "version", 0)
    snippets._elecutils = eu
    snippets.categories = categories

    helper_path = ROOT / "CEDLib.lib" / "Snippets" / "revit_helpers.py"
    helper_spec = importlib.util.spec_from_file_location("Snippets.revit_helpers", str(helper_path))
    helpers = importlib.util.module_from_spec(helper_spec)
    sys.modules[helper_spec.name] = helpers
    helper_spec.loader.exec_module(helpers)
    snippets.revit_helpers = helpers

    _module("CEDElectrical")
    domain = _module("CEDElectrical.Domain")
    settings_manager = _module("CEDElectrical.Domain.settings_manager")
    settings_manager.load_circuit_settings = lambda unused_doc: types.SimpleNamespace(
        to_json=lambda: "{}",
    )
    settings_manager.RESULT_PARAM_NAMES = []
    domain.settings_manager = settings_manager
    _module("CEDElectrical.Model")
    branch_module = _module("CEDElectrical.Model.CircuitBranch")
    branch_module.CircuitBranch = object
    branch_module.get_native_circuit_type_label = lambda circuit: getattr(circuit, "native_type", "CIRCUIT")
    settings_module = _module("CEDElectrical.Model.circuit_settings")
    settings_module.CircuitSettings = type("CircuitSettings", (), {})

    operation_path = (
        ROOT
        / "CEDLib.lib"
        / "CEDElectrical"
        / "Application"
        / "operations"
        / "calculate_circuits_operation.py"
    )
    operation_spec = importlib.util.spec_from_file_location(
        "calculate_circuits_operation_under_test",
        str(operation_path),
    )
    operation = importlib.util.module_from_spec(operation_spec)
    operation_spec.loader.exec_module(operation)
    return helpers, operation


HELPERS, CALCULATE = _load_backend_modules()


class _ElementIdParameter(object):
    StorageType = _StorageType.ElementId
    HasValue = True

    def __init__(self, value):
        self.value = value
        self.writes = []

    def AsElementId(self):
        return self.value

    def Set(self, value):
        self.value = value
        self.writes.append(value)
        return True


class _IntegerParameter(object):
    StorageType = _StorageType.Integer

    def __init__(self, value=0, has_value=True):
        self.value = int(value)
        self.HasValue = bool(has_value)
        self.writes = []

    def AsInteger(self):
        return self.value

    def Set(self, value):
        self.value = int(value)
        self.HasValue = True
        self.writes.append(int(value))
        return True


def test_cleared_elementid_is_a_noop_and_native_elementids_are_supported():
    cleared = _ElementIdParameter(_ElementId.InvalidElementId)
    assert HELPERS.parameter_matches_value(cleared, None)
    assert not HELPERS.set_parameter_if_changed(cleared, None)
    assert cleared.writes == []

    changed = _ElementIdParameter(_ElementId(7))
    desired = _ElementId(9)
    assert HELPERS.set_parameter_if_changed(changed, desired)
    assert changed.writes == [desired]
    assert HELPERS.parameter_matches_value(changed, 9)


def test_unset_yesno_is_written_even_when_asinteger_returns_zero():
    parameter = _IntegerParameter(value=0, has_value=False)
    assert not HELPERS.parameter_matches_value(parameter, 0)
    assert HELPERS.set_parameter_if_changed(parameter, 0)
    assert parameter.writes == [0]
    assert parameter.HasValue


def test_first_calculation_explicitly_disables_user_override():
    operation = CALCULATE.CalculateCircuitsOperation.__new__(
        CALCULATE.CalculateCircuitsOperation
    )

    class _FirstCalculationBranch(object):
        is_special = False
        _calc_first_calculation = True
        _calc_preview_keep_existing = False
        _calc_preview_force_auto = False
        neutral_wire_quantity = 1
        isolated_ground_wire_quantity = 0
        branch_type = "BRANCH"
        panel = "P1"
        circuit_number = "1"
        load_name = "LOAD"
        rating = 20
        frame = 20
        length = 10
        circuit_notes = ""
        voltage_drop_percentage = 1.0
        hot_wire_size = "12"
        number_of_wires = 3
        number_of_sets = 1
        hot_wire_quantity = 1
        ground_wire_size = "12"
        ground_wire_quantity = 1
        neutral_wire_size = "12"
        isolated_ground_wire_size = ""
        wire_material = "CU"
        wire_temp_rating = "75"
        wire_insulation = "THHN"
        conduit_size = '3/4"'
        conduit_type = "EMT"
        conduit_fill_percentage = 0.2
        circuit_load_current = 10
        circuit_base_ampacity = 20
        wire_length_makeup = 0

        def get_wire_size_callout(self):
            return "3#12"

        def get_conduit_and_wire_size(self):
            return '3#12, 1#12G; 3/4" EMT'

    values = operation._collect_shared_param_values(_FirstCalculationBranch())
    assert values["CKT_User Override_CED"] == 0
    assert values["CKT_Include Neutral_CED"] == 1
    assert values["CKT_Include Isolated Ground_CED"] == 0


def test_compound_preview_requires_original_operation_rerun():
    operation = CALCULATE.CalculateCircuitsOperation.__new__(
        CALCULATE.CalculateCircuitsOperation
    )
    standalone = types.SimpleNamespace(options={})
    compound = types.SimpleNamespace(options={"use_existing_transaction_group": True})
    assert operation._preview_can_use_staged_apply(standalone)
    assert not operation._preview_can_use_staged_apply(compound)


def test_dependency_version_change_invalidates_stage():
    operation = CALCULATE.CalculateCircuitsOperation.__new__(
        CALCULATE.CalculateCircuitsOperation
    )
    element = types.SimpleNamespace(Id=_ElementId(22), VersionGuid="new")
    doc = types.SimpleNamespace(GetElement=lambda unused_id: element)
    snapshot = {
        "dependency_versions": [
            {"element_id": 22, "version_guid": "old"},
        ],
    }
    valid, reason = operation._validate_dependency_versions(doc, snapshot)
    assert not valid
    assert reason == "calculation_dependency_changed"


def test_keep_existing_applies_fully_recalculated_variant():
    writes = []
    alerts = []

    class _Writer(object):
        def write_circuit_parameters(self, unused_circuit, values):
            writes.append(dict(values))

        def write_connected_elements(self, *args, **kwargs):
            return 0, 0

    class _AlertStore(object):
        def write_alert_payload(self, unused_circuit, payload):
            alerts.append(payload)

        def clear_alert_payload(self, unused_circuit):
            alerts.append(None)

    operation = CALCULATE.CalculateCircuitsOperation(None, _Writer(), _AlertStore())
    operation._load_effective_settings = lambda unused_doc, unused_request: types.SimpleNamespace()
    circuit = types.SimpleNamespace(Id=_ElementId(1))
    operation._rehydrate_staged_item = lambda unused_doc, unused_item: (
        True,
        "",
        circuit,
        [],
    )
    request = types.SimpleNamespace(
        options={"calc_preview_decision": "keep_existing", "show_output": False}
    )
    stage = {
        "version": 2,
        "circuits": [
            {
                "circuit_id": 1,
                "is_special": False,
                "preview_changed": True,
                "parameter_values": {"hot": "new", "dependent": "new-derived"},
                "keep_existing_parameter_values": {
                    "hot": "existing",
                    "dependent": "existing-derived",
                },
                "builtin_values": {},
                "connected_element_ids": [],
                "alert_payload": {"variant": "new"},
                "keep_existing_alert_payload": {"variant": "existing"},
                "runtime_alert_rows": [],
                "keep_existing_runtime_alert_rows": [],
            }
        ],
        "preview_rows": [{"circuit_id": 1}],
        "locked_rows": [],
    }
    result = operation.apply_staged_result(
        request,
        types.SimpleNamespace(),
        stage,
        validate=False,
    )
    assert result["status"] == "ok"
    assert writes == [{"hot": "existing", "dependent": "existing-derived"}]
    assert alerts == [{"variant": "existing"}]


def test_keep_existing_variant_runs_the_complete_engineering_sequence():
    calls = []

    class _KeepBranch(object):
        def __init__(self, circuit, settings=None, preview_values=None, connected_elements=None):
            self.circuit = circuit
            self.preview_values = dict(preview_values or {})
            self.connected_elements = list(connected_elements or [])

        def calculate_hot_wire_size(self):
            calls.append("hot")

        def calculate_neutral_wire_size(self):
            calls.append("neutral")

        def calculate_ground_wire_size(self):
            calls.append("ground")

        def calculate_isolated_ground_wire_size(self):
            calls.append("ig")

        def calculate_conduit_size(self):
            calls.append("conduit")

    source = types.SimpleNamespace(
        circuit=types.SimpleNamespace(Id=_ElementId(3)),
        connected_elements=[types.SimpleNamespace()],
    )
    operation = CALCULATE.CalculateCircuitsOperation.__new__(
        CALCULATE.CalculateCircuitsOperation
    )
    original_branch_type = CALCULATE.CircuitBranch
    CALCULATE.CircuitBranch = _KeepBranch
    try:
        rebuilt = operation._rebuild_branches_with_existing_sizes(
            [source],
            {
                3: {
                    "CKT_Number of Sets_CED": 2,
                    "CKT_Wire Hot Size_CEDT": "1/0",
                    "CKT_Wire Ground Size_CEDT": "6",
                    "Conduit Size_CEDT": '2"',
                    "Conduit Type_CEDT": "EMT",
                }
            },
            {},
            types.SimpleNamespace(),
            changed_ids={3},
        )
    finally:
        CALCULATE.CircuitBranch = original_branch_type

    assert calls == ["hot", "neutral", "ground", "ig", "conduit"]
    assert rebuilt[0].preview_values["CKT_User Override_CED"] == 1
    assert rebuilt[0].preview_values["CKT_Wire Hot Size_CEDT"] == "1/0"
    assert rebuilt[0]._calc_preview_keep_existing is True


def test_special_processing_is_version_gated():
    circuit = types.SimpleNamespace(native_type="SPARE")
    assert CALCULATE.get_special_processing_mode(
        types.SimpleNamespace(version=2025), circuit
    ) == CALCULATE.SPECIAL_MODE_LEGACY
    assert CALCULATE.get_special_processing_mode(
        types.SimpleNamespace(version=2026), circuit
    ) == CALCULATE.SPECIAL_MODE_REGULAR_COMPATIBLE


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
