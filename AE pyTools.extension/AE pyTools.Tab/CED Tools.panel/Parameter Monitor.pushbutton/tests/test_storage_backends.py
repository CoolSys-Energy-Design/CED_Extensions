# -*- coding: utf-8 -*-
from __future__ import print_function

import unittest

import models
import storage_service


class _ValidEntity(object):
    def IsValid(self):
        return True


class _Storage(object):
    def __init__(self, entity=None):
        self.entity = entity or _ValidEntity()

    def GetEntity(self, schema):
        return self.entity


class StorageBackendTests(unittest.TestCase):
    def setUp(self):
        self._saved = {}

    def _replace(self, name, value):
        if name not in self._saved:
            self._saved[name] = getattr(storage_service, name)
        setattr(storage_service, name, value)

    def tearDown(self):
        for name, value in self._saved.items():
            setattr(storage_service, name, value)

    def test_datastorage_contract_uses_extensible_storage_namespace(self):
        self.assertEqual(
            storage_service._DATASTORAGE_TYPE_NAMES[0],
            "Autodesk.Revit.DB.ExtensibleStorage.DataStorage, RevitAPI",
        )

    def test_find_data_storage_uses_resolved_type(self):
        expected_type = object()
        expected_storage = _Storage()
        calls = []

        class FakeCollector(object):
            def __init__(self, document):
                calls.append(("document", document))

            def OfClass(self, element_type):
                calls.append(("type", element_type))
                return [expected_storage]

        self._replace("FilteredElementCollector", FakeCollector)
        found = storage_service._find_data_storage(
            "document",
            "schema",
            data_storage_type=expected_type,
        )
        self.assertIs(found, expected_storage)
        self.assertEqual(calls[1], ("type", expected_type))

    def test_create_data_storage_uses_resolved_class(self):
        created = _Storage()

        class FakeDataStorage(object):
            @staticmethod
            def Create(document):
                self.assertEqual(document, "document")
                return created

        result = storage_service._create_data_storage(
            "document",
            data_storage_type=FakeDataStorage,
        )
        self.assertIs(result, created)

    def test_multiline_text_spec_is_preferred(self):
        multiline = object()

        class FakeStringSpecs(object):
            MultilineText = multiline
            Text = object()

        class FakeSpecTypeId(object):
            String = FakeStringSpecs

        self._replace("SpecTypeId", FakeSpecTypeId)
        self.assertIs(storage_service._multiline_text_spec(), multiline)

    def test_save_uses_global_parameter_when_extensible_storage_fails(self):
        calls = []

        def fail_primary(document, identity, payload_text, transaction_name):
            raise RuntimeError("primary unavailable")

        def save_backup(document, payload_text, transaction_name):
            calls.append((document, payload_text, transaction_name))
            return object()

        self._replace("_write_extensible_storage", fail_primary)
        self._replace("_write_fallback_payload", save_backup)
        document = type(
            "Document",
            (object,),
            {"Title": "Test", "PathName": "", "IsWorkshared": False},
        )()
        store = models.new_project_store()
        result = storage_service.save(document, store)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], document)
        self.assertIn('"tracking_sets":[]', calls[0][1])
        self.assertEqual(result.get("project_identity", {}).get("title"), "Test")

    def test_clr_unicode_string_serializes_without_a_byte_decoder(self):
        class _TypeInfo(object):
            FullName = "System.String"

        class DotNetUnicodeString(str):
            def GetType(self):  # noqa: N802
                return _TypeInfo()

            def decode(self, encoding):
                raise AssertionError("A CLR System.String must never be decoded as bytes.")

        value = DotNetUnicodeString("Vendor \u2013 Series")
        payload = storage_service._json_dumps({"name": value})
        self.assertIn("\\u2013", payload.lower())

    def test_raw_bytes_fail_with_precise_set_element_and_field_context(self):
        tracking_set = models.new_tracking_set(
            "Unicode Regression Set",
            {"source_type": "host", "display_name": "Host"},
            {"id": -1, "name": "Mechanical Equipment"},
            [{"key": "project:1:instance", "name": "MCA \u2013 CED", "scope": "instance"}],
        )
        tracking_set["elements"] = {
            "host:test": {
                "persistent_id": "host:test",
                "metadata": {
                    "friendly_name": "Vendor Unit",
                    "element_id": 42,
                    "family_name": "Vendor \u00d8 Family",
                    "type_name": "Vendor Type",
                },
                "current_metadata": {
                    "friendly_name": "Vendor Unit",
                    "element_id": 42,
                    "family_name": "Vendor \u00d8 Family",
                    "type_name": "Vendor Type",
                },
                "current_properties": {
                    "project:1:instance": {"display": b"never guess this"},
                },
            },
        }
        store = models.new_project_store()
        store["tracking_sets"] = [tracking_set]
        document = type(
            "Document", (object,),
            {"Title": "Test", "PathName": "", "IsWorkshared": False},
        )()

        with self.assertRaises(storage_service.StorageError) as raised:
            storage_service.save(document, store)

        message = str(raised.exception)
        self.assertIn("Unicode Regression Set", message)
        self.assertIn("Vendor Unit", message)
        self.assertIn("Vendor Ø Family", message)
        self.assertIn("current_properties", message)
        self.assertIn("Unexpected byte text", message)

    def test_fallback_writer_creates_multiline_global_parameter(self):
        multiline = object()
        calls = []

        class FakeStringSpecs(object):
            MultilineText = multiline

        class FakeSpecTypeId(object):
            String = FakeStringSpecs

        class FakeManager(object):
            @staticmethod
            def AreGlobalParametersAllowed(document):
                return True

            @staticmethod
            def FindByName(document, name):
                return None

        class FakeStringValue(object):
            def __init__(self, value):
                self.Value = value

        class FakeParameter(object):
            def SetValue(self, value):
                calls.append(("value", value.Value))

        class FakeGlobalParameter(object):
            @staticmethod
            def Create(document, name, spec):
                calls.append(("create", name, spec))
                return FakeParameter()

        class FakeTransaction(object):
            def __init__(self, document, name):
                calls.append(("transaction", name))

            def Start(self):
                calls.append(("start",))

            def Commit(self):
                calls.append(("commit",))

            def RollBack(self):
                calls.append(("rollback",))

        self._replace("SpecTypeId", FakeSpecTypeId)
        self._replace("GlobalParametersManager", FakeManager)
        self._replace("GlobalParameter", FakeGlobalParameter)
        self._replace("StringParameterValue", FakeStringValue)
        self._replace("Transaction", FakeTransaction)
        document = type(
            "Document",
            (object,),
            {"IsModifiable": False, "IsFamilyDocument": False},
        )()

        storage_service._write_fallback_payload(document, "payload", "Backup")

        self.assertIn(
            ("create", storage_service.FALLBACK_PARAMETER_NAME, multiline),
            calls,
        )
        self.assertIn(("value", "payload"), calls)
        self.assertIn(("commit",), calls)
        self.assertNotIn(("rollback",), calls)

    def test_load_reads_backup_when_extensible_storage_is_unavailable(self):
        original = models.new_project_store()
        original["tracking_sets"] = [{"id": "fallback-set"}]
        payload = storage_service.serialize_payload(original)

        def fail_schema():
            raise RuntimeError("schema unavailable")

        self._replace("_get_schema", fail_schema)
        self._replace("_read_fallback_payload", lambda document: payload)
        document = type(
            "Document",
            (object,),
            {"Title": "Test", "PathName": "", "IsWorkshared": False},
        )()

        restored = storage_service.load(document)

        self.assertEqual(restored["tracking_sets"][0]["id"], "fallback-set")


if __name__ == "__main__":
    unittest.main()
