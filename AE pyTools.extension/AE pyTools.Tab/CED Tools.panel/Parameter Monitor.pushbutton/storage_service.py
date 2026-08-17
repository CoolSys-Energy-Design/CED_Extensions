# -*- coding: utf-8 -*-
"""Dedicated Extensible Storage persistence with a global-parameter backup.

Parameter Monitor intentionally owns a separate fixed schema. It does not use
the generalized YAML/history schema used by MEPRFP tools, avoiding schema-name,
GUID, and payload-key collisions. The preferred host is Revit's
``Autodesk.Revit.DB.ExtensibleStorage.DataStorage``. If that API route is not
available or a write fails, the same JSON is stored in a project-wide Revit
global multiline-text parameter.
"""

from __future__ import print_function

import json

import models
import text_service

DB = None
System = None
AccessLevel = None
DataStorage = None
Entity = None
Schema = None
SchemaBuilder = None
FilteredElementCollector = None
Transaction = None
GlobalParameter = None
GlobalParametersManager = None
SpecTypeId = None
StringParameterValue = None
Array = None
Guid = None
Int32 = int
String = str

try:
    import clr
    clr.AddReference("RevitAPI")
    import System
    import Autodesk.Revit.DB as DB
    from Autodesk.Revit.DB import FilteredElementCollector, Transaction
    from Autodesk.Revit.DB.ExtensibleStorage import (
        AccessLevel,
        DataStorage,
        Entity,
        Schema,
        SchemaBuilder,
    )
    from System import Array, Guid, Int32, String
except Exception:
    pass

try:
    from Autodesk.Revit.DB import (
        GlobalParameter,
        GlobalParametersManager,
        SpecTypeId,
        StringParameterValue,
    )
except Exception:
    pass


# v2 (current). v1 (3C41541B-...) declared Vendor-level write access with
# VendorId "CEDT", which never matched the executing pyRevit add-in identity,
# so every ES write was rejected ("Writing of Entities of this Schema is not
# allowed to the current add-in") and data landed in the fallback global
# parameter. Schemas are immutable per GUID once registered in a session, so
# the fix is a new GUID with Public write access. v1 is still read for
# migration in case an entity ever was written by a matching add-in.
SCHEMA_GUID = "CBFFFBE2-5722-4544-905E-5CFA5F013E10"
SCHEMA_NAME = "CED_ParameterMonitor_v2"
LEGACY_SCHEMA_GUID = "3C41541B-82C2-4CD9-8E52-6B89E9B8E6F2"
DATASTORAGE_NAME = "CED Parameter Monitor"
FALLBACK_PARAMETER_NAME = "CED_Parameter_Monitor_Data"
VENDOR_ID = "CEDT"
OUTER_SCHEMA_VERSION = 1
FIELD_SCHEMA_VERSION = "SchemaVersion"
FIELD_TOOL_VERSION = "ToolVersion"
FIELD_PROJECT_IDENTITY = "ProjectIdentity"
FIELD_PAYLOAD_JSON = "PayloadJson"
WARNING_PAYLOAD_BYTES = 12 * 1024 * 1024
MAX_PAYLOAD_BYTES = 15 * 1024 * 1024
_DATASTORAGE_TYPE_NAMES = (
    "Autodesk.Revit.DB.ExtensibleStorage.DataStorage, RevitAPI",
)
# Once a document has loaded its fallback payload, keep using that backend for
# the lifetime of this modeless session. This avoids retrying a known-failing
# Extensible Storage transaction on every button click; projects that load
# from the fallback remain on the fallback until a future migration policy is
# implemented deliberately.
_FALLBACK_ONLY_DOCUMENTS = set()


class StorageError(RuntimeError):
    pass


class StorageCorruptionError(StorageError):
    pass


def _diagnostic_value(value, fallback=u"<unavailable>"):
    return text_service.diagnostic_text(value, fallback=fallback)


def _mapping_context(mapping, parent_context):
    """Add usable set/element/parameter identity to persistence errors."""
    if not isinstance(mapping, dict):
        return parent_context
    if mapping.get("set_id"):
        return u"{} > tracking set '{}'".format(
            parent_context, _diagnostic_value(mapping.get("name"), u"Unnamed")
        )
    if mapping.get("persistent_id"):
        metadata = (
            mapping.get("current_metadata") or mapping.get("metadata") or {}
        )
        return u"{} > element '{}' (id={}, family='{}', type='{}', persistent_id={})".format(
            parent_context,
            _diagnostic_value(metadata.get("friendly_name"), u"Unnamed"),
            _diagnostic_value(metadata.get("element_id"), u"?"),
            _diagnostic_value(metadata.get("family_name"), u"?"),
            _diagnostic_value(metadata.get("type_name"), u"?"),
            _diagnostic_value(mapping.get("persistent_id"), u"?"),
        )
    if mapping.get("key") and mapping.get("scope") and "name" in mapping:
        return u"{} > tracked parameter '{}' ({})".format(
            parent_context,
            _diagnostic_value(mapping.get("name"), u"Unnamed"),
            _diagnostic_value(mapping.get("scope"), u"unknown scope"),
        )
    return parent_context


def _unicode_safe_json_value(value, context=u"Project monitor data"):
    """Normalize text exactly once and reject unknown bytes with context.

    There is intentionally no CP1252/Latin-1/UTF-8 retry list here.  Revit API
    text is already Unicode; any raw byte string reaching persistence is a
    programming error and is reported with its tracking-set/element context.
    """
    if text_service.is_text_value(value):
        return text_service.to_text(value, context=context)
    if isinstance(value, dict):
        result = {}
        mapping_context = _mapping_context(value, context)
        for key, item in value.items():
            key_context = u"{} > field name".format(mapping_context)
            safe_key = (
                text_service.to_text(key, context=key_context)
                if text_service.is_text_value(key) else key
            )
            value_context = u"{} > field '{}'".format(
                mapping_context, _diagnostic_value(safe_key, u"<invalid key>")
            )
            result[safe_key] = _unicode_safe_json_value(item, value_context)
        return result
    if isinstance(value, (list, tuple)):
        return [
            _unicode_safe_json_value(item, u"{} > item {}".format(context, index))
            for index, item in enumerate(value)
        ]
    return value


def _document_session_key(document):
    return id(document) if document is not None else None


def _json_dumps(value, context=u"Project monitor data", already_normalized=False):
    normalized = (
        value if already_normalized
        else _unicode_safe_json_value(value, context=context)
    )
    return json.dumps(
        normalized,
        separators=(",", ":"),
        sort_keys=True,
        ensure_ascii=True,
    )


def serialize_payload(store):
    migrated = models.migrate_project_store(store)
    migrated["tool_version"] = models.TOOL_VERSION
    migrated["updated_at"] = models.utc_now_text()
    return _json_dumps(migrated, context=u"Serialized project store")


def deserialize_payload(text, project_identity=None):
    if not text:
        return models.new_project_store(project_identity=project_identity)
    try:
        raw = json.loads(text_service.to_text(text, context=u"Stored project payload"))
    except Exception as ex:
        raise StorageCorruptionError("Parameter Monitor payload JSON is invalid: {}".format(ex))
    return models.migrate_project_store(raw, project_identity=project_identity)


def _require_revit():
    if DB is None or Schema is None or Entity is None:
        raise StorageError("Revit API is required for project persistence.")


def project_identity(document):
    identity = {
        "title": text_service.to_text(
            getattr(document, "Title", "") or "", context=u"Project title"
        ),
        "path": text_service.to_text(
            getattr(document, "PathName", "") or "", context=u"Project path"
        ),
        "is_workshared": bool(getattr(document, "IsWorkshared", False)),
    }
    try:
        identity["central_guid"] = str(document.WorksharingCentralGUID)
    except Exception:
        pass
    try:
        cloud_path = document.GetCloudModelPath()
        if cloud_path is not None:
            identity["cloud_project_guid"] = str(cloud_path.GetProjectGUID())
            identity["cloud_model_guid"] = str(cloud_path.GetModelGUID())
    except Exception:
        pass
    return identity


def _get_schema():
    _require_revit()
    guid = Guid(SCHEMA_GUID)
    schema = Schema.Lookup(guid)
    if schema is not None:
        return schema
    builder = SchemaBuilder(guid)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.SetDocumentation(
        "CED Parameter Monitor stable envelope. Versioned monitor state is stored in PayloadJson."
    )
    try:
        builder.SetVendorId(VENDOR_ID)
        builder.SetReadAccessLevel(AccessLevel.Public)
        # Public write: Vendor-level write requires the executing add-in's
        # VendorId to match, and pyRevit's identity is not "CEDT".
        builder.SetWriteAccessLevel(AccessLevel.Public)
    except Exception:
        pass
    builder.AddSimpleField(FIELD_SCHEMA_VERSION, Int32)
    builder.AddSimpleField(FIELD_TOOL_VERSION, String)
    builder.AddSimpleField(FIELD_PROJECT_IDENTITY, String)
    builder.AddSimpleField(FIELD_PAYLOAD_JSON, String)
    return builder.Finish()


def _get_legacy_schema():
    """Look up the retired v1 schema for read-only migration. Never builds it."""
    if Schema is None or Guid is None:
        return None
    try:
        return Schema.Lookup(Guid(LEGACY_SCHEMA_GUID))
    except Exception:
        return None


def _resolve_datastorage_type():
    """Resolve the real Revit DataStorage type without trusting the DB proxy.

    Revit places DataStorage in ``Autodesk.Revit.DB.ExtensibleStorage``. Some
    Python/.NET hosts do not expose nested namespace types as attributes on the
    imported ``Autodesk.Revit.DB`` namespace proxy, which is why
    ``DB.DataStorage`` is intentionally never used here.
    """
    global DataStorage
    if DataStorage is not None:
        return DataStorage
    if System is None:
        return None
    for type_name in _DATASTORAGE_TYPE_NAMES:
        try:
            resolved = System.Type.GetType(type_name)
        except Exception:
            resolved = None
        if resolved is not None:
            DataStorage = resolved
            return DataStorage
    try:
        assemblies = System.AppDomain.CurrentDomain.GetAssemblies()
    except Exception:
        assemblies = []
    for assembly in assemblies:
        for full_name in ("Autodesk.Revit.DB.ExtensibleStorage.DataStorage",):
            try:
                resolved = assembly.GetType(full_name)
            except Exception:
                resolved = None
            if resolved is not None:
                DataStorage = resolved
                return DataStorage
    return None


def _collector_type():
    if FilteredElementCollector is not None:
        return FilteredElementCollector
    if DB is not None:
        try:
            return DB.FilteredElementCollector
        except Exception:
            pass
    return None


def _transaction_type():
    if Transaction is not None:
        return Transaction
    if DB is not None:
        try:
            return DB.Transaction
        except Exception:
            pass
    return None


def _find_data_storage(document, schema, data_storage_type=None):
    data_storage_type = data_storage_type or _resolve_datastorage_type()
    collector_type = _collector_type()
    if document is None or data_storage_type is None or collector_type is None:
        return None
    for storage in collector_type(document).OfClass(data_storage_type):
        try:
            entity = storage.GetEntity(schema)
            if entity is not None and entity.IsValid():
                return storage
        except Exception:
            continue
    return None


def _create_data_storage(document, data_storage_type=None):
    data_storage_type = data_storage_type or _resolve_datastorage_type()
    if data_storage_type is None:
        return None
    try:
        return data_storage_type.Create(document)
    except Exception:
        pass
    try:
        create_method = data_storage_type.GetMethod("Create")
    except Exception:
        create_method = None
    if create_method is None:
        return None
    try:
        if Array is not None and System is not None:
            return create_method.Invoke(None, Array[System.Object]([document]))
        return create_method.Invoke(None, [document])
    except Exception:
        return None


def _get_text(entity, field):
    try:
        return entity.Get[String](field)
    except Exception:
        try:
            return entity.Get[str](field)
        except Exception:
            return None


def _entity_payload(entity, schema, identity):
    if entity is None or not entity.IsValid():
        return None
    outer_version = 0
    try:
        outer_version = int(entity.Get[Int32](schema.GetField(FIELD_SCHEMA_VERSION)))
    except Exception:
        pass
    if outer_version > OUTER_SCHEMA_VERSION:
        raise StorageError(
            "Stored envelope schema {} is newer than supported schema {}.".format(
                outer_version, OUTER_SCHEMA_VERSION
            )
        )
    payload_text = _get_text(entity, schema.GetField(FIELD_PAYLOAD_JSON))
    return deserialize_payload(payload_text, project_identity=identity)


def _global_parameter_manager():
    if GlobalParametersManager is not None:
        return GlobalParametersManager
    if DB is not None:
        try:
            return DB.GlobalParametersManager
        except Exception:
            pass
    return None


def _global_parameter_type():
    if GlobalParameter is not None:
        return GlobalParameter
    if DB is not None:
        try:
            return DB.GlobalParameter
        except Exception:
            pass
    return None


def _string_parameter_value_type():
    if StringParameterValue is not None:
        return StringParameterValue
    if DB is not None:
        try:
            return DB.StringParameterValue
        except Exception:
            pass
    return None


def _find_fallback_global_parameter(document):
    manager = _global_parameter_manager()
    if document is None or manager is None:
        return None
    try:
        parameter_id = manager.FindByName(document, FALLBACK_PARAMETER_NAME)
    except Exception:
        return None
    if parameter_id is None:
        return None
    try:
        return document.GetElement(parameter_id)
    except Exception:
        return None


def _read_fallback_payload(document):
    parameter = _find_fallback_global_parameter(document)
    if parameter is None:
        return None
    try:
        value = parameter.GetValue()
    except Exception as ex:
        raise StorageError(
            "Could not read fallback global parameter {!r}: {}".format(
                FALLBACK_PARAMETER_NAME, ex
            )
        )
    expected_type = _string_parameter_value_type()
    if expected_type is not None and value is not None:
        try:
            is_string_value = isinstance(value, expected_type)
        except Exception:
            is_string_value = True
        if not is_string_value:
            raise StorageError(
                "Fallback global parameter {!r} is not a text parameter.".format(
                    FALLBACK_PARAMETER_NAME
                )
            )
    if value is None:
        return ""
    try:
        return text_service.to_text(
            value.Value or "", context=u"Fallback global parameter payload"
        )
    except Exception as ex:
        raise StorageError(
            "Fallback global parameter {!r} has no readable text value: {}".format(
                FALLBACK_PARAMETER_NAME, ex
            )
        )


def _multiline_text_spec():
    spec_type = SpecTypeId
    if spec_type is None and DB is not None:
        try:
            spec_type = DB.SpecTypeId
        except Exception:
            spec_type = None
    if spec_type is not None:
        try:
            return spec_type.String.MultilineText
        except Exception:
            try:
                return spec_type.String.Text
            except Exception:
                pass
    if DB is not None:
        try:
            return DB.ParameterType.MultilineText
        except Exception:
            try:
                return DB.ParameterType.Text
            except Exception:
                pass
    return None


def _global_parameters_allowed(document):
    manager = _global_parameter_manager()
    if manager is None:
        return False
    try:
        return bool(manager.AreGlobalParametersAllowed(document))
    except Exception:
        return not bool(getattr(document, "IsFamilyDocument", False))


def _write_fallback_payload(document, payload_text, transaction_name):
    manager = _global_parameter_manager()
    parameter_type = _global_parameter_type()
    value_type = _string_parameter_value_type()
    text_spec = _multiline_text_spec()
    transaction_type = _transaction_type()
    if manager is None or parameter_type is None or value_type is None or text_spec is None:
        raise StorageError("The Revit global text parameter API is unavailable.")
    if not _global_parameters_allowed(document):
        raise StorageError("Revit global parameters are not allowed in this document.")

    parameter = _find_fallback_global_parameter(document)
    transaction = None
    try:
        if not bool(getattr(document, "IsModifiable", False)):
            if transaction_type is None:
                raise StorageError("The Revit transaction API is unavailable.")
            transaction = transaction_type(
                document,
                str(transaction_name or "Parameter Monitor - Save Backup"),
            )
            transaction.Start()
        if parameter is None:
            parameter = parameter_type.Create(document, FALLBACK_PARAMETER_NAME, text_spec)
        parameter.SetValue(value_type(payload_text))
        if transaction is not None:
            transaction.Commit()
        return parameter
    except Exception:
        if transaction is not None:
            try:
                transaction.RollBack()
            except Exception:
                pass
        raise


def _delete_fallback_parameter(document, transaction_name):
    """Remove the stale fallback global parameter after a successful ES write.

    Returns True when a parameter was deleted. Best-effort: any failure is
    reported by the caller as a warning, never as a save failure.
    """
    parameter = _find_fallback_global_parameter(document)
    if parameter is None:
        return False
    transaction = None
    transaction_type = _transaction_type()
    try:
        if not bool(getattr(document, "IsModifiable", False)):
            if transaction_type is None:
                return False
            transaction = transaction_type(
                document,
                str(transaction_name or "Parameter Monitor - Remove Backup Parameter"),
            )
            transaction.Start()
        document.Delete(parameter.Id)
        if transaction is not None:
            transaction.Commit()
        return True
    except Exception:
        if transaction is not None:
            try:
                transaction.RollBack()
            except Exception:
                pass
        raise


def _log(logger, level, message, *args):
    if logger is None:
        return
    try:
        getattr(logger, level)(message, *args)
    except Exception:
        pass


def _log_save_context(logger, transaction_name, store):
    """Log concise Add Set/persistence context without dumping every element."""
    if logger is None:
        return
    sets = list((store or {}).get("tracking_sets") or [])
    _log(
        logger,
        "info",
        "Parameter Monitor save context: operation=%s, tracking_sets=%s.",
        transaction_name,
        len(sets),
    )
    for tracking_set in sets:
        properties = list(tracking_set.get("tracked_properties") or [])
        labels = [
            _diagnostic_value(item.get("name"), u"Unnamed parameter")
            for item in properties[:20]
        ]
        if len(properties) > 20:
            labels.append(u"… {} more".format(len(properties) - 20))
        _log(
            logger,
            "info",
            "Parameter Monitor set: name=%s, id=%s, elements=%s, properties=[%s].",
            _diagnostic_value(tracking_set.get("name"), u"Unnamed set"),
            _diagnostic_value(tracking_set.get("set_id"), u"?"),
            len(tracking_set.get("elements") or {}),
            u"; ".join(labels),
        )


def load(document, logger=None):
    identity = project_identity(document)
    extensible_error = None
    try:
        schema = _get_schema()
        storage = _find_data_storage(document, schema)
        if storage is not None:
            store = _entity_payload(storage.GetEntity(schema), schema, identity)
            if store is not None:
                _log(logger, "info", "Parameter Monitor loaded Extensible Storage data.")
                _FALLBACK_ONLY_DOCUMENTS.discard(_document_session_key(document))
                return store
        legacy_schema = _get_legacy_schema()
        if legacy_schema is not None:
            legacy_storage = _find_data_storage(document, legacy_schema)
            if legacy_storage is not None:
                store = _entity_payload(
                    legacy_storage.GetEntity(legacy_schema), legacy_schema, identity
                )
                if store is not None:
                    _log(
                        logger,
                        "info",
                        "Parameter Monitor loaded legacy v1 Extensible Storage data; "
                        "the next save migrates it to the v2 schema.",
                    )
                    _FALLBACK_ONLY_DOCUMENTS.discard(
                        _document_session_key(document)
                    )
                    return store
    except StorageCorruptionError:
        raise
    except Exception as ex:
        extensible_error = ex
        _log(
            logger,
            "warning",
            "Parameter Monitor Extensible Storage read failed; trying global parameter backup: %s",
            ex,
        )

    try:
        fallback_text = _read_fallback_payload(document)
        if fallback_text is not None:
            store = deserialize_payload(fallback_text, project_identity=identity)
            _FALLBACK_ONLY_DOCUMENTS.add(_document_session_key(document))
            _log(
                logger,
                "warning",
                "Parameter Monitor loaded data from fallback global parameter %s.",
                FALLBACK_PARAMETER_NAME,
            )
            return store
    except StorageCorruptionError:
        raise
    except Exception as fallback_error:
        if extensible_error is not None:
            raise StorageError(
                "Extensible Storage read failed ({}); fallback global parameter read failed ({}).".format(
                    extensible_error, fallback_error
                )
            )
        raise

    if extensible_error is not None:
        _log(
            logger,
            "warning",
            "Parameter Monitor is starting with an empty store because no readable backup exists.",
        )
    return models.new_project_store(project_identity=identity)


def payload_size_bytes(store):
    return len(serialize_payload(store).encode("utf-8"))


def _build_entity(schema, identity, payload_text):
    entity = Entity(schema)
    entity.Set[Int32](schema.GetField(FIELD_SCHEMA_VERSION), Int32(OUTER_SCHEMA_VERSION))
    entity.Set[String](schema.GetField(FIELD_TOOL_VERSION), String(models.TOOL_VERSION))
    entity.Set[String](
        schema.GetField(FIELD_PROJECT_IDENTITY),
        String(_json_dumps(identity)),
    )
    entity.Set[String](schema.GetField(FIELD_PAYLOAD_JSON), String(payload_text))
    return entity


def _write_extensible_storage(document, identity, payload_text, transaction_name):
    schema = _get_schema()
    data_storage_type = _resolve_datastorage_type()
    if data_storage_type is None:
        raise StorageError(
            "Autodesk.Revit.DB.ExtensibleStorage.DataStorage could not be resolved."
        )
    storage = _find_data_storage(document, schema, data_storage_type)
    transaction = None
    transaction_type = _transaction_type()
    try:
        if not bool(getattr(document, "IsModifiable", False)):
            if transaction_type is None:
                raise StorageError("The Revit transaction API is unavailable.")
            transaction = transaction_type(
                document,
                str(transaction_name or "Parameter Monitor - Save"),
            )
            transaction.Start()
        if storage is None:
            storage = _create_data_storage(document, data_storage_type)
            if storage is None:
                raise StorageError("Revit could not create a DataStorage element.")
            try:
                storage.Name = DATASTORAGE_NAME
            except Exception:
                pass
        storage.SetEntity(_build_entity(schema, identity, payload_text))
        if transaction is not None:
            transaction.Commit()
        return storage
    except Exception:
        if transaction is not None:
            try:
                transaction.RollBack()
            except Exception:
                pass
        raise


def save(document, store, transaction_name="Parameter Monitor - Save", logger=None):
    identity = project_identity(document)
    result = models.migrate_project_store(store, project_identity=identity)
    result["project_identity"] = identity
    result["tool_version"] = models.TOOL_VERSION
    result["updated_at"] = models.utc_now_text()
    _log_save_context(logger, transaction_name, result)
    try:
        # This is validation, not a decoding fallback. Revit-originated text
        # has already been converted at collection boundaries. If a raw byte
        # string still reaches this point, fail with exact model context.
        result = _unicode_safe_json_value(result, context=u"Project monitor save")
        payload_text = _json_dumps(
            result,
            context=u"Project monitor save",
            already_normalized=True,
        )
    except Exception as ex:
        _log(
            logger,
            "error",
            "Parameter Monitor text serialization failed during %s: %s",
            transaction_name,
            ex,
        )
        raise StorageError(
            "Parameter Monitor could not serialize text during {}: {}".format(
                transaction_name, ex
            )
        )
    payload_bytes = len(payload_text.encode("utf-8"))
    if payload_bytes > MAX_PAYLOAD_BYTES:
        raise StorageError(
            "Parameter Monitor data is {} bytes and exceeds the safe {} byte limit.".format(
                payload_bytes, MAX_PAYLOAD_BYTES
            )
        )
    if logger is not None and payload_bytes > WARNING_PAYLOAD_BYTES:
        try:
            logger.warning(
                "Parameter Monitor payload is approaching its storage limit: %s bytes",
                payload_bytes,
            )
        except Exception:
            pass

    extensible_error = None
    session_key = _document_session_key(document)
    fallback_only = session_key in _FALLBACK_ONLY_DOCUMENTS
    try:
        if fallback_only:
            raise StorageError(
                "Extensible Storage skipped because this document loaded the "
                "fallback backend for the current session."
            )
        _write_extensible_storage(document, identity, payload_text, transaction_name)
        _FALLBACK_ONLY_DOCUMENTS.discard(session_key)
        _log(
            logger,
            "info",
            "Parameter Monitor Extensible Storage write: %s bytes, %s set(s)",
            payload_bytes,
            len(result.get("tracking_sets") or []),
        )
        try:
            if _delete_fallback_parameter(
                document, "Parameter Monitor - Remove Backup Parameter"
            ):
                _log(
                    logger,
                    "info",
                    "Parameter Monitor removed the stale fallback global parameter %s "
                    "now that Extensible Storage writes succeed.",
                    FALLBACK_PARAMETER_NAME,
                )
        except Exception as cleanup_error:
            _log(
                logger,
                "warning",
                "Parameter Monitor could not remove the stale fallback global "
                "parameter %s: %s",
                FALLBACK_PARAMETER_NAME,
                cleanup_error,
            )
        return result
    except Exception as ex:
        extensible_error = ex
        _FALLBACK_ONLY_DOCUMENTS.add(session_key)
        _log(
            logger,
            "info" if fallback_only else "warning",
            "Parameter Monitor Extensible Storage %s; using global parameter backup: %s",
            "was skipped" if fallback_only else "write failed",
            ex,
        )

    try:
        _write_fallback_payload(
            document,
            payload_text,
            "Parameter Monitor - Save Global Backup",
        )
        _log(
            logger,
            "info" if fallback_only else "warning",
            "Parameter Monitor backup write: %s bytes in global multiline-text parameter %s.",
            payload_bytes,
            FALLBACK_PARAMETER_NAME,
        )
        return result
    except Exception as fallback_error:
        raise StorageError(
            "Parameter Monitor could not save. Extensible Storage failed ({}); "
            "global parameter backup failed ({}).".format(
                extensible_error, fallback_error
            )
        )
