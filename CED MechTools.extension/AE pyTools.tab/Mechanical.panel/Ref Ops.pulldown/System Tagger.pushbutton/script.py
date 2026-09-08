# -*- coding: utf-8 -*-
__title__ = "System Tagger"
__doc__ = "Place system ID tags on refrigerated cases from a pasted SYS NO. list."

import re

from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System import Guid, String, Int32
from pyrevit import revit, DB, forms, script

from Snippets import revit_helpers

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc


def _elid_value(item, default=0):
    return revit_helpers.get_elementid_value(item, default=default)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)

ORANGE_RGB = (255, 128, 0)
RESUME_SCHEMA_GUID = Guid("5f1a3c2e-9e4c-4e42-9f26-9c8f8f5c6f14")
RESUME_SCHEMA_NAME = "CED_SystemTagger_Resume"
RESUME_DATA_STORAGE_NAME = "CED_SystemTagger_Resume"

ACTION_MODE_TAG_AND_MARK = "Tag + Identity Mark"
ACTION_MODE_MARK_ONLY = "Identity Mark only"
ACTION_MODE_TAG_ONLY = "Tag only"


# Which equipment the picking loop will let you click, and which tag categories
# suit it. Multi-category tags always apply, so they are added on top.
CATEGORY_ORDER = ("mechanical", "electrical", "plumbing")
EQUIPMENT_CATEGORIES = {
    "mechanical": {
        "label": "Mechanical equipment",
        "equipment": ("OST_MechanicalEquipment",),
        "tags": ("OST_MechanicalEquipmentTags", "OST_SpecialityEquipmentTags"),
    },
    "electrical": {
        "label": "Electrical equipment",
        "equipment": ("OST_ElectricalEquipment",),
        "tags": ("OST_ElectricalEquipmentTags",),
    },
    "plumbing": {
        "label": "Plumbing equipment",
        "equipment": ("OST_PlumbingEquipment",),
        "tags": ("OST_PlumbingEquipmentTags",),
    },
}
ALWAYS_TAG_CATEGORIES = ("OST_MultiCategoryTags",)


def _builtin_category_ids(names):
    """Resolve BuiltInCategory names to ints, skipping any this Revit lacks."""
    ids = []
    for name in names or ():
        try:
            ids.append(int(getattr(DB.BuiltInCategory, name)))
        except Exception:
            continue
    return ids


class _CategorySelectionFilter(ISelectionFilter):
    """Allow only elements in the given BuiltInCategory ids."""

    def __init__(self, category_ids):
        self._allowed = set(int(i) for i in category_ids or ())

    def AllowElement(self, elem):  # noqa: N802
        if elem is None:
            return False
        try:
            cat = elem.Category
        except Exception:
            return False
        if cat is None:
            return False
        try:
            return _elid_value(cat.Id) in self._allowed
        except Exception:
            return False

    def AllowReference(self, reference, position):  # noqa: N802
        return False


def _rgb(r, g, b):
    return DB.Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])


def _get_resume_schema():
    schema = Schema.Lookup(RESUME_SCHEMA_GUID)
    if schema is not None:
        return schema
    sb = SchemaBuilder(RESUME_SCHEMA_GUID)
    sb.SetSchemaName(RESUME_SCHEMA_NAME)
    sb.SetReadAccessLevel(AccessLevel.Public)
    sb.SetWriteAccessLevel(AccessLevel.Public)
    sb.AddSimpleField("ExcelPath", String)
    sb.AddSimpleField("Index", Int32)
    return sb.Finish()


def _build_resume_entity(path, idx):
    schema = _get_resume_schema()
    entity = Entity(schema)
    try:
        entity.Set[String](schema.GetField("ExcelPath"), path or "")
    except Exception:
        pass
    try:
        entity.Set[Int32](schema.GetField("Index"), int(idx))
    except Exception:
        pass
    return entity


def _read_entity_values(entity, schema):
    if entity is None or not entity.IsValid():
        return None, None
    try:
        path = entity.Get[String](schema.GetField("ExcelPath"))
    except Exception:
        path = None
    try:
        idx = entity.Get[Int32](schema.GetField("Index"))
    except Exception:
        idx = None
    if idx is not None and idx < 0:
        idx = None
    return path or None, idx


def _find_resume_storage():
    if doc is None:
        return None
    schema = _get_resume_schema()
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.ExtensibleStorage.DataStorage)
    except Exception:
        return None
    for ds in collector:
        try:
            entity = ds.GetEntity(schema)
        except Exception:
            entity = None
        if entity and entity.IsValid():
            return ds
    return None


def _project_info_resume_entity():
    if doc is None:
        return None
    schema = _get_resume_schema()
    proj_info = doc.ProjectInformation
    if proj_info is None:
        return None
    try:
        entity = proj_info.GetEntity(schema)
    except Exception:
        return None
    if not entity or not entity.IsValid():
        return None
    return entity


def _migrate_from_project_info_if_needed():
    """One-time migration: if ProjectInformation has the legacy schema entity and
    no DataStorage entity exists yet, copy it across and wipe the old one.
    All in a single transaction. Returns True if a migration write occurred."""
    if doc is None:
        return False
    schema = _get_resume_schema()

    # If a DataStorage already holds the entity, nothing to do.
    if _find_resume_storage() is not None:
        return False

    legacy = _project_info_resume_entity()
    if legacy is None:
        return False

    path, idx = _read_entity_values(legacy, schema)
    proj_info = doc.ProjectInformation
    if proj_info is None:
        return False

    with revit.Transaction("System Tagger - Migrate Resume State"):
        ds = DB.ExtensibleStorage.DataStorage.Create(doc)
        try:
            ds.Name = RESUME_DATA_STORAGE_NAME
        except Exception:
            pass
        ds.SetEntity(_build_resume_entity(path, idx if idx is not None else -1))
        try:
            proj_info.DeleteEntity(schema)
        except Exception:
            # Best-effort: leave a tombstone so future loads ignore it.
            try:
                proj_info.SetEntity(_build_resume_entity("", -1))
            except Exception:
                pass
    logger.info(
        "System Tagger: migrated resume state from ProjectInformation to DataStorage."
    )
    return True


def _ensure_resume_storage():
    """Get or create the DataStorage element. Creation requires its own transaction."""
    ds = _find_resume_storage()
    if ds is not None:
        return ds
    with revit.Transaction("System Tagger - Create Resume Storage"):
        ds = DB.ExtensibleStorage.DataStorage.Create(doc)
        try:
            ds.Name = RESUME_DATA_STORAGE_NAME
        except Exception:
            pass
        ds.SetEntity(_build_resume_entity("", -1))
    return ds


def _load_resume_state():
    if doc is None:
        return None, None
    schema = _get_resume_schema()
    ds = _find_resume_storage()
    if ds is None:
        return None, None
    try:
        entity = ds.GetEntity(schema)
    except Exception:
        return None, None
    return _read_entity_values(entity, schema)


def _save_resume_state(path, idx):
    """Save resume state in its own transaction. Use at startup / for one-shot writes."""
    if doc is None:
        return
    ds = _ensure_resume_storage()
    if ds is None:
        return
    entity = _build_resume_entity(path, idx)
    with revit.Transaction("System Tagger - Save Resume State"):
        ds.SetEntity(entity)


def _save_resume_state_no_tx(path, idx):
    """Save resume state inside an already-open transaction. The DataStorage element
    must already exist (call _ensure_resume_storage() once before the picking loop)."""
    if doc is None:
        return
    ds = _find_resume_storage()
    if ds is None:
        # Fall back to opening our own tx so we don't silently drop state.
        _save_resume_state(path, idx)
        return
    ds.SetEntity(_build_resume_entity(path, idx))


def _clear_resume_state():
    _save_resume_state("", -1)


def _get_solid_fill_pattern_id():
    for fpe in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp.Target == DB.FillPatternTarget.Drafting and fp.IsSolidFill:
                return fpe.Id
        except Exception:
            continue
    return None


def _build_highlight_ogs():
    ogs = DB.OverrideGraphicSettings()
    col = _rgb(*ORANGE_RGB)
    try:
        ogs.SetProjectionLineColor(col)
    except Exception:
        pass
    try:
        ogs.SetCutLineColor(col)
    except Exception:
        pass
    pattern_id = _get_solid_fill_pattern_id()
    if pattern_id:
        try:
            ogs.SetSurfaceForegroundPatternId(pattern_id)
            ogs.SetSurfaceForegroundPatternColor(col)
        except Exception:
            pass
    return ogs


def _clear_override(view, elem_id):
    view.SetElementOverrides(elem_id, DB.OverrideGraphicSettings())


def _apply_override(view, elem_id, ogs):
    view.SetElementOverrides(elem_id, ogs)


def _apply_highlights(view, element_ids, ogs):
    if not element_ids:
        return
    for elem_int in element_ids:
        _apply_override(view, _elid_from_value(elem_int), ogs)


def _clear_highlights_in_tx(view, element_ids):
    if not element_ids:
        return
    clear_ogs = DB.OverrideGraphicSettings()
    for elem_int in element_ids:
        view.SetElementOverrides(_elid_from_value(elem_int), clear_ogs)


def _update_preview_highlights(view, ogs, new_ids, prev_ids):
    try:
        with revit.Transaction("System Tagger - Preview Highlights"):
            _clear_highlights_in_tx(view, prev_ids)
            _apply_highlights(view, new_ids, ogs)
        return list(new_ids or [])
    except Exception as ex:
        logger.warning("Preview highlight update failed: {}".format(ex))
        return list(prev_ids or [])







def _get_element_center(elem, view):
    if not elem:
        return None
    bbox = None
    try:
        bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _index_to_letters(index):
    letters = ""
    num = index
    while True:
        num, rem = divmod(num, 26)
        letters = chr(ord("A") + rem) + letters
        if num == 0:
            break
        num -= 1
    return letters


def _collect_tag_types(category_names=None):
    """Tag family symbols for the given tag categories (all of them if None)."""
    if category_names is None:
        category_names = list(ALWAYS_TAG_CATEGORIES)
        for spec in EQUIPMENT_CATEGORIES.values():
            category_names.extend(spec["tags"])
    tag_types = []
    for cat_name in category_names:
        try:
            cat = getattr(DB.BuiltInCategory, cat_name)
        except Exception:
            cat = None
        if cat is None:
            continue
        tag_types.extend(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilySymbol)
            .OfCategory(cat)
            .ToElements()
        )
    return tag_types


def _tag_type_label(tag_type):
    try:
        fam_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    except Exception:
        fam_name = None
    try:
        type_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = None
    try:
        cat_name = tag_type.Category.Name if tag_type.Category else "Tag"
    except Exception:
        cat_name = "Tag"
    return "[{}] {} : {}".format(cat_name, fam_name or "?", type_name or "?")


def _prompt_settings(prefill_ids="", resume_note=None):
    """Ask for everything up front: IDs, category, action and tag type.

    Replaces the old chain (CommandSwitchWindow for the mode, SelectFromList for
    the tag, a paste window for the IDs) with one window, so the choices can be
    seen and corrected together.
    """
    from UIClasses import settings_form as sf

    tag_types = _collect_tag_types()
    label_to_type = {}
    for tag_type in tag_types:
        label_to_type[_tag_type_label(tag_type)] = tag_type
    tag_options = [(label, label) for label in sorted(label_to_type)]

    fields = [
        sf.multiline(
            "ids", "System IDs (paste from Excel)", default=prefill_ids or "",
            hint="Paste the SYS NO. column - the whole block is fine. "
                 "The order you paste is the order you will be walked through.",
            height=110),
        sf.choice(
            "category", "Equipment category",
            [(key, EQUIPMENT_CATEGORIES[key]["label"]) for key in CATEGORY_ORDER],
            default="mechanical",
            hint="Only this category can be clicked during the picking loop."),
        sf.choice(
            "mode", "What should it write?",
            [(ACTION_MODE_TAG_AND_MARK, ACTION_MODE_TAG_AND_MARK),
             (ACTION_MODE_MARK_ONLY, ACTION_MODE_MARK_ONLY),
             (ACTION_MODE_TAG_ONLY, ACTION_MODE_TAG_ONLY)],
            default=ACTION_MODE_TAG_AND_MARK),
        sf.choice(
            "tag", "Tag type", tag_options, required=False,
            hint="Not needed if you chose Identity Mark only."),
    ]
    if resume_note:
        fields.insert(0, sf.info(resume_note))

    def _validate(values):
        problems = []
        mode = values.get("mode")
        picked_label = values.get("tag")
        if mode != ACTION_MODE_MARK_ONLY:
            if not picked_label:
                if not tag_options:
                    problems.append(
                        "No tag types are loaded in this project, so only "
                        "'Identity Mark only' can run.")
                else:
                    problems.append(
                        "Pick a tag type, or switch to 'Identity Mark only'.")
            else:
                spec = EQUIPMENT_CATEGORIES.get(values.get("category")) or {}
                allowed = set(_builtin_category_ids(
                    tuple(ALWAYS_TAG_CATEGORIES) + tuple(spec.get("tags") or ())))
                tag_type = label_to_type.get(picked_label)
                try:
                    tag_cat = _elid_value(tag_type.Category.Id)
                except Exception:
                    tag_cat = None
                if tag_cat is not None and allowed and tag_cat not in allowed:
                    problems.append(
                        "That tag type does not suit {0}. Choose a multi-category "
                        "tag or one made for that category.".format(
                            spec.get("label", "the chosen category")))
        if not _parse_system_ids(values.get("ids") or ""):
            problems.append("No System IDs could be read from the pasted text.")
        return problems

    values = sf.prompt(
        "System Tagger",
        fields,
        subtitle="Then you will click the equipment for each ID in turn - "
                 "click again to unselect, Esc to move to the next ID.",
        ok_label="Start picking",
        caller_file=__file__,
        validate=_validate,
    )
    if values is None:
        return None
    spec = EQUIPMENT_CATEGORIES[values["category"]]
    return {
        "raw_text": values["ids"],
        "system_ids": _parse_system_ids(values["ids"]),
        "action_mode": values["mode"],
        "tag_type": label_to_type.get(values.get("tag") or ""),
        "category_key": values["category"],
        "pick_label": spec["label"],
        "pick_filter": _CategorySelectionFilter(
            _builtin_category_ids(spec["equipment"])),
    }


def _parse_system_ids(raw_text):
    system_ids = []
    if not raw_text:
        return system_ids
    raw = raw_text.replace("\r", "\n").strip()
    if not raw:
        return system_ids
    tokens = [t.strip() for t in re.split(r"[\t,\n;]+", raw) if t.strip()]
    for token in tokens:
        lowered = token.lower()
        if lowered in ("system id", "sys no.", "sys no", "sys no#", "sys no #"):
            continue
        system_ids.append(token)
    return system_ids


def _toggle_pick_cases(system_id, view, ogs, pick_filter, pick_label):
    selected_ids = []
    selected_set = set()
    preview_ids = []
    preview_group = DB.TransactionGroup(doc, "System Tagger - Preview Highlights")
    preview_group.Start()

    try:
        while True:
            with forms.WarningBar(
                title="System ID {}: pick {} (ESC to finish)".format(system_id, pick_label)
            ):
                while True:
                    try:
                        ref = uidoc.Selection.PickObject(
                            ObjectType.Element,
                            pick_filter,
                            "Pick {} (click again to unselect). ESC when done.".format(pick_label)
                        )
                    except Exception:
                        break

                    elem_int = _elid_value(ref.ElementId)
                    elem_id = _elid_from_value(elem_int)

                    if elem_int in selected_set:
                        selected_set.remove(elem_int)
                        selected_ids = [i for i in selected_ids if i != elem_int]
                    else:
                        selected_set.add(elem_int)
                        selected_ids.append(elem_int)

                    preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)

            choice = forms.CommandSwitchWindow.show(
                ["Done", "Pick More", "Pause"],
                message="System ID '{}': {} case(s) selected.".format(system_id, len(selected_ids)),
            )
            if choice == "Pick More":
                continue
            if choice == "Pause":
                preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
                _update_preview_highlights(view, ogs, [], preview_ids)
                try:
                    preview_group.Assimilate()
                except Exception:
                    pass
                return "pause", selected_ids
            if choice == "Cancel" or choice is None:
                preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
                _update_preview_highlights(view, ogs, [], preview_ids)
                try:
                    preview_group.Assimilate()
                except Exception:
                    pass
                return "cancel", selected_ids
            preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
            try:
                preview_group.Assimilate()
            except Exception:
                pass
            return "done", selected_ids
    finally:
        try:
            if preview_group.HasStarted():
                preview_group.RollBack()
        except Exception:
            pass
        try:
            preview_group.Dispose()
        except Exception:
            pass


def _build_labels(system_id, element_ids):
    if not element_ids:
        return []
    if len(element_ids) == 1:
        labels = [system_id]
    else:
        labels = [system_id + _index_to_letters(i) for i in range(len(element_ids))]
    return labels

def _apply_identity_mark(element_ids, labels):
    if not element_ids or not labels:
        return
    for elem_int, label in zip(element_ids, labels):
        elem = doc.GetElement(_elid_from_value(elem_int))
        if not elem:
            continue
        try:
            param = elem.LookupParameter("Identity Mark")
            if not param:
                param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
            if param and not param.IsReadOnly:
                if param.StorageType == DB.StorageType.String:
                    param.Set(label)
                else:
                    param.SetValueString(label)
        except Exception as ex:
            logger.warning(
                "Failed to set Identity Mark for {}: {}".format(elem_int, ex)
            )


def _place_tags(element_ids, labels, view, tag_type):
    if not element_ids or not labels or tag_type is None:
        return 0
    if not tag_type.IsActive:
        tag_type.Activate()
        doc.Regenerate()
    count = 0
    for elem_int, label in zip(element_ids, labels):
        elem = doc.GetElement(_elid_from_value(elem_int))
        if not elem:
            continue
        pt = _get_element_center(elem, view)
        if not pt:
            logger.warning("No bounding box for element {}".format(elem_int))
            continue
        try:
            reference = DB.Reference(elem)
            DB.IndependentTag.Create(
                doc,
                tag_type.Id,
                view.Id,
                reference,
                False,
                DB.TagOrientation.Horizontal,
                pt,
            )
            count += 1
        except Exception as ex:
            logger.warning("Failed to place tag for {}: {}".format(elem_int, ex))
    return count


def main():
    active_view = revit.active_view
    if active_view.IsTemplate:
        forms.alert("Active view is a template. Open a working view first.", exitscript=True)

    # One-time migration from legacy ProjectInformation storage to DataStorage.
    _migrate_from_project_info_if_needed()

    # Decide about a paused run first, so the settings form can be pre-filled
    # with the saved list.
    resume_path, resume_index = _load_resume_state()
    prefill_ids = ""
    resume_note = None
    start_index = 0

    if resume_path and resume_index is not None:
        resume_ids = _parse_system_ids(resume_path)
        if resume_ids and 0 <= resume_index < len(resume_ids):
            next_id = resume_ids[resume_index]
            choice = forms.CommandSwitchWindow.show(
                ["Resume", "Start Over", "Cancel"],
                message="Resume System Tagger at '{}'? ({} of {})".format(
                    next_id, resume_index + 1, len(resume_ids)
                ),
            )
            if choice == "Resume":
                prefill_ids = resume_path
                start_index = resume_index
                resume_note = (
                    "Resuming at '{0}' - ID {1} of {2}. The list below is the saved "
                    "one; editing it starts again from the top."
                ).format(next_id, resume_index + 1, len(resume_ids))
            elif choice == "Start Over":
                _clear_resume_state()
            else:
                script.exit()
        else:
            _clear_resume_state()

    settings = _prompt_settings(prefill_ids=prefill_ids, resume_note=resume_note)
    if settings is None:
        script.exit()

    raw_text = settings["raw_text"]
    system_ids = settings["system_ids"]
    action_mode = settings["action_mode"]
    tag_type = settings["tag_type"]
    pick_filter = settings["pick_filter"]
    pick_label = settings["pick_label"]

    # If the saved list was edited, the saved position no longer means anything.
    if prefill_ids and raw_text.strip() != prefill_ids.strip():
        start_index = 0

    highlight_ogs = _build_highlight_ogs()

    if action_mode == ACTION_MODE_MARK_ONLY:
        action_text = "Identity Mark will be set on each picked element (no tag will be placed)."
    elif action_mode == ACTION_MODE_TAG_ONLY:
        action_text = "A tag will be placed on each picked element (Identity Mark will not be changed)."
    else:
        action_text = "Identity Mark will be set and a tag placed on each picked element."

    forms.alert(
        "Select Mechanical Equipment for each System ID.\n"
        "- Click an element again to unselect.\n"
        "- Press ESC when done selecting, then click Done.\n"
        "- Click Pause to save progress and exit.\n"
        "The script will advance to the next System ID.\n\n"
        + action_text
    )

    # One-shot save of starting index. Also ensures the DataStorage element exists
    # so subsequent saves can ride inside the placement transaction (no extra tx).
    _save_resume_state(raw_text, start_index)

    for idx in range(start_index, len(system_ids)):
        system_id = system_ids[idx]
        status, picked_ids = _toggle_pick_cases(system_id, active_view, highlight_ogs,
                                                   pick_filter, pick_label)
        if status == "cancel":
            return
        try:
            labels = _build_labels(system_id, picked_ids)
            with revit.Transaction("System Tagger - Place Labels {}".format(system_id)):
                _apply_highlights(active_view, picked_ids, highlight_ogs)
                if action_mode != ACTION_MODE_TAG_ONLY:
                    _apply_identity_mark(picked_ids, labels)
                if action_mode != ACTION_MODE_MARK_ONLY and tag_type is not None:
                    _place_tags(picked_ids, labels, active_view, tag_type)
                _clear_highlights_in_tx(active_view, picked_ids)
                _save_resume_state_no_tx(raw_text, idx + 1)
        except Exception:
            raise
        if status == "pause":
            forms.alert("System Tagger paused. Run again to resume.")
            return

    _clear_resume_state()


if __name__ == "__main__":
    main()
