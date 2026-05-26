# profile_engine.py
# True profile-driven placement engine.
# Iterates yaml_profiles.json + Charlotte triggers, places fixtures with offsets + yaml params.

import math, json
from Autodesk.Revit import DB
from Autodesk.Revit.DB import UnitUtils, UnitTypeId

PARAM_UNITS = {
    'Apparent Load Input_CED': UnitTypeId.VoltAmperes,
    'Wattage Input_CED':       UnitTypeId.Watts,
    'Voltage_CED':             UnitTypeId.Volts,
    'FLA Input_CED':           UnitTypeId.Amperes,
    'CKT_Rating_CED':          UnitTypeId.Amperes,
}


def apply_yaml_params(elem, params):
    """Apply param dict to a Revit element, using UnitUtils for electrical doubles."""
    for pname, val in params.items():
        if val is None: continue
        p = elem.LookupParameter(pname)
        if p is None or p.IsReadOnly: continue
        try:
            st = p.StorageType
            if st == DB.StorageType.String:
                p.Set(str(val))
            elif st == DB.StorageType.Integer:
                p.Set(int(val))
            elif st == DB.StorageType.Double:
                unit = PARAM_UNITS.get(pname)
                internal = UnitUtils.ConvertToInternalUnits(float(val), unit) if unit else float(val)
                p.Set(internal)
        except:
            pass


def rotate_offset(offset_x_ft, offset_y_ft, anchor_rot_deg):
    """Rotate an offset by anchor rotation. Returns (dx, dy) in world coords."""
    a = math.radians(anchor_rot_deg)
    c, s = math.cos(a), math.sin(a)
    dx = c*offset_x_ft - s*offset_y_ft
    dy = s*offset_x_ft + c*offset_y_ft
    return dx, dy


def place_from_profile(doc, level, find_family_symbol, place_workplane_fixture,
                       set_mounting_height_inches, profile_name, profile_data,
                       triggers_for_profile, default_mt_in=18, room_suffix=None):
    """
    Place all linked-element fixtures from one yaml profile at all trigger positions.

    Args:
        doc, level: Revit doc + L1 level.
        find_family_symbol, place_workplane_fixture, set_mounting_height_inches: helpers from place_fixture.py.
        profile_name: name of the yaml profile (e.g. 'BCS_Tanning Bed 42.4').
        profile_data: dict from yaml_profiles.json[profile_name].
        triggers_for_profile: list of (anchor_x_ft, anchor_y_ft, anchor_rot_deg).
        default_mt_in: fall-back mounting height if profile doesn't specify.
        room_suffix: optional suffix to append to CKT_Load Name_CEDT.

    Returns: list of placed FamilyInstance ids.
    """
    placed_ids = []
    if not triggers_for_profile: return placed_ids

    fixtures = profile_data.get('fixtures', [])
    for anchor_x, anchor_y, anchor_rot in triggers_for_profile:
        for f in fixtures:
            family_name = f['family']
            type_name   = f['type']
            offset_x_ft = (f.get('offset_x_in', 0) or 0) / 12.0
            offset_y_ft = (f.get('offset_y_in', 0) or 0) / 12.0
            elem_rot_deg = f.get('rotation_deg', 0) or 0

            # Apply anchor rotation to offset, then add to anchor position
            dx, dy = rotate_offset(offset_x_ft, offset_y_ft, anchor_rot)
            wx = anchor_x + dx
            wy = anchor_y + dy

            # Final fixture rotation (in radians for place_workplane_fixture)
            final_rot_rad = math.radians(anchor_rot + elem_rot_deg)

            # Lookup symbol
            try:
                sym = find_family_symbol(doc, family_name, type_name)
            except:
                continue

            # Place
            pt = DB.XYZ(wx, wy, level.Elevation)
            try:
                inst = place_workplane_fixture(doc, sym, level, pt, rotation_rad=final_rot_rad)
            except:
                continue
            set_mounting_height_inches(inst, default_mt_in)

            # Apply yaml params
            params = dict(f.get('params', {}))
            if room_suffix and 'CKT_Load Name_CEDT' in params and params['CKT_Load Name_CEDT']:
                params['CKT_Load Name_CEDT'] = '%s - %s' % (params['CKT_Load Name_CEDT'], room_suffix)
            apply_yaml_params(inst, params)

            placed_ids.append(inst.Id.IntegerValue)
    return placed_ids


def place_all_profiles(doc, level, find_family_symbol, place_workplane_fixture,
                       set_mounting_height_inches, profiles_json_path, triggers_dict,
                       suffix_for_profile=None):
    """Iterate every profile in yaml_profiles.json, place at its trigger positions.

    suffix_for_profile: optional callable(profile_name, anchor_idx) -> room_suffix string.
    """
    with open(profiles_json_path, 'r') as f:
        profiles = json.load(f)

    total = 0
    per_profile_counts = {}
    for prof_name, prof_data in profiles.items():
        trigs = triggers_dict.get(prof_name, [])
        if not trigs:
            continue
        for idx, tr in enumerate(trigs):
            suffix = None
            if suffix_for_profile:
                suffix = suffix_for_profile(prof_name, idx)
            ids = place_from_profile(doc, level, find_family_symbol, place_workplane_fixture,
                                     set_mounting_height_inches, prof_name, prof_data, [tr],
                                     room_suffix=suffix)
            total += len(ids)
            per_profile_counts[prof_name] = per_profile_counts.get(prof_name, 0) + len(ids)
    return total, per_profile_counts
