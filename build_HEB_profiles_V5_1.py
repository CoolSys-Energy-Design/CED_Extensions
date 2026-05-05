# -*- coding: utf-8 -*-
"""Build HEB_profiles_V5.1.yaml.

Goal: keep the V5 (schema 100) structure exactly, but inject the
keynote ``Keynote Value`` and ``Keynote Description`` fields that
were captured into ``HEB_profiles_V4_MODIFIED_41.yaml`` and silently
dropped during the V4 -> V5 unified-annotations migration.

V4 keynote shape (per LED, peer list ``keynotes:``)::

    keynotes:
      - parameters: null            # empty placeholder
        Keynote Value: 12           # SIBLING of `parameters` — the
        Keynote Description: '...'  # serializer flattened them out
        offsets: { ... }
        category_name: Generic Annotations
        family_name: GA_Keynote Symbol_CED
        type_name: Standard - Electrical

V5 keynote shape (unified ``annotations:`` list)::

    annotations:
      - id: SET-208-LED-004-ANN-002
        kind: keynote
        category_name: Generic Annotations
        family_name: GA_Keynote Symbol_CED
        type_name: Standard - Electrical
        parameters: null               # <-- Should hold KV/KD
        offsets: { ... }
        label: 'GA_Keynote Symbol_CED : Standard - Electrical'

Match key per keynote: (set_id, led_id, family_name, type_name,
offset signature). Set + LED ids survived the V4 -> V5 migration so
that's the strong cross-file linkage; family/type and offset
disambiguate when an LED carries more than one keynote (rare but
allowed).

Run from anywhere; reads / writes paths are absolute.
"""

import io
import sys
import os

import yaml


SCRIPT_DIR = r"c:\CED_Extensions"
V5_PATH = os.path.join(SCRIPT_DIR, "HEB_profiles_V5.yaml")
V4_PATH = os.path.join(SCRIPT_DIR, "HEB_profiles_V4_MODIFIED_41.yaml")
OUT_PATH = os.path.join(SCRIPT_DIR, "HEB_profiles_V5.1.yaml")


# ---------------------------------------------------------------------
# YAML helpers — preserve key order on round-trip
# ---------------------------------------------------------------------

def _safe_load(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _safe_dump(data, path):
    with io.open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data, f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
            width=10**9,  # don't auto-wrap long descriptions
        )


# ---------------------------------------------------------------------
# Offset signature — used to disambiguate multi-keynote LEDs
# ---------------------------------------------------------------------

def _offset_sig(offsets):
    """Stable rounded tuple representing an offset dict for keying."""
    if not isinstance(offsets, dict):
        return ("none",)
    def _r(v):
        try:
            return round(float(v or 0.0), 4)
        except (TypeError, ValueError):
            return 0.0
    return (
        _r(offsets.get("x_inches")),
        _r(offsets.get("y_inches")),
        _r(offsets.get("z_inches")),
        _r(offsets.get("rotation_deg")),
    )


# ---------------------------------------------------------------------
# V4 keynote index
# ---------------------------------------------------------------------

def build_v4_keynote_index(v4_data):
    """Return ``{(set_id, led_id, family, type, offset_sig): {KV, KD}}``.

    Falls back to (set_id, led_id, family, type) when the offset
    signature alone isn't unique — that's the typical case (one
    keynote per LED).
    """
    by_full_key = {}
    by_loose_key = {}  # without offset_sig
    for profile in (v4_data or {}).get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for s in profile.get("linked_sets") or []:
            if not isinstance(s, dict):
                continue
            sid = s.get("id") or ""
            for led in s.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                lid = led.get("id") or ""
                for kn in led.get("keynotes") or []:
                    if not isinstance(kn, dict):
                        continue
                    kv = kn.get("Keynote Value")
                    kd = kn.get("Keynote Description")
                    if kv is None and kd is None:
                        continue
                    family = kn.get("family_name") or ""
                    type_name = kn.get("type_name") or ""
                    offs = kn.get("offsets") or {}
                    full_key = (sid, lid, family, type_name, _offset_sig(offs))
                    loose_key = (sid, lid, family, type_name)
                    payload = {
                        "Keynote Value": kv,
                        "Keynote Description": kd,
                    }
                    by_full_key[full_key] = payload
                    by_loose_key.setdefault(loose_key, payload)
    return by_full_key, by_loose_key


# ---------------------------------------------------------------------
# Patch V5 in place
# ---------------------------------------------------------------------

def patch_v5_keynotes(v5_data, full_index, loose_index):
    patched = 0
    skipped_already_filled = 0
    no_match = 0
    no_match_examples = []
    keynote_total = 0
    for profile in (v5_data or {}).get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        pid = profile.get("id") or ""
        for s in profile.get("linked_sets") or []:
            if not isinstance(s, dict):
                continue
            sid = s.get("id") or ""
            for led in s.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                lid = led.get("id") or ""
                for ann in led.get("annotations") or []:
                    if not isinstance(ann, dict):
                        continue
                    if ann.get("kind") != "keynote":
                        continue
                    keynote_total += 1
                    family = ann.get("family_name") or ""
                    type_name = ann.get("type_name") or ""
                    offs = ann.get("offsets") or {}
                    full_key = (sid, lid, family, type_name, _offset_sig(offs))
                    loose_key = (sid, lid, family, type_name)

                    payload = full_index.get(full_key) or loose_index.get(loose_key)
                    if payload is None:
                        no_match += 1
                        if len(no_match_examples) < 5:
                            no_match_examples.append(
                                "profile={} set={} led={} ann={} family={} type={}".format(
                                    pid, sid, lid,
                                    ann.get("id") or "?",
                                    family, type_name,
                                )
                            )
                        continue

                    params = ann.get("parameters")
                    if not isinstance(params, dict):
                        params = {}
                    # Don't overwrite if the user has already filled it
                    # in (some V5.x build may have run a different
                    # patch); preserve what's there.
                    if (params.get("Keynote Value") not in (None, "") and
                            params.get("Keynote Description") not in (None, "")):
                        skipped_already_filled += 1
                        ann["parameters"] = params
                        continue
                    if payload.get("Keynote Value") is not None:
                        params.setdefault("Keynote Value",
                                          payload["Keynote Value"])
                    if payload.get("Keynote Description") is not None:
                        params.setdefault("Keynote Description",
                                          payload["Keynote Description"])
                    ann["parameters"] = params
                    patched += 1
    return {
        "keynote_total": keynote_total,
        "patched": patched,
        "skipped_already_filled": skipped_already_filled,
        "no_match": no_match,
        "no_match_examples": no_match_examples,
    }


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def main():
    print("Loading V5 from {}".format(V5_PATH))
    v5 = _safe_load(V5_PATH)
    if not isinstance(v5, dict):
        print("ERROR: V5 didn't parse to a mapping")
        sys.exit(1)

    print("Loading V4 from {}".format(V4_PATH))
    v4 = _safe_load(V4_PATH)
    if not isinstance(v4, dict):
        print("ERROR: V4 didn't parse to a mapping")
        sys.exit(1)

    print("Building V4 keynote index...")
    full_index, loose_index = build_v4_keynote_index(v4)
    print("  {} keynote(s) indexed (full keys), {} unique loose keys".format(
        len(full_index), len(loose_index),
    ))

    print("Patching V5...")
    stats = patch_v5_keynotes(v5, full_index, loose_index)
    print("  {} V5 keynote annotation(s) total".format(stats["keynote_total"]))
    print("  {} patched with V4 KV/KD".format(stats["patched"]))
    print("  {} already had values; left alone".format(
        stats["skipped_already_filled"]
    ))
    print("  {} had no V4 match".format(stats["no_match"]))
    for example in stats.get("no_match_examples") or ():
        print("    - {}".format(example))

    print("Writing {}".format(OUT_PATH))
    _safe_dump(v5, OUT_PATH)
    size = os.path.getsize(OUT_PATH)
    print("  wrote {} bytes".format(size))
    print("Done.")


if __name__ == "__main__":
    main()
