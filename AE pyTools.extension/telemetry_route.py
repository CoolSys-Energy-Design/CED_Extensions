# -*- coding: utf-8 -*-
"""CED telemetry routing helpers for ACC usage transfer."""

import getpass
import json
import os
import time
from collections import deque


ORG_NAME = "CoolSys"
PROJECT_NAME = "CED Content Collection"

PROJECT_FILES_RELATIVE_PATH = "Project Files"
USAGE_RELATIVE_PATH = os.path.join("Project Files", "03 Automations", "Usage")
ROUTE_KEY_RELATIVE_PATH = os.path.join("Project Files", "py", ".ced_usage_route_key.json")

EXPECTED_ROUTE_ID = "ced-telemetry-usage-v1"
STATE_FILE_NAME = ".ced_usage_route_status.json"
STATE_SCHEMA_VERSION = 1

DISCOVERY_TOP_LEVEL_NAMES = set(["acc", "dc", "accdocs"])
DISCOVERY_NAME_CONTAINS = "forma"
DISCOVERY_MAX_DEPTH = 8
DISCOVERY_MAX_DIRS = 5000
DISCOVERY_MAX_SECONDS = 4.0


def _norm(path):
    if not path:
        return ""
    try:
        return os.path.normcase(os.path.normpath(path))
    except Exception:
        return str(path or "")


def _utc_now():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _to_text(value):
    if value is None:
        return ""
    try:
        return str(value)
    except Exception:
        return repr(value)


def get_username():
    try:
        return getpass.getuser()
    except Exception:
        return os.environ.get("USERNAME", "UnknownUser")


def telemetry_source_folder():
    appdata = os.environ.get("APPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Roaming"))
    return os.path.join(appdata, "pyRevit", "Extensions", "CED_pyTelemetry")


def ensure_telemetry_source_folder():
    source_folder = telemetry_source_folder()
    if os.path.isdir(source_folder):
        return source_folder, True, None
    try:
        os.makedirs(source_folder)
        return source_folder, True, None
    except Exception as ex:
        return source_folder, False, ex


def state_file_path(source_folder=None):
    folder = source_folder or telemetry_source_folder()
    return os.path.join(folder, STATE_FILE_NAME)


def _read_json(path):
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r") as fp:
            return json.load(fp)
    except Exception:
        return {}


def _write_json(path, payload):
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with open(path, "w") as fp:
        json.dump(payload, fp, indent=2, sort_keys=True)


def load_state(source_folder=None):
    path = state_file_path(source_folder)
    payload = _read_json(path)
    if not isinstance(payload, dict):
        payload = {}
    payload.setdefault("schema_version", STATE_SCHEMA_VERSION)
    payload.setdefault("updated_utc", "")
    payload.setdefault("approved_root", "")
    payload.setdefault("approved_by", "")
    payload.setdefault("approved_updated_utc", "")
    return payload


def save_state(state_payload, source_folder=None):
    path = state_file_path(source_folder)
    state_payload = dict(state_payload or {})
    state_payload["schema_version"] = STATE_SCHEMA_VERSION
    state_payload["updated_utc"] = _utc_now()
    _write_json(path, state_payload)
    return path


def _load_route_key(route_key_path):
    result = {
        "exists": False,
        "valid_json": False,
        "route_id": "",
        "schema_version": "",
        "matches_expected_route": False,
        "error": "",
    }
    if not os.path.isfile(route_key_path):
        return result

    result["exists"] = True
    try:
        with open(route_key_path, "r") as fp:
            payload = json.load(fp)
        result["valid_json"] = True
        result["route_id"] = _to_text(payload.get("route_id", ""))
        result["schema_version"] = _to_text(payload.get("schema_version", ""))
        result["matches_expected_route"] = (result["route_id"] == EXPECTED_ROUTE_ID)
    except Exception as ex:
        result["error"] = _to_text(ex)
    return result


def _safe_listdirs(path):
    if not os.path.isdir(path):
        return []
    try:
        return sorted([
            name for name in os.listdir(path)
            if os.path.isdir(os.path.join(path, name))
        ])
    except Exception:
        return []


def _project_sibling_count(candidate_root):
    org_root = _norm(os.path.dirname(candidate_root))
    if not org_root or not os.path.isdir(org_root):
        return 0
    try:
        return len([
            name for name in os.listdir(org_root)
            if os.path.isdir(os.path.join(org_root, name))
        ])
    except Exception:
        return 0


def _top_level_anchor_dirs(base_dir):
    roots = []
    if not base_dir or not os.path.isdir(base_dir):
        return roots
    try:
        names = os.listdir(base_dir)
    except Exception:
        return roots

    for name in names:
        full = os.path.join(base_dir, name)
        if not os.path.isdir(full):
            continue
        low = _norm(name).lower()
        if low in DISCOVERY_TOP_LEVEL_NAMES or DISCOVERY_NAME_CONTAINS in low:
            normed = _norm(full)
            if normed and normed not in roots:
                roots.append(normed)
    return roots


def _discover_project_roots_under(anchor_dir):
    target = _norm(PROJECT_NAME).lower()
    found = []
    if not anchor_dir or not os.path.isdir(anchor_dir):
        return found

    started = time.time()
    scanned = 0
    queue = deque([(anchor_dir, 0)])

    while queue:
        if (time.time() - started) > DISCOVERY_MAX_SECONDS:
            break
        current_dir, depth = queue.popleft()
        if depth > DISCOVERY_MAX_DEPTH:
            continue

        try:
            names = os.listdir(current_dir)
        except Exception:
            continue

        for name in names:
            if (time.time() - started) > DISCOVERY_MAX_SECONDS:
                break
            child = os.path.join(current_dir, name)
            if not os.path.isdir(child):
                continue
            scanned += 1
            if scanned >= DISCOVERY_MAX_DIRS:
                break

            child_name = _norm(name).lower()
            if child_name == target:
                normed = _norm(child)
                if normed and normed not in found:
                    found.append(normed)
                continue

            if depth < DISCOVERY_MAX_DEPTH:
                try:
                    if os.path.islink(child):
                        continue
                except Exception:
                    pass
                queue.append((child, depth + 1))

        if scanned >= DISCOVERY_MAX_DIRS:
            break

    return found


def _discover_candidates_by_scope():
    drive_root = os.path.splitdrive(os.path.expanduser("~"))[0] + os.sep
    home_root = os.path.expanduser("~")

    anchors = _top_level_anchor_dirs(drive_root)
    if not anchors:
        anchors = _top_level_anchor_dirs(home_root)

    candidates = []
    for anchor in anchors:
        for candidate in _discover_project_roots_under(anchor):
            if candidate and candidate not in candidates:
                candidates.append(candidate)
    return candidates


def build_candidate_roots(source_folder=None, extra_roots=None):
    candidates = []

    for path in _discover_candidates_by_scope():
        normed = _norm(path)
        if normed and normed not in candidates:
            candidates.append(normed)

    state_payload = load_state(source_folder=source_folder)
    approved_root = _norm(state_payload.get("approved_root"))
    if approved_root and approved_root not in candidates:
        candidates.append(approved_root)

    for root in list(extra_roots or []):
        normed = _norm(root)
        if normed and normed not in candidates:
            candidates.append(normed)

    return candidates


def score_candidate(candidate_root, username=None, approved_root=None):
    username = _to_text(username or get_username()).strip()
    approved_root = _norm(approved_root)
    candidate_root = _norm(candidate_root)

    project_files_path = os.path.join(candidate_root, PROJECT_FILES_RELATIVE_PATH)
    usage_base_path = os.path.join(candidate_root, USAGE_RELATIVE_PATH)
    route_key_path = os.path.join(candidate_root, ROUTE_KEY_RELATIVE_PATH)

    root_exists = os.path.isdir(candidate_root)
    project_files_exists = os.path.isdir(project_files_path)
    usage_base_exists = os.path.isdir(usage_base_path)

    project_files_subfolders = _safe_listdirs(project_files_path)
    usage_subfolders = _safe_listdirs(usage_base_path)

    project_files_subfolder_count = len(project_files_subfolders)
    usage_subfolder_count = len(usage_subfolders)

    has_automations_subfolder = any(
        _norm(name).lower() == _norm("03 Automations").lower()
        for name in project_files_subfolders
    )
    project_files_only_automations = (
        project_files_subfolder_count == 1 and has_automations_subfolder
    )

    usage_only_current_user = False
    if usage_subfolder_count == 1 and username:
        usage_only_current_user = (
            _norm(usage_subfolders[0]).lower() == _norm(username).lower()
        )

    key_info = _load_route_key(route_key_path)
    sibling_count = _project_sibling_count(candidate_root)
    approved_root_match = bool(approved_root and candidate_root == approved_root)

    score = 0
    if root_exists:
        score += 20
    if project_files_exists:
        score += 20
    if has_automations_subfolder:
        score += 10
    if project_files_subfolder_count > 1:
        score += 20
    if project_files_subfolder_count >= 5:
        score += 10
    if project_files_only_automations:
        score -= 120

    if usage_base_exists:
        score += 35
    if usage_subfolder_count > 1:
        score += 30
    elif usage_subfolder_count == 1:
        if usage_only_current_user:
            score -= 80
        else:
            score -= 30
    if usage_subfolder_count >= 5:
        score += 10

    if key_info["exists"]:
        score += 8
    if key_info["valid_json"]:
        score += 4
    if key_info["matches_expected_route"]:
        score += 8

    if sibling_count > 1:
        score += 8

    if approved_root_match:
        score += 10

    return {
        "root": candidate_root,
        "score": score,
        "root_exists": root_exists,
        "project_files_exists": project_files_exists,
        "project_files_path": project_files_path,
        "project_files_subfolder_count": project_files_subfolder_count,
        "project_files_subfolders": project_files_subfolders,
        "has_automations_subfolder": has_automations_subfolder,
        "project_files_only_automations": project_files_only_automations,
        "usage_base_path": usage_base_path,
        "usage_base_exists": usage_base_exists,
        "usage_subfolder_count": usage_subfolder_count,
        "usage_subfolders": usage_subfolders,
        "usage_only_current_user": usage_only_current_user,
        "route_key_path": route_key_path,
        "key_info": key_info,
        "sibling_count": sibling_count,
        "approved_root_match": approved_root_match,
    }


def score_candidates(candidates, username=None, approved_root=None):
    scored = [
        score_candidate(root, username=username, approved_root=approved_root)
        for root in list(candidates or [])
    ]
    return sorted(scored, key=lambda item: item["score"], reverse=True)


def _is_viable(candidate):
    if not candidate:
        return False
    if not candidate.get("root_exists"):
        return False
    if not candidate.get("usage_base_exists"):
        return False
    if candidate.get("project_files_only_automations"):
        return False
    return True


def resolve_usage_route(username=None, source_folder=None, persist=True):
    username = _to_text(username or get_username()).strip()
    state_payload = load_state(source_folder=source_folder)
    approved_root = _norm(state_payload.get("approved_root"))

    candidates = build_candidate_roots(source_folder=source_folder)
    scored = score_candidates(candidates, username=username, approved_root=approved_root)
    viable = [item for item in scored if _is_viable(item)]

    best = scored[0] if scored else None
    second = scored[1] if len(scored) > 1 else None
    margin = (best["score"] - second["score"]) if (best and second) else (best["score"] if best else 0)

    resolved = None
    status = "not_found"
    reason = "no_candidates"

    if viable:
        if len(viable) == 1:
            resolved = viable[0]
            status = "resolved"
            reason = "single_viable_candidate"
        else:
            approved_candidate = None
            for item in viable:
                if item.get("approved_root_match"):
                    approved_candidate = item
                    break

            if approved_candidate and approved_candidate["score"] >= (viable[0]["score"] - 15):
                resolved = approved_candidate
                status = "resolved"
                reason = "approved_root_preferred"
            elif viable[0]["score"] >= 75 and margin >= 20:
                resolved = viable[0]
                status = "resolved"
                reason = "high_confidence_scoring"
            else:
                status = "ambiguous"
                reason = "multiple_viable_candidates"
    elif scored:
        status = "not_found"
        reason = "no_viable_candidates"

    resolved_root = resolved["root"] if resolved else ""
    usage_base_path = resolved["usage_base_path"] if resolved else ""

    result = {
        "status": status,
        "reason": reason,
        "resolved_root": resolved_root,
        "usage_base_path": usage_base_path,
        "username": username,
        "best_score": best["score"] if best else 0,
        "margin": margin,
        "candidate_count": len(scored),
        "viable_count": len(viable),
        "scored_candidates": scored,
        "state_file": state_file_path(source_folder=source_folder),
    }

    if persist:
        state_payload["last_resolution"] = {
            "status": status,
            "reason": reason,
            "resolved_root": resolved_root,
            "usage_base_path": usage_base_path,
            "candidate_count": len(scored),
            "viable_count": len(viable),
            "best_score": best["score"] if best else 0,
            "margin": margin,
            "updated_utc": _utc_now(),
        }
        state_payload["last_candidates"] = [{
            "root": item["root"],
            "score": item["score"],
            "root_exists": item["root_exists"],
            "usage_base_exists": item["usage_base_exists"],
            "project_files_subfolder_count": item["project_files_subfolder_count"],
            "usage_subfolder_count": item["usage_subfolder_count"],
            "usage_only_current_user": item["usage_only_current_user"],
            "project_files_only_automations": item["project_files_only_automations"],
            "key_match": bool(item["key_info"]["matches_expected_route"]),
        } for item in scored]
        save_state(state_payload, source_folder=source_folder)

    return result


def validate_acc_root(candidate_root):
    inspected = score_candidate(candidate_root, username=get_username())
    valid = bool(
        inspected["root_exists"]
        and inspected["project_files_exists"]
        and inspected["usage_base_exists"]
    )
    return valid, inspected


def set_manual_approved_root(acc_root, username=None, source_folder=None):
    root = _norm(acc_root)
    username = _to_text(username or get_username()).strip()
    valid, inspected = validate_acc_root(root)

    result = {
        "success": False,
        "root": root,
        "state_file": state_file_path(source_folder=source_folder),
        "reason": "",
        "inspected": inspected,
    }

    if not valid:
        result["reason"] = "selected_root_invalid"
        return result

    state_payload = load_state(source_folder=source_folder)
    state_payload["approved_root"] = root
    state_payload["approved_by"] = "manual"
    state_payload["approved_updated_utc"] = _utc_now()
    state_payload["last_manual_username"] = username
    save_state(state_payload, source_folder=source_folder)

    resolve_usage_route(username=username, source_folder=source_folder, persist=True)

    result["success"] = True
    result["reason"] = "manual_approval_saved"
    return result


def ensure_user_folder(resolved_root, username=None):
    username = _to_text(username or get_username()).strip()
    usage_base_path = os.path.join(_norm(resolved_root), USAGE_RELATIVE_PATH)
    if not os.path.isdir(usage_base_path):
        return {"ok": False, "created": False, "reason": "usage_base_missing", "path": usage_base_path}

    user_folder = os.path.join(usage_base_path, username)
    if os.path.isdir(user_folder):
        return {"ok": True, "created": False, "reason": "already_exists", "path": user_folder}

    try:
        os.mkdir(user_folder)
        return {"ok": True, "created": True, "reason": "created", "path": user_folder}
    except Exception as ex:
        return {"ok": False, "created": False, "reason": _to_text(ex), "path": user_folder}


def cleanup_stale_user_folders(approved_root, username=None, source_folder=None):
    username = _to_text(username or get_username()).strip()
    approved_root = _norm(approved_root)
    removed = []
    skipped = []

    for root in build_candidate_roots(source_folder=source_folder):
        root_norm = _norm(root)
        if not root_norm or root_norm == approved_root:
            continue
        user_folder = os.path.join(root_norm, USAGE_RELATIVE_PATH, username)
        if not os.path.isdir(user_folder):
            continue
        try:
            if os.listdir(user_folder):
                skipped.append(user_folder)
                continue
            os.rmdir(user_folder)
            removed.append(user_folder)
        except Exception:
            skipped.append(user_folder)

    return {"removed": removed, "skipped": skipped}


def record_transfer_state(status, username=None, resolved_root="", files_found=0, files_copied=0, files_failed=0, source_folder=None, note=""):
    state_payload = load_state(source_folder=source_folder)
    state_payload["last_transfer"] = {
        "status": _to_text(status),
        "username": _to_text(username or get_username()),
        "resolved_root": _to_text(resolved_root),
        "files_found": int(files_found or 0),
        "files_copied": int(files_copied or 0),
        "files_failed": int(files_failed or 0),
        "note": _to_text(note),
        "updated_utc": _utc_now(),
    }
    save_state(state_payload, source_folder=source_folder)
