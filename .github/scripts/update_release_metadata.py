#!/usr/bin/env python3
"""Update toolbar_version/build fields in about.yaml deterministically."""

import argparse
import io
import os
import re
import sys


WIP_VERSION_RE = re.compile(r"^(\d+\.\d+\.\d+)(?:-WIP)?$")


def _read_lines(path):
    with io.open(path, "r", encoding="utf-8") as handle:
        return handle.read().splitlines()


def _write_lines(path, lines):
    text = "\n".join(lines).rstrip("\n") + "\n"
    with io.open(path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def _find_key_index(lines, key_name):
    pattern = re.compile(r"^" + re.escape(key_name) + r"\s*:")
    for idx, line in enumerate(lines):
        if pattern.match(line):
            return idx
    return -1


def _get_key_value(lines, key_name):
    idx = _find_key_index(lines, key_name)
    if idx < 0:
        return ""
    value = lines[idx].split(":", 1)[1].split("#", 1)[0].strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        value = value[1:-1].strip()
    return value


def _set_or_insert_key(lines, key_name, value, insert_after_key=None):
    new_line = "{}: {}".format(key_name, value)
    idx = _find_key_index(lines, key_name)
    changed = False

    if idx >= 0:
        if lines[idx] != new_line:
            lines[idx] = new_line
            changed = True
        return changed

    insert_at = 0
    if insert_after_key:
        after_idx = _find_key_index(lines, insert_after_key)
        if after_idx >= 0:
            insert_at = after_idx + 1
    lines.insert(insert_at, new_line)
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--build", required=True)
    parser.add_argument("--toolbar-version", default="")
    parser.add_argument(
        "--mark-wip",
        action="store_true",
        help="Preserve the numeric version and append the -WIP prerelease marker.",
    )
    args = parser.parse_args()

    file_path = args.file
    build_value = str(args.build).strip()
    toolbar_version = str(args.toolbar_version or "").strip()

    if not os.path.exists(file_path):
        print("Missing file: {}".format(file_path), file=sys.stderr)
        return 1

    lines = _read_lines(file_path)
    changed = False

    if args.mark_wip and toolbar_version:
        print(
            "--mark-wip and --toolbar-version cannot be used together.",
            file=sys.stderr,
        )
        return 1

    if args.mark_wip:
        current_version = _get_key_value(lines, "toolbar_version")
        version_match = WIP_VERSION_RE.match(current_version)
        if not version_match:
            print(
                "Cannot mark invalid toolbar_version as WIP: {}".format(
                    current_version or "<missing>"
                ),
                file=sys.stderr,
            )
            return 1
        toolbar_version = "{}-WIP".format(version_match.group(1))

    if toolbar_version:
        changed = _set_or_insert_key(lines, "toolbar_version", toolbar_version) or changed

    changed = _set_or_insert_key(lines, "build", build_value, insert_after_key="toolbar_version") or changed

    if changed:
        _write_lines(file_path, lines)
        print("updated")
    else:
        print("unchanged")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
