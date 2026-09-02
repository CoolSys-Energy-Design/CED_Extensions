# -*- coding: utf-8 -*-
"""Presentation-independent state helpers for Revision Cloud Manager."""


def text(value, default=""):
    if value is None:
        return default
    try:
        return str(value)
    except Exception:
        return default


def normalized(value):
    return text(value).strip().lower()


class BadgeItem(object):
    def __init__(self, value, severity="Info"):
        self.text = text(value)
        self.severity = text(severity, "Info")


class NavigationItem(object):
    def __init__(self, value, target_id=0):
        self.text = text(value)
        self.target_id = int(target_id or 0)


class CloudRow(object):
    """One canonical revision-cloud row at the modeless UI boundary."""

    def __init__(self, **values):
        self.cloud_id = int(values.get("cloud_id") or 0)
        self.revision_id = int(values.get("revision_id") or 0)
        self.owner_view_id = int(values.get("owner_view_id") or 0)
        self.revision_number = text(values.get("revision_number"), "N/A")
        self.revision_display = text(values.get("revision_display"), self.revision_number)
        self.revision_sort = int(values.get("revision_sort") or 0)
        self.revision_date = text(values.get("revision_date"))
        self.revision_description = text(values.get("revision_description"))
        self.sheet_ids = list(values.get("sheet_ids") or [])
        self.sheet_numbers_list = list(values.get("sheet_numbers_list") or [])
        self.sheet_names_list = list(values.get("sheet_names_list") or [])
        self.sheet_numbers = ", ".join([text(x) for x in self.sheet_numbers_list if text(x)]) or "N/A"
        self.sheet_names = ", ".join([text(x) for x in self.sheet_names_list if text(x)]) or "N/A"
        self.sheet_number_links = [
            NavigationItem(number, sheet_id)
            for number, sheet_id in zip(self.sheet_numbers_list, self.sheet_ids)
            if text(number)
        ] or [NavigationItem("N/A", 0)]
        self.sheet_name_links = [
            NavigationItem(name, sheet_id)
            for name, sheet_id in zip(self.sheet_names_list, self.sheet_ids)
            if text(name)
        ] or [NavigationItem("N/A", 0)]
        # Both link lists always hold at least one item (the "N/A" fallback above),
        # so index 0 is always safe. These let the review grid render the common
        # one-sheet row as a single TextBlock instead of building an ItemsControl
        # and WrapPanel per cell; the multi-sheet path is unchanged.
        self.has_multiple_sheet_numbers = len(self.sheet_number_links) > 1
        self.has_multiple_sheet_names = len(self.sheet_name_links) > 1
        self.single_sheet_number = self.sheet_number_links[0].text
        self.single_sheet_number_id = self.sheet_number_links[0].target_id
        self.single_sheet_name = self.sheet_name_links[0].text
        self.single_sheet_name_id = self.sheet_name_links[0].target_id
        self.sheet_sort = self.sheet_numbers if self.sheet_numbers != "N/A" else "~~~~"
        self.view_name = text(values.get("view_name"), "Unknown View")
        self.view_type = text(values.get("view_type"))
        self.comments = text(values.get("comments"))
        self.placement = text(values.get("placement"), "Unknown")
        # A cloud that lives on a sheet already says so in the Sheet column, so the
        # review grid shows a dash rather than repeating the host view, and the cell
        # is not a navigation link (target 0 renders as plain text).
        self.on_sheet = self.placement == "On Sheet"
        self.view_display = "-" if self.on_sheet else self.view_name
        self.view_display_id = 0 if self.on_sheet else self.owner_view_id
        self.worksharing_owner = text(values.get("worksharing_owner"))
        self.created_by = text(values.get("created_by"))
        self.edited_by = text(values.get("edited_by"))
        self.current_user = text(values.get("current_user"))
        self.owned_by_other = bool(values.get("owned_by_other"))
        self.missing_comment = bool(values.get("missing_comment"))
        self.cloud_in_view = bool(values.get("cloud_in_view"))
        self.not_on_sheet = bool(values.get("not_on_sheet"))
        self.revision_parameters = dict(values.get("revision_parameters") or {})
        self.cloud_parameters = dict(values.get("cloud_parameters") or {})
        issues = []
        if self.missing_comment:
            issues.append("Missing Comment")
        if self.not_on_sheet:
            issues.append("Not on Sheet")
        elif self.cloud_in_view:
            issues.append("Review Placement")
        self.issues = issues
        self.issues_text = "; ".join(issues)
        if self.missing_comment or self.not_on_sheet:
            self.issue_severity = "Error"
        elif self.cloud_in_view:
            self.issue_severity = "Warning"
        else:
            self.issue_severity = "None"
        self.placement_badges = [BadgeItem(self.placement, "Placement")]
        self.issue_badges = [
            BadgeItem(issue, "Error" if issue in ("Missing Comment", "Not on Sheet") else "Warning")
            for issue in issues
        ]
        search_values = [
            self.cloud_id,
            self.revision_number,
            self.revision_display,
            self.revision_date,
            self.revision_description,
            self.sheet_numbers,
            self.sheet_names,
            self.view_name,
            self.view_type,
            self.comments,
            self.placement,
            self.issues_text,
            self.created_by,
            self.edited_by,
            self.worksharing_owner,
        ]
        self.search_text = normalized(" | ".join([text(x) for x in search_values]))

    @property
    def owner_display(self):
        if self.sheet_numbers != "N/A" and self.placement == "On Sheet":
            return "{} - {}".format(self.sheet_numbers, self.sheet_names)
        return self.view_name

    @property
    def worksharing_owner_display(self):
        return self.worksharing_owner or "Available"

    @property
    def created_by_display(self):
        return self.created_by or "Unknown"

    @property
    def edited_by_display(self):
        return self.edited_by or "Unknown"

    @property
    def has_issue(self):
        return bool(self.issues)

    def update_comment(self, value):
        self.comments = text(value)
        self.missing_comment = not bool(self.comments.strip())
        issues = []
        if self.missing_comment:
            issues.append("Missing Comment")
        if self.not_on_sheet:
            issues.append("Not on Sheet")
        elif self.cloud_in_view:
            issues.append("Review Placement")
        self.issues = issues
        self.issues_text = "; ".join(issues)
        if self.missing_comment or self.not_on_sheet:
            self.issue_severity = "Error"
        elif self.cloud_in_view:
            self.issue_severity = "Warning"
        else:
            self.issue_severity = "None"
        self.issue_badges = [
            BadgeItem(issue, "Error" if issue in ("Missing Comment", "Not on Sheet") else "Warning")
            for issue in issues
        ]
        self.search_text = normalized(" | ".join([
            text(self.cloud_id), self.revision_number, self.revision_display, self.revision_date,
            self.revision_description, self.sheet_numbers, self.sheet_names,
            self.view_name, self.view_type, self.comments, self.placement,
            self.issues_text, self.created_by, self.edited_by, self.worksharing_owner,
        ]))


class RevisionOption(object):
    def __init__(self, revision_id, number, description, date_value="", sequence=0, selected=True):
        self.revision_id = int(revision_id or 0)
        self.number = text(number, "N/A")
        self.description = text(description)
        self.date = text(date_value)
        self.sequence = int(sequence or 0)
        self.is_selected = bool(selected)
        suffix = " - {}".format(self.description) if self.description else ""
        self.display_text = "{}{}".format(self.number, suffix)


class ReportColumnOption(object):
    def __init__(self, key, label, source, parameter_name="", selected=False,
                 width_weight=1.35, bold=False):
        self.key = text(key)
        self.label = text(label)
        self.source = text(source)
        self.parameter_name = text(parameter_name)
        self.is_selected = bool(selected)
        self.is_bold = bool(bold)
        try:
            self.width_weight = max(0.5, min(4.0, float(width_weight)))
        except Exception:
            self.width_weight = 1.35


BASE_REPORT_COLUMN_DEFINITIONS = (
    ("sheet_number", "Sheet Number", "Cloud / Sheet", True),
    ("sheet_name", "Sheet Name", "Cloud / Sheet", True),
    ("comments", "Comments", "Revision Cloud", True),
    ("placement", "Placement", "Audit", True),
    ("issues", "Issues", "Audit", True),
)


WORKSHARING_REPORT_COLUMN_DEFINITIONS = (
    ("created_by", "Created By", "Worksharing", False),
    ("edited_by", "Edited By", "Worksharing", False),
    ("owned_by", "Owned By", "Worksharing", False),
)


def default_report_column_width(key):
    key = text(key)
    if key == "comments":
        return 2.4
    if key in ("placement", "issues"):
        return 1.35
    if key == "sheet_number":
        return 1.0
    return 1.35


def build_report_column_options(revision_parameter_names=None, cloud_parameter_names=None, saved_schema=None):
    """Return ordered report columns, merging current parameters with saved preferences."""
    discovered = []
    for key, label, source, selected in BASE_REPORT_COLUMN_DEFINITIONS:
        discovered.append(ReportColumnOption(
            key, label, source, selected=selected,
            width_weight=default_report_column_width(key)))
    for key, label, source, selected in WORKSHARING_REPORT_COLUMN_DEFINITIONS:
        discovered.append(ReportColumnOption(
            key, label, source, selected=selected,
            width_weight=default_report_column_width(key)))
    excluded_cloud_names = set(["comments", "design option", "type name", "workset"])
    for name in sorted(set([text(item) for item in list(cloud_parameter_names or []) if text(item)])):
        if normalized(name) in excluded_cloud_names:
            continue
        discovered.append(ReportColumnOption(
            "cloud::{}".format(name), name, "Revision Cloud", parameter_name=name, selected=False,
            width_weight=default_report_column_width("cloud::{}".format(name))))

    by_key = dict([(item.key, item) for item in discovered])
    ordered = []
    for record in list(saved_schema or []):
        try:
            key = text(record.get("key"))
            item = by_key.pop(key, None)
            if item is None:
                continue
            item.is_selected = bool(record.get("selected"))
            item.is_bold = bool(record.get("bold", False))
            try:
                item.width_weight = max(0.5, min(4.0, float(record.get("width", item.width_weight))))
            except Exception:
                pass
            ordered.append(item)
        except Exception:
            continue
    ordered.extend(discovered_item for discovered_item in discovered if discovered_item.key in by_key)
    return ordered


def selected_report_columns(options):
    return [item for item in list(options or []) if item.is_selected]


def report_column_value(row, column):
    key = text(getattr(column, "key", column))
    if key.startswith("revision::"):
        return text((row.get("revision_parameters") or {}).get(key.split("::", 1)[1]))
    if key.startswith("cloud::"):
        return text((row.get("cloud_parameters") or {}).get(key.split("::", 1)[1]))
    return text(row.get(key))


def filter_clouds(rows, search="", revision="All revisions", issue="All issue status",
                  placement="All placement", sheet="All sheets"):
    query = normalized(search)
    revision_key = normalized(revision)
    issue_key = normalized(issue)
    placement_key = normalized(placement)
    sheet_key = normalized(sheet)
    filtered = []
    for row in list(rows or []):
        if query and query not in row.search_text:
            continue
        if revision_key and revision_key != "all revisions" and normalized(row.revision_number) != revision_key:
            continue
        if issue_key == "issues only" and not row.has_issue:
            continue
        if issue_key == "no issues" and row.has_issue:
            continue
        if issue_key == "missing comment" and not row.missing_comment:
            continue
        if issue_key == "review placement" and not row.cloud_in_view:
            continue
        if issue_key == "not on sheet" and not row.not_on_sheet:
            continue
        if placement_key and placement_key != "all placement" and normalized(row.placement) != placement_key:
            continue
        if sheet_key and sheet_key != "all sheets":
            if sheet_key not in [normalized(x) for x in row.sheet_numbers_list]:
                continue
        filtered.append(row)
    return filtered


def selected_revision_ids(options):
    return set([item.revision_id for item in list(options or []) if item.is_selected])


def build_report_rows(rows, revision_ids, deduplicate=True):
    """Expand canonical clouds into sheet/report rows."""
    selected = set([int(x) for x in list(revision_ids or [])])
    report_rows = []
    seen = set()
    for cloud in sorted(list(rows or []), key=lambda x: (x.revision_sort, x.sheet_sort, x.cloud_id)):
        if cloud.revision_id not in selected:
            continue
        pairs = list(zip(cloud.sheet_numbers_list, cloud.sheet_names_list))
        if not pairs:
            pairs = [("N/A", "N/A")]
        for sheet_number, sheet_name in pairs:
            key = (cloud.revision_id, normalized(sheet_number), normalized(cloud.comments))
            if deduplicate and cloud.comments.strip() and key in seen:
                continue
            if deduplicate and cloud.comments.strip():
                seen.add(key)
            report_rows.append({
                "revision_id": cloud.revision_id,
                "revision_number": cloud.revision_number,
                "revision_display": cloud.revision_display,
                "revision_date": cloud.revision_date,
                "revision_description": cloud.revision_description,
                "sequence": cloud.revision_sort,
                "cloud_id": cloud.cloud_id,
                "sheet_number": text(sheet_number, "N/A") or "N/A",
                "sheet_name": text(sheet_name, "N/A") or "N/A",
                "view_name": cloud.view_name,
                "comments": cloud.comments,
                "placement": cloud.placement,
                "issues": cloud.issues_text,
                "created_by": cloud.created_by,
                "edited_by": cloud.edited_by,
                "owned_by": cloud.worksharing_owner_display,
                "revision_parameters": dict(cloud.revision_parameters),
                "cloud_parameters": dict(cloud.cloud_parameters),
            })
    return report_rows


def group_report_rows(report_rows):
    groups = []
    by_id = {}
    for row in list(report_rows or []):
        revision_id = int(row.get("revision_id") or 0)
        group = by_id.get(revision_id)
        if group is None:
            group = {
                "revision_id": revision_id,
                "number": text(row.get("revision_number"), "N/A"),
                "date": text(row.get("revision_date")),
                "description": text(row.get("revision_description")),
                "sequence": int(row.get("sequence") or 0),
                "rows": [],
            }
            by_id[revision_id] = group
            groups.append(group)
        group["rows"].append(row)
    return sorted(groups, key=lambda x: x["sequence"])


def plain_text_report(metadata, groups, columns):
    lines = ["COOLSYS ENERGY DESIGN", "PROJECT REVISION SUMMARY", ""]
    for label, key in (
        ("Project Number", "project_number"),
        ("Client", "client"),
        ("Project Name", "project_name"),
        ("Report Date", "report_date"),
    ):
        lines.append("{}: {}".format(label, text(metadata.get(key))))
    lines.append("")
    for group in groups:
        header = "Revision {} | Date: {} | Description: {}".format(
            group["number"], group["date"], group["description"])
        lines.extend([header, "-" * len(header)])
        lines.append("\t".join([item.label for item in columns]))
        if not group["rows"]:
            lines.append("No revision clouds found for this revision.")
        for row in group["rows"]:
            values = [report_column_value(row, column) for column in columns]
            lines.append("\t".join([text(x) for x in values]))
        lines.append("")
    return "\r\n".join(lines)
