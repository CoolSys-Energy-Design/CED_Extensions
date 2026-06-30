# -*- coding: utf-8 -*-
__title__ = "Category\nConverter"
__doc__ = (
    "Batch-convert the family category of one or more .rfa files. "
    "Pick the families, choose a single target category, and each is "
    "opened silently, reassigned, saved, and closed."
)

import os

from pyrevit import DB, forms, script
from pyrevit import HOST_APP

logger = script.get_logger()
output = script.get_output()


def _family_categories(fam_doc):
    """name -> Category for the family-assignable categories of a document.

    Returns the top-level *model* categories (the ones that show up in
    Revit's "Family Category and Parameters" dialog). Tag/annotation and
    sub-categories are excluded. Not every entry is guaranteed assignable
    to every family, so the actual set is still wrapped in try/except.
    """
    result = {}
    for cat in fam_doc.Settings.Categories:
        try:
            if cat.CategoryType != DB.CategoryType.Model:
                continue
            if cat.Parent is not None:
                continue
        except Exception:
            continue
        result[cat.Name] = cat
    return result


def _current_category_name(fam_doc):
    try:
        cat = fam_doc.OwnerFamily.FamilyCategory
        return cat.Name if cat else "<none>"
    except Exception:
        return "<unknown>"


def _convert_category(fam_doc, target_name):
    """Set ``fam_doc``'s family category to ``target_name``.

    Returns (ok, message). Re-resolves the Category object from the
    target document (Category objects are document-specific, so the
    selection made on the first family can't be reused on the rest).
    """
    cats = _family_categories(fam_doc)
    new_cat = cats.get(target_name)
    if new_cat is None:
        return False, "target category '{}' not available in this family".format(target_name)

    current = fam_doc.OwnerFamily.FamilyCategory
    if current is not None and current.Name == target_name:
        return True, "already {}".format(target_name)

    t = DB.Transaction(fam_doc, "Convert Family Category")
    t.Start()
    try:
        fam_doc.OwnerFamily.FamilyCategory = new_cat
        t.Commit()
        return True, "converted to {}".format(target_name)
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return False, "Revit refused the conversion: {}".format(ex)


def main():
    paths = forms.pick_file(
        file_ext="rfa",
        multi_file=True,
        title="Select RFA families to convert",
    )
    if not paths:
        script.exit()
    if isinstance(paths, str):
        paths = [paths]

    app = HOST_APP.app
    output.close_others()
    output.print_md("# Category Converter")

    target_name = None
    results = []  # (file_name, status_message)

    for path in paths:
        name = os.path.basename(path)
        fam_doc = None
        try:
            try:
                fam_doc = app.OpenDocumentFile(path)
            except Exception as ex:
                results.append((name, "OPEN FAILED: {}".format(ex)))
                continue

            if not fam_doc.IsFamilyDocument:
                results.append((name, "SKIPPED: not a family (.rfa) document"))
                fam_doc.Close(False)
                continue

            current_name = _current_category_name(fam_doc)

            # Prompt for the target category once, off the first family's
            # category list, then apply that choice to every file.
            if target_name is None:
                options = sorted(_family_categories(fam_doc).keys())
                target_name = forms.SelectFromList.show(
                    options,
                    title="Convert family category to...  (first family is '{}')".format(
                        current_name
                    ),
                    multiselect=False,
                    button_name="Convert",
                )
                if not target_name:
                    fam_doc.Close(False)
                    forms.alert("No target category selected. Nothing was changed.")
                    script.exit()

            ok, message = _convert_category(fam_doc, target_name)
            if ok:
                try:
                    fam_doc.Save()
                    results.append((name, "{} -> {}".format(current_name, message)))
                except Exception as ex:
                    results.append((name, "CONVERTED BUT SAVE FAILED: {}".format(ex)))
            else:
                results.append((name, "FAILED: {}".format(message)))

            fam_doc.Close(False)
        except Exception as ex:
            logger.warning("Unhandled error on {}: {}".format(name, ex))
            results.append((name, "ERROR: {}".format(ex)))
            try:
                if fam_doc is not None:
                    fam_doc.Close(False)
            except Exception:
                pass

    output.print_md("**Target category:** {}".format(target_name or "<none>"))
    output.print_md("**{} file(s) processed**".format(len(results)))
    for fname, msg in results:
        output.print_md("- **{}** — {}".format(fname, msg))


if __name__ == "__main__":
    main()
