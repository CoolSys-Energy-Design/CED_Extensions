# -*- coding: utf-8 -*-
from __future__ import print_function

import errno
import os

SOURCE_ROOT = r"C:\ACC\ACCDocs\CoolSys"

DESTINATION_ROOTS = [
    r"C:\Users\Aevelina\ACC\ACCDocs",
    r"C:\Users\Aevelina\DC\ACCDocs",
    r"C:\DC\ACCDocs"
]

FULL_TREE_PROJECTS = [
    "CED Content Collection"
]

DRY_RUN = False
SKIP_ERRORS = True


def make_folder(folder_path):
    if DRY_RUN:
        print("Would create: {}".format(folder_path))
        return True

    try:
        os.makedirs(folder_path)
        print("Created: {}".format(folder_path))
        return True

    except OSError as err:
        if err.errno == errno.EEXIST and os.path.isdir(folder_path):
            print("Already exists: {}".format(folder_path))
            return True

        print("FAILED: {}".format(folder_path))
        print("Reason: {}".format(err))

        if SKIP_ERRORS:
            return False

        raise


def get_destination_coolsys_root(destination_root, source_root):
    """
    Destination roots are expected to be ACCDocs-level folders.

    Example:
    C:\\Users\\Aevelina\\ACC\\ACCDocs

    This creates/uses:
    C:\\Users\\Aevelina\\ACC\\ACCDocs\\CoolSys

    If the destination root already ends with CoolSys, it uses it directly.
    """
    source_folder_name = os.path.basename(os.path.normpath(source_root))
    destination_root = os.path.normpath(destination_root)

    if os.path.basename(destination_root).lower() == source_folder_name.lower():
        return destination_root

    return os.path.join(destination_root, source_folder_name)


def is_full_tree_project(folder_name):
    for project_name in FULL_TREE_PROJECTS:
        if folder_name.lower() == project_name.lower():
            return True
    return False


def replicate_project_parent(source_project_path, destination_project_path):
    """
    Create only the top-level project folder.
    Does not create any child folders.
    """
    make_folder(destination_project_path)


def replicate_full_tree(source_project_path, destination_project_path):
    """
    Recreate every folder under the selected project.
    Files are ignored.
    """
    for current_folder, subfolders, filenames in os.walk(source_project_path):
        relative_path = os.path.relpath(current_folder, source_project_path)

        if relative_path == ".":
            target_folder = destination_project_path
        else:
            target_folder = os.path.join(destination_project_path, relative_path)

        make_folder(target_folder)


def replicate_coolsys_structure(source_root, destination_roots):
    source_root = os.path.normpath(source_root.strip().strip('"'))

    if not os.path.isdir(source_root):
        raise ValueError("Source root does not exist or is not a folder: {}".format(source_root))

    project_folders = []

    for item_name in os.listdir(source_root):
        source_item_path = os.path.join(source_root, item_name)

        if os.path.isdir(source_item_path):
            project_folders.append(item_name)

    project_folders.sort()

    print("Project folders found: {}".format(len(project_folders)))
    print("")

    for destination_root in destination_roots:
        destination_coolsys_root = get_destination_coolsys_root(destination_root, source_root)

        print("")
        print("Destination CoolSys root:")
        print(destination_coolsys_root)

        make_folder(destination_coolsys_root)

        for project_folder in project_folders:
            source_project_path = os.path.join(source_root, project_folder)
            destination_project_path = os.path.join(destination_coolsys_root, project_folder)

            if is_full_tree_project(project_folder):
                print("")
                print("Replicating FULL tree: {}".format(project_folder))
                replicate_full_tree(source_project_path, destination_project_path)
            else:
                print("")
                print("Creating parent folder only: {}".format(project_folder))
                replicate_project_parent(source_project_path, destination_project_path)

    print("")
    print("Finished.")


def main():
    replicate_coolsys_structure(SOURCE_ROOT, DESTINATION_ROOTS)


if __name__ == "__main__":
    main()