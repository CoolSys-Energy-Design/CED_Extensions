# -*- coding: utf-8 -*-
"""Offline documentation browser services shared by CED pyRevit tools."""

from Documentation.catalog import Catalog, CatalogError
from Documentation.history import NavigationHistory
from Documentation.pathing import DocumentationPathError, resolve_documentation_root

__all__ = (
    "Catalog",
    "CatalogError",
    "DocumentationPathError",
    "NavigationHistory",
    "resolve_documentation_root",
)

