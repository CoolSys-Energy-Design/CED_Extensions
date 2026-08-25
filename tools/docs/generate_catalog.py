# -*- coding: utf-8 -*-
"""Named entry point for generating docs/user-guide/catalog.json."""

from __future__ import print_function

import sys

from documentation_build import run_generate

if __name__ == "__main__":
    sys.exit(run_generate())

