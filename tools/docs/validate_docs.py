# -*- coding: utf-8 -*-
"""Named entry point for documentation source validation."""

from __future__ import print_function

import sys

from documentation_build import run_validate

if __name__ == "__main__":
    sys.exit(run_validate())

