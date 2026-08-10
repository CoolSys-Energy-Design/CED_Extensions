# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)


if __name__ == "__main__":
    suite = unittest.defaultTestLoader.discover(HERE, pattern="test_*.py")
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)
