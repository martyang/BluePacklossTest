#!/usr/local/bin/python
# -*- coding: utf-8 -*-
# ============================================================
# TEST1.PY
#
# Note:
# ============================================================

import sys
from PyQt5.QtWidgets import QApplication
from mainWindows import UiWindows

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = UiWindows()
    sys.exit(app.exec_())
