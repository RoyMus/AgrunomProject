# PyInstaller runtime hook: runs before main.py inside the built app.
# A windowed build (console=False) has no console, so sys.stdout / sys.stderr are None.
# Some libraries (numpy.f2py, imported by scipy/statsmodels) call sys.stderr.write at import time
# and crash with "'NoneType' object has no attribute 'write'". Point them at a null device instead.
import os
import sys

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")
