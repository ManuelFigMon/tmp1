# One-line workspace bootstrap: %run cgs_ai_setup  — adds workspace root to sys.path and imports cgs_ai.
import sys as _sys, importlib as _importlib
# Ensure the workspace root is on sys.path so cgs_ai can be found
if '' not in _sys.path:
    _sys.path.insert(0, '')
import cgs_ai
_importlib.reload(cgs_ai)  # picks up any edits made to __init__.py mid-session
