"""Import the versioned Python entry-point using its exact path for QA."""
from __future__ import annotations
import importlib.util
from pathlib import Path
import sys

_module = None

def main_module():
    global _module
    if _module is None:
        path = Path(__file__).parent / "DailyWorkReports_15.2.py"
        spec = importlib.util.spec_from_file_location("dwr_version_15_2", path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        _module = module
    return _module
