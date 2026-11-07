# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

PROJECT_ROOT = Path(SPECPATH).resolve()
MAIN = PROJECT_ROOT / "DailyWorkReports_15.2.py"
ICON = PROJECT_ROOT / "daily_work_reports_icon.ico"
VERSION_FILE = PROJECT_ROOT / "version_info_public.txt"

a = Analysis(
    [str(MAIN)],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=[
        (str(ICON), "."),
        *collect_data_files("reportlab"),
        *collect_data_files("pptx"),
    ],
    hiddenimports=[
        "daily_reports_ai", "structure_studio", "modern_ui_v153", "public_setup_v155", "tkinter", "tkinter.filedialog",
        "tkinter.messagebox", "tkinter.simpledialog", "tkinter.ttk",
        "openpyxl", "openpyxl.chart", "openpyxl.styles",
        "PIL.Image", "PIL.ImageDraw", "PIL.ImageFont",
        "reportlab", "reportlab.platypus", *collect_submodules("pptx"),
    ],
    hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="DailyWorkReports_Public_v15.5.0", debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False, icon=str(ICON),
    version=str(VERSION_FILE), uac_admin=False,
)
