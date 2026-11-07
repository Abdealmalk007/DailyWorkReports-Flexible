from __future__ import annotations

import calendar
import copy
import csv
import ctypes
import getpass
import hashlib
import json
import os
import re
import secrets
import shutil
import sqlite3
import sys
import uuid
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from daily_reports_ai import (
    neural_forecast,
    productivity_insights,
    report_quality_score,
    robust_anomalies,
    scenario_forecast,
)
from structure_studio import StructureStudio, active_category_conflicts
from modern_ui_v153 import ModernHome


_TK_DLL_DIRECTORY_HANDLES: list[object] = []
_TK_LOADED_TCL_DLLS: list[object] = []


def configure_tk_runtime() -> None:
    """Prepare Tcl/Tk for portable Python environments before Tk is imported."""
    candidates: list[tuple[Path, Path, Path | None]] = []
    env_tcl = os.environ.get("TCL_LIBRARY", "").strip()
    env_tk = os.environ.get("TK_LIBRARY", "").strip()
    if env_tcl and env_tk:
        candidates.append((Path(env_tcl), Path(env_tk), Path(env_tcl).parent))
    if getattr(sys, "frozen", False):
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        executable_root = Path(sys.executable).resolve().parent
        candidates.extend(
            [
                (bundle_root / "_tcl_data", bundle_root / "_tk_data", bundle_root),
                (executable_root / "_tcl_data", executable_root / "_tk_data", executable_root),
            ]
        )
    for base in (Path(sys.base_prefix), Path(sys.prefix), Path(sys.exec_prefix)):
        candidates.append((base / "tcl" / "tcl8.6", base / "tcl" / "tk8.6", base / "DLLs"))

    seen: set[tuple[Path, Path, Path | None]] = set()
    for tcl_library, tk_library, dll_dir in candidates:
        key = (tcl_library, tk_library, dll_dir)
        if key in seen:
            continue
        seen.add(key)
        if not (tcl_library / "init.tcl").is_file() or not (tk_library / "tk.tcl").is_file():
            continue
        os.environ["TCL_LIBRARY"] = tcl_library.as_posix()
        os.environ["TK_LIBRARY"] = tk_library.as_posix()
        if dll_dir is not None and dll_dir.is_dir():
            dll_text = str(dll_dir)
            path_parts = os.environ.get("PATH", "").split(os.pathsep)
            if dll_text not in path_parts:
                os.environ["PATH"] = dll_text + os.pathsep + os.environ.get("PATH", "")
            if os.name == "nt" and hasattr(os, "add_dll_directory"):
                try:
                    _TK_DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(dll_text))
                except OSError:
                    pass
        if os.name == "nt":
            _warm_windows_tcl(tcl_library, dll_dir)
        return


def _warm_windows_tcl(tcl_library: Path, dll_dir: Path | None) -> None:
    dll_candidates: list[Path] = []
    if dll_dir is not None:
        dll_candidates.append(dll_dir / "tcl86t.dll")
    for base in (Path(sys.base_prefix), Path(sys.prefix)):
        dll_candidates.append(base / "DLLs" / "tcl86t.dll")
    tcl_dll_path = next((path for path in dll_candidates if path.is_file()), None)
    if tcl_dll_path is None:
        return
    try:
        tcl_dll = ctypes.WinDLL(str(tcl_dll_path))
    except OSError:
        return
    _TK_LOADED_TCL_DLLS.append(tcl_dll)
    tcl_dll.Tcl_FindExecutable.argtypes = [ctypes.c_char_p]
    tcl_dll.Tcl_FindExecutable.restype = None
    tcl_dll.Tcl_CreateInterp.argtypes = []
    tcl_dll.Tcl_CreateInterp.restype = ctypes.c_void_p
    tcl_dll.Tcl_Eval.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    tcl_dll.Tcl_Eval.restype = ctypes.c_int
    tcl_dll.Tcl_DeleteInterp.argtypes = [ctypes.c_void_p]
    tcl_dll.Tcl_DeleteInterp.restype = None
    tcl_dll.Tcl_FindExecutable(str(Path(sys.executable)).encode("utf-8"))
    interp = tcl_dll.Tcl_CreateInterp()
    if not interp:
        return
    try:
        init_tcl = (tcl_library / "init.tcl").as_posix()
        for command in ("pwd", "file exists {C:/Windows}", f"file exists {{{init_tcl}}}"):
            tcl_dll.Tcl_Eval(interp, command.encode("utf-8"))
    finally:
        tcl_dll.Tcl_DeleteInterp(interp)


configure_tk_runtime()

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from PIL import Image, ImageDraw, ImageFont

    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.platypus import (
        KeepTogether,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table as ExcelTable
    from openpyxl.worksheet.table import TableStyleInfo

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from pptx import Presentation

    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


ARABIC_APP_NAME = "تقارير العمل اليومية"
APP_NAME = f"{ARABIC_APP_NAME} | Daily Work Report"
APP_VERSION = "15.5.0"
CONFIG_VERSION = 11

REPORT_STATUSES = {
    "draft": "مسودة",
    "completed": "مكتمل",
    "approved": "معتمد",
}

FIELD_KIND_LABELS = {
    "number": "رقمي",
    "text": "نصي",
    "choice": "قائمة منسدلة",
    "date": "تاريخ",
    "time": "وقت",
    "boolean": "نعم / لا",
    "percentage": "نسبة مئوية",
    "currency": "مبلغ مالي",
    "phone": "هاتف / جوال",
}
FIELD_KIND_BY_LABEL = {label: kind for kind, label in FIELD_KIND_LABELS.items()}
SUPPORTED_FIELD_KINDS = set(FIELD_KIND_LABELS)

APPEARANCE_MODES = {
    "light": "فاتح",
    "dark": "داكن",
}
DENSITY_LABELS = {
    "compact": "مضغوط",
    "comfortable": "مريح",
    "spacious": "واسع",
}
ROLE_LABELS = {
    "admin": "مدير النظام",
    "editor": "محرر",
    "viewer": "مشاهد",
}
ROLE_PERMISSIONS = {
    "admin": {"*", "view", "add", "edit", "delete", "approve", "manage", "sync", "recover", "intelligence"},
    "editor": {"view", "add", "edit"},
    "viewer": {"view"},
}
REPORT_THEMES = {
    "executive_teal": {
        "label": "الرسمي الفاخر - أخضر بترولي وذهبي",
        "primary": "#214E52",
        "primary_2": "#32696A",
        "soft": "#E9F2F0",
        "accent": "#C39A5B",
        "accent_dark": "#9C7440",
        "accent_soft": "#F6EFE4",
        "alt_row": "#F3F7F6",
        "border": "#CFDDDA",
    },
    "royal_blue": {
        "label": "الإداري العصري - أزرق ملكي وفضي",
        "primary": "#25446D",
        "primary_2": "#3E6393",
        "soft": "#EAF0F8",
        "accent": "#9CA9B8",
        "accent_dark": "#657487",
        "accent_soft": "#F1F3F6",
        "alt_row": "#F4F7FB",
        "border": "#CDD6E2",
    },
    "burgundy": {
        "label": "الكلاسيكي الرسمي - عنابي ورملي",
        "primary": "#6B3341",
        "primary_2": "#8A4B5A",
        "soft": "#F5EDEF",
        "accent": "#C5A36B",
        "accent_dark": "#927243",
        "accent_soft": "#F7F0E4",
        "alt_row": "#FAF6F7",
        "border": "#DECED2",
    },
}

DEFAULT_CONFIG: dict[str, Any] = {
    "version": CONFIG_VERSION, "report_theme": "executive_teal",
    "report_layout": {"orientation": "landscape"},
    "app_state": {"company_id": "", "template_id": "", "dashboard_filters": {},
                  "history_filters": {}, "appearance": {"mode": "light", "font_scale": 1.0,
                  "density": "comfortable"}, "entry_table_widths": {}},
    "identity": {"report_title": "", "company_name": "", "user_name": "",
                 "default_code": "", "default_report_number": "1", "report_footer": ""},
    "category_order": [], "categories": {},
}

def make_qa_config() -> dict[str, Any]:
    """Synthetic fixture used only by tests; never normal startup."""
    config = copy.deepcopy(DEFAULT_CONFIG)
    config["identity"]["report_title"] = APP_NAME
    config["category_order"] = ["drawers", "technicians", "showroom_employees"]
    config["categories"] = {
        "drawers": {"label": "قسم اختباري 1", "person_label": "الشخص",
                    "total_label": "الإجمالي", "people": ["موظف اختبار"],
                    "fields": [
                        {"id": "colored", "label": "رقم أ", "kind": "number"},
                        {"id": "edits", "label": "رقم ب", "kind": "number"},
                        {"id": "black_white", "label": "رقم ج", "kind": "number"}]},
        "technicians": {"label": "قسم اختباري 2", "person_label": "الشخص",
                        "total_label": "الإجمالي", "people": ["موظف اختبار"],
                        "fields": [
                            {"id": name, "label": f"حقل {i}", "kind": "text" if name == "notes" else "number"}
                            for i, name in enumerate(["new_measure", "confirmed_measure", "outlets",
                                                       "cancelled", "new_requests", "notes"], 1)]},
        "showroom_employees": {"label": "قسم اختباري 3", "person_label": "اسم الموظف",
                               "total_label": "الإجمالي", "people": ["موظف اختبار"],
                               "fields": [
                                   {"id": "measurements_count", "label": "رقم د", "kind": "number"},
                                   {"id": "drawings_count", "label": "رقم هـ", "kind": "number"}]},
    }
    return config


ARABIC_WEEKDAYS = (
    "الإثنين",
    "الثلاثاء",
    "الأربعاء",
    "الخميس",
    "الجمعة",
    "السبت",
    "الأحد",
)
ARABIC_MONTHS = (
    "يناير",
    "فبراير",
    "مارس",
    "أبريل",
    "مايو",
    "يونيو",
    "يوليو",
    "أغسطس",
    "سبتمبر",
    "أكتوبر",
    "نوفمبر",
    "ديسمبر",
)


class Palette:
    PRIMARY = "#315E60"
    PRIMARY_DARK = "#24494B"
    PRIMARY_LIGHT = "#DDEBE8"
    ACCENT = "#B48A57"
    ACCENT_LIGHT = "#F1E7D8"
    SUCCESS = "#4F7668"
    DANGER = "#A65D57"
    INFO = "#547A8B"
    INK = "#263536"
    MUTED = "#708080"
    PALE = "#F5F7F6"
    PALE_BLUE = "#EAF2F0"
    ALT_ROW = "#F0F5F3"
    WHITE = "#FFFFFF"
    BORDER = "#CBD9D6"
    NAVY = PRIMARY_DARK
    BLUE = PRIMARY
    BLUE_2 = INFO
    GOLD = ACCENT_LIGHT
    ORANGE = ACCENT
    GREEN = SUCCESS
    RED = DANGER

    @classmethod
    def use_mode(cls, mode: str) -> None:
        if str(mode).lower() == "dark":
            cls.PRIMARY = "#4D8585"
            cls.PRIMARY_DARK = "#17292B"
            cls.PRIMARY_LIGHT = "#263B3C"
            cls.ACCENT = "#D1A96A"
            cls.ACCENT_LIGHT = "#4B3E2C"
            cls.SUCCESS = "#5F947F"
            cls.DANGER = "#C8756E"
            cls.INFO = "#6E9AAE"
            cls.INK = "#E7EFEF"
            cls.MUTED = "#AAB9B8"
            cls.PALE = "#182223"
            cls.PALE_BLUE = "#223334"
            cls.ALT_ROW = "#243031"
            cls.WHITE = "#202A2B"
            cls.BORDER = "#415253"
        else:
            cls.PRIMARY = "#315E60"
            cls.PRIMARY_DARK = "#24494B"
            cls.PRIMARY_LIGHT = "#DDEBE8"
            cls.ACCENT = "#B48A57"
            cls.ACCENT_LIGHT = "#F1E7D8"
            cls.SUCCESS = "#4F7668"
            cls.DANGER = "#A65D57"
            cls.INFO = "#547A8B"
            cls.INK = "#263536"
            cls.MUTED = "#708080"
            cls.PALE = "#F5F7F6"
            cls.PALE_BLUE = "#EAF2F0"
            cls.ALT_ROW = "#F0F5F3"
            cls.WHITE = "#FFFFFF"
            cls.BORDER = "#CBD9D6"
        cls.NAVY = cls.PRIMARY_DARK
        cls.BLUE = cls.PRIMARY
        cls.BLUE_2 = cls.INFO
        cls.GOLD = cls.ACCENT_LIGHT
        cls.ORANGE = cls.ACCENT
        cls.GREEN = cls.SUCCESS
        cls.RED = cls.DANGER


def bundled_resource_path(filename: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


def _is_writable_directory(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / f".write_test_{uuid.uuid4().hex}.tmp"
        probe.write_text("ok", encoding="ascii")
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def app_data_dir() -> Path:
    override = os.environ.get("DAILY_WORK_REPORTS_DATA_DIR")
    if override:
        override_path = Path(override).expanduser().resolve()
        if _is_writable_directory(override_path):
            return override_path

    app_dir = Path(__file__).resolve().parent
    portable_root = os.environ.get("DAILY_WORK_REPORTS_PORTABLE_ROOT", "").strip()
    candidates: list[Path] = []
    if portable_root:
        candidates.append(Path(portable_root).expanduser().resolve() / "DailyWorkReportsFlexible_v155_Data")
    if (app_dir / "portable_mode.flag").exists():
        candidates.append(app_dir / "DailyWorkReportsFlexible_v155_Data")

    base = os.environ.get("LOCALAPPDATA")
    if base:
        candidates.append(Path(base) / "DailyWorkReportsFlexible_v155")
    candidates.extend(
        [
            Path.home() / ".daily_work_reports_flexible_v155",
            app_dir / "DailyWorkReportsFlexible_v155_Data",
            Path.cwd() / "DailyWorkReportsFlexible_v155_Data",
        ]
    )
    for candidate in candidates:
        if _is_writable_directory(candidate):
            return candidate.resolve()
    return candidates[-1].resolve()


def atomic_write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def read_json(path: Path, default: Any) -> Any:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return copy.deepcopy(default)
    return copy.deepcopy(default)


def unique_clean_strings(values: Any) -> list[str]:
    if not isinstance(values, list):
        return []
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        name = str(value).strip()
        key = name.casefold()
        if name and key not in seen:
            seen.add(key)
            result.append(name)
    return result


BUILTIN_FIELD_ORDER = {
    "drawers": ("colored", "edits", "black_white"),
    "technicians": (
        "new_measure",
        "confirmed_measure",
        "outlets",
        "cancelled",
        "new_requests",
        "notes",
    ),
    "showroom_employees": (
        "measurements_count",
        "drawings_count",
    ),
}


def _config_version_number(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def canonical_builtin_fields(
    category_id: str,
    fields: list[dict[str, Any]],
    *,
    restore_missing: bool,
) -> list[dict[str, Any]]:
    """Return built-in fields in the official RTL order, then custom fields."""
    requested_order = BUILTIN_FIELD_ORDER.get(category_id)
    if not requested_order:
        return fields
    by_id = {str(field.get("id", "")): field for field in fields}
    defaults = {
        str(field.get("id", "")): field
        for field in DEFAULT_CONFIG["categories"][category_id]["fields"]
    }
    ordered: list[dict[str, Any]] = []
    for field_id in requested_order:
        field = by_id.get(field_id)
        if field is None and restore_missing and field_id in defaults:
            field = copy.deepcopy(defaults[field_id])
            field.setdefault("options", [])
        if field is not None:
            ordered.append(field)
    ordered.extend(
        field
        for field in fields
        if str(field.get("id", "")) not in requested_order
    )
    return ordered


def move_field_by_drop(
    fields: list[dict[str, Any]],
    source_id: str,
    target_id: str,
    *,
    place_after: bool,
) -> bool:
    """Move one field around another; the list order is the RTL report order."""
    if not source_id or not target_id or source_id == target_id:
        return False
    source_index = next(
        (index for index, field in enumerate(fields) if field.get("id") == source_id),
        None,
    )
    if source_index is None:
        return False
    moved = fields.pop(source_index)
    target_index = next(
        (index for index, field in enumerate(fields) if field.get("id") == target_id),
        None,
    )
    if target_index is None:
        fields.insert(source_index, moved)
        return False
    fields.insert(target_index + (1 if place_after else 0), moved)
    return True


def normalize_config(raw: Any) -> dict[str, Any]:
    source = raw if isinstance(raw, dict) else {}
    source_config_version = _config_version_number(source.get("version", 0))
    source_categories = source.get("categories", {})
    if not isinstance(source_categories, dict):
        source_categories = {}

    source_identity = source.get("identity", {})
    if not isinstance(source_identity, dict):
        source_identity = {}
    default_identity = DEFAULT_CONFIG["identity"]
    report_title = str(
        source_identity.get("report_title", default_identity["report_title"])
    ).strip()
    if report_title in {ARABIC_APP_NAME, "تقرير العمل اليومي"}:
        report_title = APP_NAME
    identity = {
        "report_title": report_title,
        "company_name": str(
            source_identity.get("company_name", default_identity["company_name"])
        ).strip(),
        "user_name": str(
            source_identity.get("user_name", default_identity["user_name"])
        ).strip(),
        "default_code": str(
            source_identity.get("default_code", default_identity["default_code"])
        ).strip(),
        "default_report_number": str(
            source_identity.get(
                "default_report_number",
                default_identity["default_report_number"],
            )
        ).strip()
        or "1",
        "report_footer": str(source_identity.get("report_footer", default_identity.get("report_footer", ""))).strip(),
    }

    requested_order = unique_clean_strings(source.get("category_order", []))
    if not requested_order:
        requested_order = [
            category_id
            for category_id in DEFAULT_CONFIG["category_order"]
            if category_id in source_categories
        ]
    for raw_category_id in source_categories:
        clean_id = str(raw_category_id).strip()
        if clean_id and clean_id not in requested_order:
            requested_order.append(clean_id)
    if not requested_order:
        requested_order = list(DEFAULT_CONFIG["category_order"])
        source_categories = copy.deepcopy(DEFAULT_CONFIG["categories"])

    normalized: dict[str, Any] = {
        "version": CONFIG_VERSION,
        "report_theme": (
            str(source.get("report_theme", "executive_teal"))
            if str(source.get("report_theme", "executive_teal")) in REPORT_THEMES
            else "executive_teal"
        ),
        "report_layout": {
            "orientation": (
                str((source.get("report_layout", {}) if isinstance(source.get("report_layout", {}), dict) else {}).get("orientation", "landscape"))
                if str((source.get("report_layout", {}) if isinstance(source.get("report_layout", {}), dict) else {}).get("orientation", "landscape")) in {"landscape", "portrait"}
                else "landscape"
            )
        },
        "app_state": {
            "company_id": str(
                (
                    source.get("app_state", {})
                    if isinstance(source.get("app_state", {}), dict)
                    else {}
                ).get("company_id", "")
            ),
            "template_id": str(
                (
                    source.get("app_state", {})
                    if isinstance(source.get("app_state", {}), dict)
                    else {}
                ).get("template_id", "")
            ),
            "dashboard_filters": copy.deepcopy(
                (
                    source.get("app_state", {})
                    if isinstance(source.get("app_state", {}), dict)
                    else {}
                ).get("dashboard_filters", {})
                if isinstance(
                    (
                        source.get("app_state", {})
                        if isinstance(source.get("app_state", {}), dict)
                        else {}
                    ).get("dashboard_filters", {}),
                    dict,
                )
                else {}
            ),
            "history_filters": copy.deepcopy(
                (
                    source.get("app_state", {})
                    if isinstance(source.get("app_state", {}), dict)
                    else {}
                ).get("history_filters", {})
                if isinstance(
                    (
                        source.get("app_state", {})
                        if isinstance(source.get("app_state", {}), dict)
                        else {}
                    ).get("history_filters", {}),
                    dict,
                )
                else {}
            ),
            "appearance": {},
            "entry_table_widths": {},
        },
        "identity": identity,
        "category_order": [],
        "categories": {},
    }
    source_state = source.get("app_state", {}) if isinstance(source.get("app_state", {}), dict) else {}
    raw_appearance = source_state.get("appearance", {}) if isinstance(source_state.get("appearance", {}), dict) else {}
    mode = str(raw_appearance.get("mode", "light")).strip().lower()
    if mode not in APPEARANCE_MODES:
        mode = "light"
    density = str(raw_appearance.get("density", "comfortable")).strip().lower()
    if density not in DENSITY_LABELS:
        density = "comfortable"
    try:
        font_scale = float(raw_appearance.get("font_scale", 1.0))
    except (TypeError, ValueError):
        font_scale = 1.0
    font_scale = max(0.8, min(1.4, font_scale))
    normalized["app_state"]["appearance"] = {
        "mode": mode, "font_scale": round(font_scale, 2), "density": density,
    }
    widths_source = source_state.get("entry_table_widths", {})
    if isinstance(widths_source, dict):
        normalized["app_state"]["entry_table_widths"] = {
            str(key): max(45, min(1200, int(value)))
            for key, value in widths_source.items()
            if str(value).isdigit()
        }
    used_labels: set[str] = set()
    for category_id in requested_order:
        default_category = DEFAULT_CONFIG["categories"].get(
            category_id,
            {
                "label": "قسم جديد",
                "person_label": "الاسم",
                "total_label": "الإجمالي",
                "people": [],
                "fields": [],
            },
        )
        supplied = source_categories.get(category_id, {})
        if not isinstance(supplied, dict):
            supplied = copy.deepcopy(default_category)

        people_source = (
            supplied["people"] if "people" in supplied else default_category["people"]
        )
        fields_source = (
            supplied["fields"] if "fields" in supplied else default_category["fields"]
        )
        people = unique_clean_strings(people_source)

        fields: list[dict[str, Any]] = []
        field_ids: set[str] = set()
        field_labels: set[str] = set()
        if isinstance(fields_source, list):
            for supplied_field in fields_source:
                if not isinstance(supplied_field, dict):
                    continue
                label = str(supplied_field.get("label", "")).strip()
                if not label or label.casefold() in field_labels:
                    continue
                field_id = str(supplied_field.get("id", "")).strip()
                if not field_id or field_id in field_ids:
                    field_id = f"field_{uuid.uuid4().hex[:12]}"
                kind = str(supplied_field.get("kind", "number")).strip().lower()
                if kind not in SUPPORTED_FIELD_KINDS:
                    kind = "number"
                options = unique_clean_strings(supplied_field.get("options", []))
                if kind != "choice":
                    options = []
                field_ids.add(field_id)
                field_labels.add(label.casefold())
                fields.append(
                    {
                        "id": field_id,
                        "label": label,
                        "kind": kind,
                        "options": options,
                        "report_visible": bool(supplied_field.get("report_visible", True)),
                        "report_label": str(supplied_field.get("report_label", "")).strip(),
                        "required": bool(supplied_field.get("required", False)),
                    }
                )
        if source_config_version < 7:
            fields = canonical_builtin_fields(
                category_id,
                fields,
                restore_missing=True,
            )
        elif source_config_version < 11 and category_id in {"technicians", "showroom_employees"}:
            # v15.2 migration: add the new technician counter and apply the
            # requested official field order without touching custom fields.
            fields = canonical_builtin_fields(
                category_id,
                fields,
                restore_missing=True,
            )

        label = str(supplied.get("label", default_category["label"])).strip()
        if not label:
            label = "قسم جديد"
        base_label = label
        suffix = 2
        while label.casefold() in used_labels:
            label = f"{base_label} {suffix}"
            suffix += 1
        used_labels.add(label.casefold())
        normalized["category_order"].append(category_id)
        person_label = str(
            supplied.get(
                "person_label",
                default_category.get("person_label", "الاسم"),
            )
        ).strip() or "الاسم"
        if (
            source_config_version < 11
            and category_id == "showroom_employees"
            and person_label in {"الموظف", "الاسم"}
        ):
            person_label = "اسم الموظف"
        normalized["categories"][category_id] = {
            "label": label,
            "person_label": person_label,
            "total_label": str(
                supplied.get(
                    "total_label",
                    default_category.get("total_label", "الإجمالي"),
                )
            ).strip()
            or "الإجمالي",
            "people": people,
            "fields": fields,
        }
    return normalized


def category_ids(config: dict[str, Any]) -> list[str]:
    categories = config.get("categories", {})
    if not isinstance(categories, dict):
        return []
    ordered = [
        category_id
        for category_id in config.get("category_order", [])
        if category_id in categories
    ]
    ordered.extend(category_id for category_id in categories if category_id not in ordered)
    return ordered


def default_user_name() -> str:
    try:
        return getpass.getuser().strip()
    except (OSError, KeyError):
        return ""


def report_fingerprint(snapshot: dict[str, Any], company_id: str = "") -> str:
    normalized_rows: list[dict[str, Any]] = []
    for row in snapshot.get("rows", []):
        if not isinstance(row, dict):
            continue
        normalized_rows.append(
            {
                "category": str(row.get("category", "")).strip().casefold(),
                "name": str(row.get("name", "")).strip().casefold(),
                "fields": [
                    {
                        "id": str(field.get("id", "")).strip().casefold(),
                        "value": str(field.get("value", "")).strip().casefold(),
                    }
                    for field in row.get("fields", [])
                    if isinstance(field, dict)
                ],
                "total": coerce_nonnegative_int(row.get("total", 0)),
            }
        )
    identity = {
        "company": str(company_id or snapshot.get("company_name", "")).strip().casefold(),
        "date": str(snapshot.get("date", "")).strip(),
        "code": str(snapshot.get("code", "")).strip().casefold(),
        "number": str(snapshot.get("page", "")).strip().casefold(),
        "rows": normalized_rows,
    }
    payload = json.dumps(identity, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def report_content_hash(snapshot: dict[str, Any]) -> str:
    """Return a stable integrity hash without self-referential metadata."""
    payload = copy.deepcopy(snapshot)
    for key in ("verification_code", "revision_number", "saved_at", "updated_at"):
        payload.pop(key, None)
    canonical = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def password_digest(password: str, salt_hex: str | None = None) -> tuple[str, str]:
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        240_000,
    )
    return salt.hex(), digest.hex()


def infer_row_status(row: dict[str, Any], fallback: str = "draft") -> str:
    for field in row.get("fields", []):
        if not isinstance(field, dict):
            continue
        marker = f"{field.get('id', '')} {field.get('label', '')}".casefold()
        if "حالة" in marker or "status" in marker:
            value = str(field.get("value", "")).strip()
            return value or REPORT_STATUSES.get(fallback, fallback)
    return REPORT_STATUSES.get(fallback, fallback)


class AutoClosingConnection(sqlite3.Connection):
    """Keep existing commit/rollback semantics and close the connection on exit."""
    def __exit__(self, exc_type, exc_value, traceback):
        try:
            return super().__exit__(exc_type, exc_value, traceback)
        finally:
            self.close()


class ReportDatabase:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30, factory=AutoClosingConnection)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA synchronous = NORMAL")
        connection.execute("PRAGMA busy_timeout = 30000")
        connection.execute("PRAGMA temp_store = MEMORY")
        connection.execute("PRAGMA cache_size = -8000")
        return connection

    @staticmethod
    def _ensure_column(
        connection: sqlite3.Connection,
        table: str,
        column: str,
        definition: str,
    ) -> None:
        columns = {
            str(row["name"])
            for row in connection.execute(f"PRAGMA table_info({table})").fetchall()
        }
        if column not in columns:
            connection.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS reports (
                    id TEXT PRIMARY KEY,
                    report_date TEXT NOT NULL,
                    saved_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    company_name TEXT NOT NULL DEFAULT '',
                    user_name TEXT NOT NULL DEFAULT '',
                    code TEXT NOT NULL DEFAULT '',
                    report_number TEXT NOT NULL DEFAULT '',
                    row_count INTEGER NOT NULL DEFAULT 0,
                    snapshot_json TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS companies (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    logo_path TEXT NOT NULL DEFAULT '',
                    address TEXT NOT NULL DEFAULT '',
                    phone TEXT NOT NULL DEFAULT '',
                    tax_number TEXT NOT NULL DEFAULT '',
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS templates (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    config_json TEXT NOT NULL,
                    is_default INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS users (
                    id TEXT PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    display_name TEXT NOT NULL,
                    salt TEXT NOT NULL DEFAULT '',
                    password_hash TEXT NOT NULL DEFAULT '',
                    role TEXT NOT NULL DEFAULT 'viewer',
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    last_login TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS report_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_id TEXT NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                    row_index INTEGER NOT NULL DEFAULT 0,
                    category_id TEXT NOT NULL DEFAULT '',
                    category_label TEXT NOT NULL DEFAULT '',
                    person_name TEXT NOT NULL DEFAULT '',
                    work_status TEXT NOT NULL DEFAULT '',
                    total INTEGER NOT NULL DEFAULT 0,
                    search_text TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    username TEXT NOT NULL DEFAULT '',
                    action TEXT NOT NULL,
                    entity_type TEXT NOT NULL,
                    entity_id TEXT NOT NULL DEFAULT '',
                    details_json TEXT NOT NULL DEFAULT '{}'
                );
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS report_revisions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_id TEXT NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                    revision_number INTEGER NOT NULL,
                    snapshot_json TEXT NOT NULL,
                    content_hash TEXT NOT NULL,
                    changed_by TEXT NOT NULL DEFAULT '',
                    changed_at TEXT NOT NULL,
                    UNIQUE(report_id, revision_number)
                );
                CREATE TABLE IF NOT EXISTS report_approval_steps (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_id TEXT NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                    stage_order INTEGER NOT NULL,
                    stage_name TEXT NOT NULL,
                    decision TEXT NOT NULL DEFAULT 'pending'
                        CHECK(decision IN ('pending','approved','rejected')),
                    decided_by TEXT NOT NULL DEFAULT '',
                    decided_at TEXT NOT NULL DEFAULT '',
                    note TEXT NOT NULL DEFAULT '',
                    UNIQUE(report_id, stage_order)
                );
                CREATE TABLE IF NOT EXISTS sync_import_log (
                    package_id TEXT PRIMARY KEY,
                    package_hash TEXT NOT NULL,
                    imported_at TEXT NOT NULL,
                    imported_by TEXT NOT NULL DEFAULT '',
                    summary_json TEXT NOT NULL DEFAULT '{}'
                );
                CREATE TABLE IF NOT EXISTS saved_views (
                    id TEXT PRIMARY KEY,
                    owner_user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    scope TEXT NOT NULL,
                    name TEXT NOT NULL COLLATE NOCASE,
                    filters_json TEXT NOT NULL DEFAULT '{}',
                    is_shared INTEGER NOT NULL DEFAULT 0,
                    is_default INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(owner_user_id, scope, name)
                );
                CREATE TABLE IF NOT EXISTS import_profiles (
                    id TEXT PRIMARY KEY,
                    owner_user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    name TEXT NOT NULL COLLATE NOCASE,
                    file_type TEXT NOT NULL,
                    mapping_json TEXT NOT NULL DEFAULT '{}',
                    options_json TEXT NOT NULL DEFAULT '{}',
                    is_shared INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(owner_user_id, file_type, name)
                );
                """
            )
            for column, definition in (
                ("company_id", "TEXT NOT NULL DEFAULT ''"),
                ("template_id", "TEXT NOT NULL DEFAULT ''"),
                ("status", "TEXT NOT NULL DEFAULT 'draft'"),
                ("is_locked", "INTEGER NOT NULL DEFAULT 0"),
                ("locked_at", "TEXT NOT NULL DEFAULT ''"),
                ("locked_by", "TEXT NOT NULL DEFAULT ''"),
                ("fingerprint", "TEXT NOT NULL DEFAULT ''"),
                ("verification_code", "TEXT NOT NULL DEFAULT ''"),
                ("revision_number", "INTEGER NOT NULL DEFAULT 0"),
                ("archived_at", "TEXT NOT NULL DEFAULT ''"),
                ("archived_by", "TEXT NOT NULL DEFAULT ''"),
            ):
                self._ensure_column(connection, "reports", column, definition)
            connection.executescript(
                """
                CREATE INDEX IF NOT EXISTS idx_reports_date
                    ON reports(report_date, updated_at);
                CREATE INDEX IF NOT EXISTS idx_reports_company_date
                    ON reports(company_id, report_date);
                CREATE INDEX IF NOT EXISTS idx_reports_company_status_date
                    ON reports(company_id, status, report_date);
                CREATE INDEX IF NOT EXISTS idx_reports_template_date
                    ON reports(template_id, report_date);
                CREATE INDEX IF NOT EXISTS idx_reports_status
                    ON reports(status, is_locked);
                CREATE INDEX IF NOT EXISTS idx_reports_updated_at
                    ON reports(updated_at);
                CREATE INDEX IF NOT EXISTS idx_reports_archived_date
                    ON reports(archived_at, report_date);
                CREATE INDEX IF NOT EXISTS idx_reports_fingerprint
                    ON reports(fingerprint);
                CREATE INDEX IF NOT EXISTS idx_rows_report
                    ON report_rows(report_id, row_index);
                CREATE INDEX IF NOT EXISTS idx_rows_report_category
                    ON report_rows(report_id, category_id);
                CREATE INDEX IF NOT EXISTS idx_rows_report_person
                    ON report_rows(report_id, person_name);
                CREATE INDEX IF NOT EXISTS idx_rows_category
                    ON report_rows(category_id, category_label);
                CREATE INDEX IF NOT EXISTS idx_rows_person
                    ON report_rows(person_name);
                CREATE INDEX IF NOT EXISTS idx_rows_status
                    ON report_rows(work_status);
                CREATE INDEX IF NOT EXISTS idx_audit_time
                    ON audit_log(created_at);
                CREATE INDEX IF NOT EXISTS idx_audit_entity
                    ON audit_log(entity_type, entity_id);
                CREATE INDEX IF NOT EXISTS idx_audit_user_time
                    ON audit_log(username, created_at);
                CREATE INDEX IF NOT EXISTS idx_revisions_report_number
                    ON report_revisions(report_id, revision_number DESC);
                CREATE INDEX IF NOT EXISTS idx_revisions_hash
                    ON report_revisions(content_hash);
                CREATE INDEX IF NOT EXISTS idx_approval_report_stage
                    ON report_approval_steps(report_id, stage_order);
                CREATE INDEX IF NOT EXISTS idx_approval_decision
                    ON report_approval_steps(decision, report_id);
                CREATE INDEX IF NOT EXISTS idx_sync_import_time
                    ON sync_import_log(imported_at DESC);
                CREATE INDEX IF NOT EXISTS idx_saved_views_scope_owner
                    ON saved_views(scope, owner_user_id, updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_saved_views_shared
                    ON saved_views(scope, is_shared, updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_import_profiles_type_owner
                    ON import_profiles(file_type, owner_user_id, updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_import_profiles_shared
                    ON import_profiles(file_type, is_shared, updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_companies_active_name
                    ON companies(active, name);
                CREATE INDEX IF NOT EXISTS idx_templates_default_name
                    ON templates(is_default, name);
                CREATE INDEX IF NOT EXISTS idx_users_active_name
                    ON users(active, display_name);
                """
            )
            existing = connection.execute(
                "SELECT id, company_name, snapshot_json FROM reports"
            ).fetchall()
            for report in existing:
                try:
                    snapshot = json.loads(report["snapshot_json"])
                except (TypeError, json.JSONDecodeError):
                    continue
                if not isinstance(snapshot, dict):
                    continue
                company_id = str(snapshot.get("company_id", ""))
                if not company_id:
                    company_id = self._company_id_for_name(
                        connection,
                        str(report["company_name"]),
                    )
                snapshot["company_id"] = company_id
                fingerprint = report_fingerprint(snapshot, company_id)
                content_hash = report_content_hash(snapshot)
                revision_row = connection.execute(
                    "SELECT MAX(revision_number) AS number FROM report_revisions WHERE report_id = ?",
                    (report["id"],),
                ).fetchone()
                revision_number = int(revision_row["number"] or 0) if revision_row else 0
                if revision_number == 0:
                    revision_number = 1
                    connection.execute(
                        """INSERT INTO report_revisions (
                               report_id,revision_number,snapshot_json,content_hash,changed_by,changed_at
                           ) VALUES (?,?,?,?,?,?)""",
                        (
                            report["id"], revision_number,
                            json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")),
                            content_hash, "migration",
                            datetime.now().isoformat(timespec="seconds"),
                        ),
                    )
                snapshot["revision_number"] = revision_number
                snapshot["verification_code"] = content_hash
                connection.execute(
                    """
                    UPDATE reports
                    SET company_id = CASE WHEN company_id = '' THEN ? ELSE company_id END,
                        fingerprint = CASE WHEN fingerprint = '' THEN ? ELSE fingerprint END,
                        verification_code = ?, revision_number = ?, snapshot_json = ?
                    WHERE id = ?
                    """,
                    (
                        company_id, fingerprint, content_hash, revision_number,
                        json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")),
                        report["id"],
                    ),
                )
                row_exists = connection.execute(
                    "SELECT 1 FROM report_rows WHERE report_id = ? LIMIT 1",
                    (report["id"],),
                ).fetchone()
                if row_exists is None:
                    self._write_normalized_rows(
                        connection,
                        str(report["id"]),
                        snapshot,
                        str(snapshot.get("status", "draft")),
                    )
            self._bootstrap_user(connection)
            try:
                connection.execute("PRAGMA optimize")
            except sqlite3.Error:
                pass

    def _bootstrap_user(self, connection: sqlite3.Connection) -> None:
        existing = connection.execute("SELECT 1 FROM users LIMIT 1").fetchone()
        if existing:
            return
        now = datetime.now().isoformat(timespec="seconds")
        username = default_user_name() or "admin"
        connection.execute(
            """
            INSERT INTO users (
                id, username, display_name, role, active, created_at, updated_at
            ) VALUES (?, ?, ?, 'admin', 1, ?, ?)
            """,
            (uuid.uuid4().hex, username, username, now, now),
        )

    @staticmethod
    def _company_id_for_name(
        connection: sqlite3.Connection,
        company_name: str,
    ) -> str:
        clean_name = company_name.strip() or "الشركة الافتراضية"
        row = connection.execute(
            "SELECT id FROM companies WHERE name = ? COLLATE NOCASE LIMIT 1",
            (clean_name,),
        ).fetchone()
        if row:
            return str(row["id"])
        company_id = uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        connection.execute(
            """
            INSERT INTO companies (id, name, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            """,
            (company_id, clean_name, now, now),
        )
        return company_id

    @staticmethod
    def _write_normalized_rows(
        connection: sqlite3.Connection,
        report_id: str,
        snapshot: dict[str, Any],
        report_status: str,
    ) -> None:
        connection.execute("DELETE FROM report_rows WHERE report_id = ?", (report_id,))
        payload: list[tuple[Any, ...]] = []
        for index, row in enumerate(snapshot.get("rows", []), 1):
            if not isinstance(row, dict):
                continue
            field_fragments: list[str] = []
            for field in row.get("fields", []):
                if isinstance(field, dict):
                    field_fragments.extend(
                        [str(field.get("label", "")), str(field.get("value", ""))]
                    )
            search_text = " ".join(
                [
                    str(snapshot.get("company_name", "")),
                    str(snapshot.get("user_name", "")),
                    str(snapshot.get("code", "")),
                    str(snapshot.get("page", "")),
                    str(row.get("category_label", "")),
                    str(row.get("name", "")),
                    *field_fragments,
                ]
            ).casefold()
            payload.append(
                (
                    report_id,
                    index,
                    str(row.get("category", "")),
                    str(row.get("category_label", "")),
                    str(row.get("name", "")),
                    infer_row_status(row, report_status),
                    coerce_nonnegative_int(row.get("total", 0)),
                    search_text,
                )
            )
        if payload:
            connection.executemany(
                """
                INSERT INTO report_rows (
                    report_id, row_index, category_id, category_label,
                    person_name, work_status, total, search_text
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                payload,
            )

    @staticmethod
    def _write_audit(
        connection: sqlite3.Connection,
        username: str,
        action: str,
        entity_type: str,
        entity_id: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        connection.execute(
            """
            INSERT INTO audit_log (
                created_at, username, action, entity_type, entity_id, details_json
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                username,
                action,
                entity_type,
                entity_id,
                json.dumps(details or {}, ensure_ascii=False, separators=(",", ":")),
            ),
        )

    def audit(
        self,
        username: str,
        action: str,
        entity_type: str,
        entity_id: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        with self._connect() as connection:
            self._write_audit(
                connection,
                username,
                action,
                entity_type,
                entity_id,
                details,
            )

    def ensure_defaults(self, config: dict[str, Any]) -> tuple[str, str]:
        with self._connect() as connection:
            company_id = self._company_id_for_name(
                connection,
                str(config.get("identity", {}).get("company_name", "")),
            )
            now = datetime.now().isoformat(timespec="seconds")
            existing_count = int(
                connection.execute(
                    "SELECT COUNT(*) AS count FROM templates"
                ).fetchone()["count"]
            )
            for index, (theme_id, theme) in enumerate(REPORT_THEMES.items()):
                template_id = f"builtin_{theme_id}"
                exists = connection.execute(
                    "SELECT 1 FROM templates WHERE id = ?",
                    (template_id,),
                ).fetchone()
                if exists:
                    continue
                themed_config = normalize_config(copy.deepcopy(config))
                themed_config["report_theme"] = theme_id
                connection.execute(
                    """
                    INSERT INTO templates (
                        id, name, config_json, is_default, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        template_id,
                        theme["label"],
                        json.dumps(themed_config, ensure_ascii=False),
                        1 if existing_count == 0 and index == 0 else 0,
                        now,
                        now,
                    ),
                )
            template = connection.execute(
                "SELECT id FROM templates ORDER BY is_default DESC, created_at LIMIT 1"
            ).fetchone()
            template_id = str(template["id"])
        return company_id, template_id

    def save_snapshot(
        self,
        snapshot: dict[str, Any],
        report_id: str | None = None,
        *,
        saved_at: str | None = None,
        company_id: str = "",
        template_id: str = "",
        status: str = "draft",
        actor: str = "",
        allow_locked: bool = False,
    ) -> str:
        clean_id = str(report_id or uuid.uuid4().hex)
        now = datetime.now().isoformat(timespec="seconds")
        first_saved = str(saved_at or now)
        clean_status = status if status in REPORT_STATUSES else "draft"
        payload_snapshot = copy.deepcopy(snapshot)
        payload_snapshot["company_id"] = company_id
        payload_snapshot["template_id"] = template_id
        payload_snapshot["status"] = clean_status
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT saved_at,is_locked,revision_number,verification_code FROM reports WHERE id = ?",
                (clean_id,),
            ).fetchone()
            if existing and int(existing["is_locked"]) and not allow_locked:
                raise PermissionError("التقرير معتمد ومقفل. يجب إلغاء القفل أولاً.")
            if not company_id:
                company_id = self._company_id_for_name(
                    connection,
                    str(payload_snapshot.get("company_name", "")),
                )
                payload_snapshot["company_id"] = company_id
            content_hash = report_content_hash(payload_snapshot)
            latest_revision = connection.execute(
                """SELECT revision_number,content_hash FROM report_revisions
                   WHERE report_id=? ORDER BY revision_number DESC LIMIT 1""",
                (clean_id,),
            ).fetchone()
            if latest_revision and str(latest_revision["content_hash"]) == content_hash:
                revision_number = int(latest_revision["revision_number"])
                revision_changed = False
            else:
                revision_number = (
                    int(latest_revision["revision_number"]) + 1
                    if latest_revision else 1
                )
                revision_changed = True
            payload_snapshot["revision_number"] = revision_number
            payload_snapshot["verification_code"] = content_hash
            payload = json.dumps(
                payload_snapshot,
                ensure_ascii=False,
                separators=(",", ":"),
            )
            fingerprint = report_fingerprint(payload_snapshot, company_id)
            connection.execute(
                """
                INSERT INTO reports (
                    id, report_date, saved_at, updated_at,
                    company_name, user_name, code, report_number,
                    row_count, snapshot_json, company_id, template_id,
                    status, fingerprint, verification_code, revision_number
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    report_date = excluded.report_date,
                    updated_at = excluded.updated_at,
                    company_name = excluded.company_name,
                    user_name = excluded.user_name,
                    code = excluded.code,
                    report_number = excluded.report_number,
                    row_count = excluded.row_count,
                    snapshot_json = excluded.snapshot_json,
                    company_id = excluded.company_id,
                    template_id = excluded.template_id,
                    status = excluded.status,
                    fingerprint = excluded.fingerprint,
                    verification_code = excluded.verification_code,
                    revision_number = excluded.revision_number
                """,
                (
                    clean_id,
                    str(payload_snapshot.get("date", "")),
                    str(existing["saved_at"] if existing else first_saved),
                    now,
                    str(payload_snapshot.get("company_name", "")),
                    str(payload_snapshot.get("user_name", "")),
                    str(payload_snapshot.get("code", "")),
                    str(payload_snapshot.get("page", "")),
                    len(payload_snapshot.get("rows", [])),
                    payload,
                    company_id,
                    template_id,
                    clean_status,
                    fingerprint,
                    content_hash,
                    revision_number,
                ),
            )
            self._write_normalized_rows(
                connection,
                clean_id,
                payload_snapshot,
                clean_status,
            )
            if revision_changed:
                connection.execute(
                    """INSERT INTO report_revisions (
                           report_id,revision_number,snapshot_json,content_hash,changed_by,changed_at
                       ) VALUES (?,?,?,?,?,?)""",
                    (
                        clean_id, revision_number, payload, content_hash,
                        actor or str(payload_snapshot.get("user_name", "")), now,
                    ),
                )
                connection.execute(
                    """UPDATE report_approval_steps
                       SET decision='pending',decided_by='',decided_at='',note=''
                       WHERE report_id=?""",
                    (clean_id,),
                )
            self._write_audit(
                connection,
                actor or str(payload_snapshot.get("user_name", "")),
                "update" if existing else "create",
                "report",
                clean_id,
                {
                    "date": payload_snapshot.get("date", ""),
                    "code": payload_snapshot.get("code", ""),
                    "rows": len(payload_snapshot.get("rows", [])),
                    "revision": revision_number,
                    "verification": content_hash[:16],
                },
            )
        return clean_id

    def find_duplicate(
        self,
        snapshot: dict[str, Any],
        company_id: str = "",
        exclude_id: str | None = None,
    ) -> dict[str, Any] | None:
        fingerprint = report_fingerprint(snapshot, company_id)
        query = """
            SELECT id, report_date, code, company_name, updated_at
            FROM reports
            WHERE fingerprint = ?
        """
        parameters: list[Any] = [fingerprint]
        if exclude_id:
            query += " AND id <> ?"
            parameters.append(exclude_id)
        query += " ORDER BY updated_at DESC LIMIT 1"
        with self._connect() as connection:
            row = connection.execute(query, parameters).fetchone()
        return dict(row) if row else None

    def list_reports(
        self,
        start_date: str | None = None,
        end_date: str | None = None,
        *,
        search: str = "",
        company_id: str = "",
        category_id: str = "",
        person_name: str = "",
        status: str = "",
        include_archived: bool = False,
        archived_only: bool = False,
    ) -> list[dict[str, Any]]:
        conditions: list[str] = []
        if archived_only:
            conditions.append("r.archived_at <> ''")
        elif not include_archived:
            conditions.append("r.archived_at = ''")
        parameters: list[Any] = []
        if start_date:
            conditions.append("r.report_date >= ?")
            parameters.append(start_date)
        if end_date:
            conditions.append("r.report_date <= ?")
            parameters.append(end_date)
        if company_id:
            conditions.append("r.company_id = ?")
            parameters.append(company_id)
        if status:
            conditions.append("r.status = ?")
            parameters.append(status)
        if category_id:
            conditions.append(
                "EXISTS (SELECT 1 FROM report_rows rr "
                "WHERE rr.report_id = r.id AND rr.category_id = ?)"
            )
            parameters.append(category_id)
        if person_name:
            conditions.append(
                "EXISTS (SELECT 1 FROM report_rows rr "
                "WHERE rr.report_id = r.id AND rr.person_name = ?)"
            )
            parameters.append(person_name)
        if search.strip():
            pattern = f"%{search.strip().casefold()}%"
            conditions.append(
                """
                (
                    lower(r.company_name) LIKE ? OR lower(r.user_name) LIKE ?
                    OR lower(r.code) LIKE ? OR lower(r.report_number) LIKE ?
                    OR EXISTS (
                        SELECT 1 FROM report_rows rr
                        WHERE rr.report_id = r.id AND rr.search_text LIKE ?
                    )
                )
                """
            )
            parameters.extend([pattern] * 5)
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        query = f"""
            SELECT r.id, r.report_date, r.saved_at, r.updated_at, r.company_name,
                   r.user_name, r.code, r.report_number, r.row_count,
                   r.company_id, r.template_id, r.status, r.is_locked,
                   r.locked_at, r.locked_by, r.verification_code, r.revision_number,
                   r.archived_at, r.archived_by
            FROM reports r
            {where}
            ORDER BY r.report_date DESC, r.updated_at DESC
        """
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(query, parameters).fetchall()
            ]

    def get_report(self, report_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT id, report_date, saved_at, updated_at, snapshot_json,
                       company_id, template_id, status, is_locked, locked_at, locked_by,
                       verification_code, revision_number
                FROM reports
                WHERE id = ?
                """,
                (report_id,),
            ).fetchone()
        if row is None:
            return None
        try:
            snapshot = json.loads(row["snapshot_json"])
        except (TypeError, json.JSONDecodeError):
            return None
        if not isinstance(snapshot, dict):
            return None
        return {
            "id": row["id"],
            "report_date": row["report_date"],
            "saved_at": row["saved_at"],
            "updated_at": row["updated_at"],
            "company_id": row["company_id"],
            "template_id": row["template_id"],
            "status": row["status"],
            "is_locked": bool(row["is_locked"]),
            "locked_at": row["locked_at"],
            "locked_by": row["locked_by"],
            "verification_code": row["verification_code"],
            "revision_number": int(row["revision_number"] or 0),
            "snapshot": snapshot,
        }

    def set_report_lock(
        self,
        report_id: str,
        locked: bool,
        actor: str,
    ) -> bool:
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            cursor = connection.execute(
                """
                UPDATE reports
                SET is_locked = ?, locked_at = ?, locked_by = ?,
                    status = CASE WHEN ? = 1 THEN 'approved' ELSE status END,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    1 if locked else 0,
                    now if locked else "",
                    actor if locked else "",
                    1 if locked else 0,
                    now,
                    report_id,
                ),
            )
            if cursor.rowcount:
                self._write_audit(
                    connection,
                    actor,
                    "lock" if locked else "unlock",
                    "report",
                    report_id,
                )
        return bool(cursor.rowcount)

    def list_report_revisions(self, report_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT revision_number,content_hash,changed_by,changed_at
                   FROM report_revisions WHERE report_id=?
                   ORDER BY revision_number DESC""",
                (report_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_report_revision(
        self, report_id: str, revision_number: int
    ) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT revision_number,snapshot_json,content_hash,changed_by,changed_at
                   FROM report_revisions WHERE report_id=? AND revision_number=?""",
                (report_id, int(revision_number)),
            ).fetchone()
        if row is None:
            return None
        try:
            snapshot = json.loads(row["snapshot_json"])
        except (TypeError, json.JSONDecodeError):
            return None
        result = dict(row)
        result.pop("snapshot_json", None)
        result["snapshot"] = snapshot
        return result

    def restore_report_revision(
        self, report_id: str, revision_number: int, actor: str
    ) -> str:
        revision = self.get_report_revision(report_id, revision_number)
        current = self.get_report(report_id)
        if revision is None or current is None:
            raise ValueError("تعذر العثور على إصدار التقرير المطلوب.")
        snapshot = copy.deepcopy(revision["snapshot"])
        snapshot.pop("revision_number", None)
        snapshot.pop("verification_code", None)
        snapshot["status"] = "draft"
        with self._connect() as connection:
            connection.execute(
                """UPDATE reports SET is_locked=0,locked_at='',locked_by='',status='draft'
                   WHERE id=?""",
                (report_id,),
            )
            self._write_audit(
                connection,
                actor,
                "restore_revision",
                "report",
                report_id,
                {"source_revision": int(revision_number)},
            )
        return self.save_snapshot(
            snapshot,
            report_id,
            saved_at=current["saved_at"],
            company_id=current["company_id"],
            template_id=current["template_id"],
            status="draft",
            actor=actor,
            allow_locked=True,
        )

    def verify_report(self, report_id: str) -> dict[str, Any]:
        report = self.get_report(report_id)
        if report is None:
            return {"valid": False, "reason": "التقرير غير موجود.", "code": ""}
        computed = report_content_hash(report["snapshot"])
        stored = str(report.get("verification_code") or "")
        with self._connect() as connection:
            revision = connection.execute(
                """SELECT content_hash FROM report_revisions
                   WHERE report_id=? ORDER BY revision_number DESC LIMIT 1""",
                (report_id,),
            ).fetchone()
        revision_hash = str(revision["content_hash"] or "") if revision else ""
        valid = bool(stored) and computed == stored and revision_hash == stored
        return {
            "valid": valid,
            "code": stored,
            "short_code": stored[:16].upper(),
            "computed": computed,
            "revision": report.get("revision_number", 0),
            "reason": "البصمة مطابقة وسجل الإصدار سليم." if valid else "البصمة لا تطابق محتوى التقرير الحالي.",
        }

    def configure_approval_workflow(
        self, report_id: str, stage_names: list[str], actor: str
    ) -> list[dict[str, Any]]:
        clean_names: list[str] = []
        for raw_name in stage_names:
            name = str(raw_name).strip()
            if name and name.casefold() not in {item.casefold() for item in clean_names}:
                clean_names.append(name[:80])
        if not clean_names:
            raise ValueError("أضف مرحلة موافقة واحدة على الأقل.")
        if len(clean_names) > 8:
            raise ValueError("الحد الأعلى لمسار الموافقات ثماني مراحل.")
        with self._connect() as connection:
            if connection.execute("SELECT 1 FROM reports WHERE id=?", (report_id,)).fetchone() is None:
                raise ValueError("التقرير غير موجود.")
            connection.execute("DELETE FROM report_approval_steps WHERE report_id=?", (report_id,))
            connection.executemany(
                """INSERT INTO report_approval_steps (
                       report_id,stage_order,stage_name,decision
                   ) VALUES (?,?,?,'pending')""",
                [
                    (report_id, index, name)
                    for index, name in enumerate(clean_names, 1)
                ],
            )
            connection.execute(
                "UPDATE reports SET is_locked=0,locked_at='',locked_by='',status='completed' WHERE id=?",
                (report_id,),
            )
            self._write_audit(
                connection, actor, "configure_approvals", "report", report_id,
                {"stages": clean_names},
            )
        return self.list_report_approvals(report_id)

    def list_report_approvals(self, report_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT stage_order,stage_name,decision,decided_by,decided_at,note
                   FROM report_approval_steps WHERE report_id=? ORDER BY stage_order""",
                (report_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def decide_report_approval(
        self,
        report_id: str,
        stage_order: int,
        decision: str,
        actor: str,
        note: str = "",
    ) -> list[dict[str, Any]]:
        clean_decision = str(decision).strip().lower()
        if clean_decision not in {"approved", "rejected"}:
            raise ValueError("قرار الموافقة غير صحيح.")
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            steps = connection.execute(
                """SELECT stage_order,decision FROM report_approval_steps
                   WHERE report_id=? ORDER BY stage_order""",
                (report_id,),
            ).fetchall()
            if not steps:
                raise ValueError("لم يتم إعداد مسار موافقات لهذا التقرير.")
            target = next((row for row in steps if int(row["stage_order"]) == int(stage_order)), None)
            if target is None:
                raise ValueError("مرحلة الموافقة غير موجودة.")
            if clean_decision == "approved" and any(
                int(row["stage_order"]) < int(stage_order) and row["decision"] != "approved"
                for row in steps
            ):
                raise PermissionError("يجب اعتماد المراحل السابقة أولاً.")
            connection.execute(
                """UPDATE report_approval_steps
                   SET decision=?,decided_by=?,decided_at=?,note=?
                   WHERE report_id=? AND stage_order=?""",
                (clean_decision, actor, now, str(note).strip()[:500], report_id, int(stage_order)),
            )
            if clean_decision == "rejected":
                connection.execute(
                    """UPDATE report_approval_steps
                       SET decision='pending',decided_by='',decided_at='',note=''
                       WHERE report_id=? AND stage_order>?""",
                    (report_id, int(stage_order)),
                )
                connection.execute(
                    """UPDATE reports SET status='completed',is_locked=0,locked_at='',locked_by='',updated_at=?
                       WHERE id=?""",
                    (now, report_id),
                )
            else:
                pending = int(
                    connection.execute(
                        "SELECT COUNT(*) FROM report_approval_steps WHERE report_id=? AND decision<>'approved'",
                        (report_id,),
                    ).fetchone()[0]
                )
                if pending == 0:
                    connection.execute(
                        """UPDATE reports SET status='approved',is_locked=1,locked_at=?,locked_by=?,updated_at=?
                           WHERE id=?""",
                        (now, actor, now, report_id),
                    )
            self._write_audit(
                connection, actor, f"approval_{clean_decision}", "report", report_id,
                {"stage": int(stage_order), "note": str(note).strip()[:120]},
            )
        return self.list_report_approvals(report_id)

    def export_sync_package(
        self, destination: Path | str, report_ids: list[str], actor: str
    ) -> Path:
        clean_ids = list(dict.fromkeys(str(value) for value in report_ids if value))
        if not clean_ids:
            raise ValueError("اختر تقريرًا واحدًا على الأقل للمزامنة.")
        reports: list[dict[str, Any]] = []
        for report_id in clean_ids:
            record = self.get_report(report_id)
            if record is None:
                continue
            reports.append(
                {
                    "id": report_id,
                    "saved_at": record["saved_at"],
                    "updated_at": record["updated_at"],
                    "company_id": record["company_id"],
                    "template_id": record["template_id"],
                    "status": record["status"],
                    "is_locked": record["is_locked"],
                    "verification_code": record["verification_code"],
                    "revision_number": record["revision_number"],
                    "snapshot": record["snapshot"],
                    "approvals": self.list_report_approvals(report_id),
                }
            )
        if not reports:
            raise ValueError("لم يتم العثور على التقارير المحددة.")
        package_id = uuid.uuid4().hex
        payload_bytes = json.dumps(
            {"schema": 1, "reports": reports},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        payload_hash = hashlib.sha256(payload_bytes).hexdigest()
        manifest = {
            "package_id": package_id,
            "schema": 1,
            "app_version": APP_VERSION,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "created_by": actor,
            "report_count": len(reports),
            "payload_sha256": payload_hash,
        }
        target = Path(destination)
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_suffix(target.suffix + ".tmp")
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            archive.writestr("reports.json", payload_bytes)
        os.replace(temporary, target)
        self.audit(
            actor, "export_sync", "sync_package", package_id,
            {"reports": len(reports), "path": str(target), "sha256": payload_hash},
        )
        return target

    def import_sync_package(self, source: Path | str, actor: str) -> dict[str, Any]:
        path = Path(source)
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())
            if names != {"manifest.json", "reports.json"}:
                raise ValueError("حزمة المزامنة تحتوي ملفات غير متوقعة.")
            for name in names:
                info = archive.getinfo(name)
                if info.file_size > 100 * 1024 * 1024:
                    raise ValueError("حزمة المزامنة أكبر من الحد الآمن.")
            manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
            payload_bytes = archive.read("reports.json")
        if not isinstance(manifest, dict) or int(manifest.get("schema", 0)) != 1:
            raise ValueError("إصدار حزمة المزامنة غير مدعوم.")
        payload_hash = hashlib.sha256(payload_bytes).hexdigest()
        if not secrets.compare_digest(payload_hash, str(manifest.get("payload_sha256", ""))):
            raise ValueError("فشل التحقق من سلامة حزمة المزامنة.")
        package_id = str(manifest.get("package_id", "")).strip()
        if not package_id:
            raise ValueError("معرّف حزمة المزامنة مفقود.")
        with self._connect() as connection:
            duplicate = connection.execute(
                "SELECT summary_json FROM sync_import_log WHERE package_id=?", (package_id,)
            ).fetchone()
        if duplicate:
            summary = json.loads(duplicate["summary_json"])
            summary["duplicate_package"] = True
            return summary
        payload = json.loads(payload_bytes.decode("utf-8"))
        if not isinstance(payload, dict) or not isinstance(payload.get("reports"), list):
            raise ValueError("بيانات حزمة المزامنة غير صالحة.")
        summary: dict[str, Any] = {
            "package_id": package_id,
            "created": 0,
            "updated": 0,
            "unchanged": 0,
            "conflicts": 0,
            "duplicate_package": False,
        }
        for item in payload["reports"]:
            if not isinstance(item, dict) or not isinstance(item.get("snapshot"), dict):
                summary["conflicts"] += 1
                continue
            report_id = str(item.get("id", "")).strip()
            if not report_id:
                summary["conflicts"] += 1
                continue
            snapshot = item["snapshot"]
            remote_hash = report_content_hash(snapshot)
            local = self.get_report(report_id)
            if local and str(local.get("verification_code", "")) == remote_hash:
                summary["unchanged"] += 1
                continue
            if local and (
                local.get("is_locked")
                or str(local.get("updated_at", "")) >= str(item.get("updated_at", ""))
            ):
                summary["conflicts"] += 1
                continue
            self.save_snapshot(
                snapshot,
                report_id,
                saved_at=str(item.get("saved_at", "")) or None,
                company_id=str(item.get("company_id", "")),
                template_id=str(item.get("template_id", "")),
                status=str(item.get("status", "draft")),
                actor=actor,
                allow_locked=True,
            )
            approvals = item.get("approvals", [])
            if isinstance(approvals, list) and approvals:
                names = [str(step.get("stage_name", "")) for step in approvals if isinstance(step, dict)]
                if any(name.strip() for name in names):
                    self.configure_approval_workflow(report_id, names, actor)
            summary["updated" if local else "created"] += 1
        imported_at = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO sync_import_log (
                       package_id,package_hash,imported_at,imported_by,summary_json
                   ) VALUES (?,?,?,?,?)""",
                (package_id, payload_hash, imported_at, actor, json.dumps(summary, ensure_ascii=False)),
            )
            self._write_audit(
                connection, actor, "import_sync", "sync_package", package_id, summary,
            )
        return summary

    def archive_reports(
        self, report_ids: list[str], *, actor: str = "", allow_locked: bool = False
    ) -> int:
        clean_ids = [str(report_id) for report_id in report_ids if report_id]
        if not clean_ids:
            return 0
        placeholders = ",".join("?" for _ in clean_ids)
        with self._connect() as connection:
            if not allow_locked:
                locked = connection.execute(
                    f"SELECT COUNT(*) AS count FROM reports WHERE id IN ({placeholders}) AND is_locked=1",
                    clean_ids,
                ).fetchone()
                if locked and int(locked["count"]):
                    raise PermissionError("يتضمن التحديد تقريراً معتمداً ومقفلاً. ألغِ القفل أولاً.")
            now = datetime.now().isoformat(timespec="seconds")
            cursor = connection.execute(
                f"UPDATE reports SET archived_at=?, archived_by=? WHERE id IN ({placeholders}) AND archived_at=''",
                [now, actor, *clean_ids],
            )
            changed = max(0, cursor.rowcount)
            if changed:
                self._write_audit(connection, actor, "archive", "report", ",".join(clean_ids), {"count": changed})
            return changed

    def restore_archived_reports(self, report_ids: list[str], *, actor: str = "") -> int:
        clean_ids = [str(report_id) for report_id in report_ids if report_id]
        if not clean_ids:
            return 0
        placeholders = ",".join("?" for _ in clean_ids)
        with self._connect() as connection:
            cursor = connection.execute(
                f"UPDATE reports SET archived_at='', archived_by='' WHERE id IN ({placeholders}) AND archived_at<>''",
                clean_ids,
            )
            changed = max(0, cursor.rowcount)
            if changed:
                self._write_audit(connection, actor, "restore_archive", "report", ",".join(clean_ids), {"count": changed})
            return changed

    def purge_archived_reports(self, report_ids: list[str], *, actor: str = "") -> int:
        clean_ids = [str(report_id) for report_id in report_ids if report_id]
        if not clean_ids:
            return 0
        placeholders = ",".join("?" for _ in clean_ids)
        with self._connect() as connection:
            cursor = connection.execute(
                f"DELETE FROM reports WHERE id IN ({placeholders}) AND archived_at<>''", clean_ids
            )
            changed = max(0, cursor.rowcount)
            if changed:
                self._write_audit(connection, actor, "purge", "report", ",".join(clean_ids), {"count": changed})
            return changed

    def delete_reports(
        self,
        report_ids: list[str],
        *,
        actor: str = "",
        allow_locked: bool = False,
    ) -> int:
        clean_ids = [str(report_id) for report_id in report_ids if report_id]
        if not clean_ids:
            return 0
        placeholders = ",".join("?" for _ in clean_ids)
        with self._connect() as connection:
            if not allow_locked:
                locked = connection.execute(
                    f"""
                    SELECT COUNT(*) AS count FROM reports
                    WHERE id IN ({placeholders}) AND is_locked = 1
                    """,
                    clean_ids,
                ).fetchone()
                if locked and int(locked["count"]):
                    raise PermissionError(
                        "يتضمن التحديد تقريراً معتمداً ومقفلاً. ألغِ القفل أولاً."
                    )
            cursor = connection.execute(
                f"DELETE FROM reports WHERE id IN ({placeholders})",
                clean_ids,
            )
            deleted = max(0, cursor.rowcount)
            if deleted:
                self._write_audit(
                    connection,
                    actor,
                    "delete",
                    "report",
                    ",".join(clean_ids),
                    {"count": deleted},
                )
            return deleted

    def count(self, company_id: str = "", *, include_archived: bool = False) -> int:
        query = "SELECT COUNT(*) AS count FROM reports"
        parameters: tuple[Any, ...] = ()
        conditions: list[str] = []
        if not include_archived:
            conditions.append("archived_at = ''")
        if company_id:
            conditions.append("company_id = ?")
            parameters = (company_id,)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        with self._connect() as connection:
            row = connection.execute(query, parameters).fetchone()
        return int(row["count"]) if row else 0

    def list_filter_values(self, company_id: str = "") -> dict[str, list[dict[str, str]]]:
        parameters: tuple[Any, ...] = (company_id,) if company_id else ()
        report_filter = "WHERE r.company_id = ?" if company_id else ""
        row_filter = (
            "JOIN reports r ON r.id = rr.report_id WHERE r.company_id = ?"
            if company_id
            else ""
        )
        with self._connect() as connection:
            categories = [
                {"id": str(row["category_id"]), "label": str(row["category_label"])}
                for row in connection.execute(
                    f"""
                    SELECT rr.category_id, MAX(rr.category_label) AS category_label
                    FROM report_rows rr {row_filter}
                    GROUP BY rr.category_id ORDER BY category_label
                    """,
                    parameters,
                ).fetchall()
            ]
            people = [
                {"id": str(row["person_name"]), "label": str(row["person_name"])}
                for row in connection.execute(
                    f"""
                    SELECT rr.person_name
                    FROM report_rows rr {row_filter}
                    WHERE rr.person_name <> ''
                    GROUP BY rr.person_name ORDER BY rr.person_name
                    """
                    if not company_id
                    else """
                    SELECT rr.person_name
                    FROM report_rows rr
                    JOIN reports r ON r.id = rr.report_id
                    WHERE r.company_id = ? AND rr.person_name <> ''
                    GROUP BY rr.person_name ORDER BY rr.person_name
                    """,
                    parameters,
                ).fetchall()
            ]
            users = [
                {"id": str(row["user_name"]), "label": str(row["user_name"])}
                for row in connection.execute(
                    f"""
                    SELECT r.user_name FROM reports r {report_filter}
                    WHERE r.user_name <> ''
                    GROUP BY r.user_name ORDER BY r.user_name
                    """
                    if not company_id
                    else """
                    SELECT r.user_name FROM reports r
                    WHERE r.company_id = ? AND r.user_name <> ''
                    GROUP BY r.user_name ORDER BY r.user_name
                    """,
                    parameters,
                ).fetchall()
            ]
        return {"categories": categories, "people": people, "users": users}

    def list_companies(self, active_only: bool = False) -> list[dict[str, Any]]:
        query = "SELECT * FROM companies"
        if active_only:
            query += " WHERE active = 1"
        query += " ORDER BY active DESC, name"
        with self._connect() as connection:
            return [dict(row) for row in connection.execute(query).fetchall()]

    def get_company(self, company_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM companies WHERE id = ?",
                (company_id,),
            ).fetchone()
        return dict(row) if row else None

    def save_company(
        self,
        data: dict[str, Any],
        actor: str,
        company_id: str = "",
    ) -> str:
        clean_id = company_id or uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            exists = connection.execute(
                "SELECT 1 FROM companies WHERE id = ?",
                (clean_id,),
            ).fetchone()
            connection.execute(
                """
                INSERT INTO companies (
                    id, name, logo_path, address, phone, tax_number,
                    active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    name = excluded.name, logo_path = excluded.logo_path,
                    address = excluded.address, phone = excluded.phone,
                    tax_number = excluded.tax_number, active = excluded.active,
                    updated_at = excluded.updated_at
                """,
                (
                    clean_id,
                    str(data.get("name", "")).strip() or "شركة بدون اسم",
                    str(data.get("logo_path", "")),
                    str(data.get("address", "")),
                    str(data.get("phone", "")),
                    str(data.get("tax_number", "")),
                    1 if data.get("active", True) else 0,
                    now,
                    now,
                ),
            )
            self._write_audit(
                connection,
                actor,
                "update" if exists else "create",
                "company",
                clean_id,
                {"name": data.get("name", "")},
            )
        return clean_id

    def deactivate_company(self, company_id: str, actor: str) -> bool:
        with self._connect() as connection:
            cursor = connection.execute(
                "UPDATE companies SET active = 0, updated_at = ? WHERE id = ?",
                (datetime.now().isoformat(timespec="seconds"), company_id),
            )
            if cursor.rowcount:
                self._write_audit(
                    connection,
                    actor,
                    "deactivate",
                    "company",
                    company_id,
                )
        return bool(cursor.rowcount)

    def list_templates(self) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT id, name, config_json, is_default, created_at, updated_at
                FROM templates ORDER BY is_default DESC, name
                """
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            try:
                item["config"] = normalize_config(json.loads(item.pop("config_json")))
            except (TypeError, json.JSONDecodeError):
                item["config"] = normalize_config(DEFAULT_CONFIG)
            result.append(item)
        return result

    def save_template(
        self,
        name: str,
        config: dict[str, Any],
        actor: str,
        template_id: str = "",
        is_default: bool = False,
    ) -> str:
        clean_id = template_id or uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            exists = connection.execute(
                "SELECT 1 FROM templates WHERE id = ?",
                (clean_id,),
            ).fetchone()
            if is_default:
                connection.execute("UPDATE templates SET is_default = 0")
            connection.execute(
                """
                INSERT INTO templates (
                    id, name, config_json, is_default, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    name = excluded.name, config_json = excluded.config_json,
                    is_default = excluded.is_default, updated_at = excluded.updated_at
                """,
                (
                    clean_id,
                    name.strip() or "قالب بدون اسم",
                    json.dumps(normalize_config(config), ensure_ascii=False),
                    1 if is_default else 0,
                    now,
                    now,
                ),
            )
            self._write_audit(
                connection,
                actor,
                "update" if exists else "create",
                "template",
                clean_id,
                {"name": name},
            )
        return clean_id

    def delete_template(self, template_id: str, actor: str) -> bool:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT is_default FROM templates WHERE id = ?",
                (template_id,),
            ).fetchone()
            if row and int(row["is_default"]):
                raise ValueError("لا يمكن حذف القالب الافتراضي.")
            cursor = connection.execute(
                "DELETE FROM templates WHERE id = ?",
                (template_id,),
            )
            if cursor.rowcount:
                self._write_audit(
                    connection,
                    actor,
                    "delete",
                    "template",
                    template_id,
                )
        return bool(cursor.rowcount)

    def list_users(self, active_only: bool = False) -> list[dict[str, Any]]:
        query = """
            SELECT id, username, display_name, role, active,
                   created_at, updated_at, last_login,
                   CASE WHEN password_hash = '' THEN 0 ELSE 1 END AS has_password
            FROM users
        """
        if active_only:
            query += " WHERE active = 1"
        query += " ORDER BY active DESC, display_name"
        with self._connect() as connection:
            return [dict(row) for row in connection.execute(query).fetchall()]

    def save_user(
        self,
        username: str,
        display_name: str,
        role: str,
        active: bool,
        actor: str,
        *,
        user_id: str = "",
        password: str | None = None,
    ) -> str:
        clean_role = role if role in ROLE_LABELS else "viewer"
        clean_username = username.strip()
        if not clean_username:
            raise ValueError("اسم المستخدم مطلوب.")
        clean_id = user_id or uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT salt, password_hash, role, active FROM users WHERE id = ?",
                (clean_id,),
            ).fetchone()
            if (
                existing
                and str(existing["role"]) == "admin"
                and int(existing["active"])
                and (clean_role != "admin" or not active)
            ):
                other_admin = connection.execute(
                    """
                    SELECT 1 FROM users
                    WHERE id <> ? AND role = 'admin' AND active = 1
                    LIMIT 1
                    """,
                    (clean_id,),
                ).fetchone()
                if other_admin is None:
                    raise ValueError("يجب الإبقاء على مدير نظام فعّال واحد على الأقل.")
            salt = str(existing["salt"]) if existing else ""
            digest = str(existing["password_hash"]) if existing else ""
            if password is not None:
                salt, digest = password_digest(password)
            connection.execute(
                """
                INSERT INTO users (
                    id, username, display_name, salt, password_hash,
                    role, active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    username = excluded.username,
                    display_name = excluded.display_name,
                    salt = excluded.salt,
                    password_hash = excluded.password_hash,
                    role = excluded.role,
                    active = excluded.active,
                    updated_at = excluded.updated_at
                """,
                (
                    clean_id,
                    clean_username,
                    display_name.strip() or clean_username,
                    salt,
                    digest,
                    clean_role,
                    1 if active else 0,
                    now,
                    now,
                ),
            )
            self._write_audit(
                connection,
                actor,
                "update" if existing else "create",
                "user",
                clean_id,
                {"username": username, "role": clean_role, "active": active},
            )
        return clean_id

    def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM users WHERE username = ? AND active = 1",
                (username.strip(),),
            ).fetchone()
            if row is None:
                return None
            if row["password_hash"]:
                _salt, digest = password_digest(password, str(row["salt"]))
                if not secrets.compare_digest(digest, str(row["password_hash"])):
                    return None
            connection.execute(
                "UPDATE users SET last_login = ? WHERE id = ?",
                (datetime.now().isoformat(timespec="seconds"), row["id"]),
            )
        return {
            "id": row["id"],
            "username": row["username"],
            "display_name": row["display_name"],
            "role": row["role"],
        }

    def bootstrap_login(self) -> dict[str, Any] | None:
        username = default_user_name() or "admin"
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT id, username, display_name, role
                FROM users
                WHERE username = ? AND active = 1 AND password_hash = ''
                """,
                (username,),
            ).fetchone()
        return dict(row) if row else None

    def list_audit(
        self,
        search: str = "",
        limit: int = 1000,
    ) -> list[dict[str, Any]]:
        query = """
            SELECT id, created_at, username, action, entity_type,
                   entity_id, details_json
            FROM audit_log
        """
        parameters: list[Any] = []
        if search.strip():
            query += """
                WHERE lower(username) LIKE ? OR lower(action) LIKE ?
                   OR lower(entity_type) LIKE ? OR lower(details_json) LIKE ?
            """
            pattern = f"%{search.strip().casefold()}%"
            parameters.extend([pattern] * 4)
        query += " ORDER BY id DESC LIMIT ?"
        parameters.append(max(1, min(limit, 10_000)))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(query, parameters).fetchall()
            ]

    def save_view(
        self,
        scope: str,
        name: str,
        filters: dict[str, Any],
        actor: dict[str, Any],
        *,
        is_shared: bool = False,
        is_default: bool = False,
    ) -> str:
        clean_scope = str(scope or "").strip().casefold()
        clean_name = str(name or "").strip()
        if clean_scope not in {"history", "dashboard"}:
            raise ValueError("نوع المنظر غير مدعوم.")
        if not clean_name or len(clean_name) > 80:
            raise ValueError("اسم المنظر مطلوب وبحد أقصى 80 حرفاً.")
        if not isinstance(filters, dict):
            raise ValueError("مرشحات المنظر غير صالحة.")
        owner_id = str(actor.get("id") or "").strip()
        if not owner_id:
            raise PermissionError("يلزم تسجيل الدخول لحفظ المنظر.")
        if is_shared and str(actor.get("role")) != "admin":
            raise PermissionError("مشاركة المناظر متاحة لمدير النظام فقط.")
        payload = json.dumps(filters, ensure_ascii=False, sort_keys=True)
        if len(payload.encode("utf-8")) > 32_000:
            raise ValueError("حجم مرشحات المنظر أكبر من الحد المسموح.")
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            if is_default:
                connection.execute(
                    "UPDATE saved_views SET is_default=0 WHERE owner_user_id=? AND scope=?",
                    (owner_id, clean_scope),
                )
            existing = connection.execute(
                "SELECT id FROM saved_views WHERE owner_user_id=? AND scope=? AND name=?",
                (owner_id, clean_scope, clean_name),
            ).fetchone()
            view_id = str(existing["id"]) if existing else uuid.uuid4().hex
            connection.execute(
                """INSERT INTO saved_views(
                       id,owner_user_id,scope,name,filters_json,is_shared,is_default,created_at,updated_at
                   ) VALUES(?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(id) DO UPDATE SET
                       name=excluded.name,filters_json=excluded.filters_json,
                       is_shared=excluded.is_shared,is_default=excluded.is_default,
                       updated_at=excluded.updated_at""",
                (
                    view_id, owner_id, clean_scope, clean_name, payload,
                    int(is_shared), int(is_default),
                    now if not existing else now, now,
                ),
            )
            self._write_audit(
                connection,
                str(actor.get("username") or actor.get("display_name") or ""),
                "update" if existing else "create",
                "saved_view",
                view_id,
                {"scope": clean_scope, "name": clean_name, "shared": bool(is_shared)},
            )
        return view_id

    def list_views(self, scope: str, actor: dict[str, Any]) -> list[dict[str, Any]]:
        owner_id = str(actor.get("id") or "").strip()
        if not owner_id:
            return []
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT sv.*,u.display_name AS owner_name
                   FROM saved_views sv JOIN users u ON u.id=sv.owner_user_id
                   WHERE sv.scope=? AND (sv.owner_user_id=? OR sv.is_shared=1)
                   ORDER BY sv.is_default DESC,sv.is_shared ASC,sv.name COLLATE NOCASE""",
                (str(scope or "").strip().casefold(), owner_id),
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            try:
                filters = json.loads(str(item.pop("filters_json") or "{}"))
            except (TypeError, ValueError):
                filters = {}
            item["filters"] = filters if isinstance(filters, dict) else {}
            result.append(item)
        return result

    def delete_view(self, view_id: str, actor: dict[str, Any]) -> None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id,owner_user_id,scope,name FROM saved_views WHERE id=?",
                (str(view_id),),
            ).fetchone()
            if row is None:
                raise ValueError("المنظر غير موجود.")
            is_owner = str(row["owner_user_id"]) == str(actor.get("id") or "")
            if not is_owner and str(actor.get("role")) != "admin":
                raise PermissionError("لا تملك صلاحية حذف هذا المنظر.")
            connection.execute("DELETE FROM saved_views WHERE id=?", (str(view_id),))
            self._write_audit(
                connection,
                str(actor.get("username") or actor.get("display_name") or ""),
                "delete",
                "saved_view",
                str(view_id),
                {"scope": row["scope"], "name": row["name"]},
            )

    @staticmethod
    def _profile_file_type(file_type: str) -> str:
        clean = str(file_type or "").strip().casefold()
        if clean and not clean.startswith("."):
            clean = f".{clean}"
        if clean not in {".xlsx", ".csv", ".tsv", ".txt", ".pptx", ".json"}:
            raise ValueError("نوع ملف الاستيراد غير مدعوم.")
        return clean

    def save_import_profile(
        self,
        name: str,
        file_type: str,
        mapping: dict[str, Any],
        actor: dict[str, Any],
        *,
        options: dict[str, Any] | None = None,
        is_shared: bool = False,
    ) -> str:
        clean_name = str(name or "").strip()
        clean_type = self._profile_file_type(file_type)
        owner_id = str(actor.get("id") or "").strip()
        if not clean_name or len(clean_name) > 80:
            raise ValueError("اسم ملف التعريف مطلوب وبحد أقصى 80 حرفاً.")
        if not owner_id:
            raise PermissionError("يلزم تسجيل الدخول لحفظ ملف التعريف.")
        if is_shared and str(actor.get("role")) != "admin":
            raise PermissionError("مشاركة ملفات الاستيراد متاحة لمدير النظام فقط.")
        if not isinstance(mapping, dict) or not all(
            isinstance(key, str) and isinstance(value, str)
            for key, value in mapping.items()
        ):
            raise ValueError("ربط أعمدة الاستيراد غير صالح.")
        mapping_payload = json.dumps(mapping, ensure_ascii=False, sort_keys=True)
        options_payload = json.dumps(options or {}, ensure_ascii=False, sort_keys=True)
        if len(mapping_payload.encode("utf-8")) + len(options_payload.encode("utf-8")) > 64_000:
            raise ValueError("حجم ملف تعريف الاستيراد أكبر من الحد المسموح.")
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT id,created_at FROM import_profiles WHERE owner_user_id=? AND file_type=? AND name=?",
                (owner_id, clean_type, clean_name),
            ).fetchone()
            profile_id = str(existing["id"]) if existing else uuid.uuid4().hex
            connection.execute(
                """INSERT INTO import_profiles(
                       id,owner_user_id,name,file_type,mapping_json,options_json,
                       is_shared,created_at,updated_at
                   ) VALUES(?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(id) DO UPDATE SET
                       name=excluded.name,file_type=excluded.file_type,
                       mapping_json=excluded.mapping_json,options_json=excluded.options_json,
                       is_shared=excluded.is_shared,updated_at=excluded.updated_at""",
                (
                    profile_id, owner_id, clean_name, clean_type,
                    mapping_payload, options_payload, int(is_shared),
                    str(existing["created_at"]) if existing else now, now,
                ),
            )
            self._write_audit(
                connection,
                str(actor.get("username") or actor.get("display_name") or ""),
                "update" if existing else "create",
                "import_profile",
                profile_id,
                {"name": clean_name, "file_type": clean_type, "shared": bool(is_shared)},
            )
        return profile_id

    def list_import_profiles(
        self, file_type: str, actor: dict[str, Any]
    ) -> list[dict[str, Any]]:
        clean_type = self._profile_file_type(file_type)
        owner_id = str(actor.get("id") or "").strip()
        if not owner_id:
            return []
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT ip.*,u.display_name AS owner_name
                   FROM import_profiles ip JOIN users u ON u.id=ip.owner_user_id
                   WHERE ip.file_type=? AND (ip.owner_user_id=? OR ip.is_shared=1)
                   ORDER BY ip.is_shared ASC,ip.name COLLATE NOCASE""",
                (clean_type, owner_id),
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            for source_key, target_key in (("mapping_json", "mapping"), ("options_json", "options")):
                try:
                    parsed = json.loads(str(item.pop(source_key) or "{}"))
                except (TypeError, ValueError):
                    parsed = {}
                item[target_key] = parsed if isinstance(parsed, dict) else {}
            result.append(item)
        return result

    def delete_import_profile(self, profile_id: str, actor: dict[str, Any]) -> None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id,owner_user_id,name,file_type FROM import_profiles WHERE id=?",
                (str(profile_id),),
            ).fetchone()
            if row is None:
                raise ValueError("ملف تعريف الاستيراد غير موجود.")
            is_owner = str(row["owner_user_id"]) == str(actor.get("id") or "")
            if not is_owner and str(actor.get("role")) != "admin":
                raise PermissionError("لا تملك صلاحية حذف ملف التعريف هذا.")
            connection.execute("DELETE FROM import_profiles WHERE id=?", (str(profile_id),))
            self._write_audit(
                connection,
                str(actor.get("username") or actor.get("display_name") or ""),
                "delete",
                "import_profile",
                str(profile_id),
                {"name": row["name"], "file_type": row["file_type"]},
            )

    def analytics_summary(
        self,
        start_date: str,
        end_date: str,
        company_id: str = "",
    ) -> dict[str, Any]:
        filters = ["r.report_date >= ?", "r.report_date <= ?"]
        parameters: list[Any] = [start_date, end_date]
        if company_id:
            filters.append("r.company_id = ?")
            parameters.append(company_id)
        where = " AND ".join(filters)
        with self._connect() as connection:
            kpi = connection.execute(
                f"""
                SELECT COUNT(DISTINCT r.id) AS reports,
                       COUNT(rr.id) AS records,
                       COALESCE(SUM(rr.total), 0) AS work_total,
                       MAX(r.updated_at) AS freshness
                FROM reports r
                LEFT JOIN report_rows rr ON rr.report_id = r.id
                WHERE {where}
                """,
                parameters,
            ).fetchone()
            daily = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT r.report_date AS label,
                           COUNT(DISTINCT r.id) AS reports,
                           COUNT(rr.id) AS records,
                           COALESCE(SUM(rr.total), 0) AS total
                    FROM reports r
                    LEFT JOIN report_rows rr ON rr.report_id = r.id
                    WHERE {where}
                    GROUP BY r.report_date ORDER BY r.report_date
                    """,
                    parameters,
                ).fetchall()
            ]
            categories = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT rr.category_label AS label,
                           COUNT(rr.id) AS records,
                           COALESCE(SUM(rr.total), 0) AS total
                    FROM reports r
                    JOIN report_rows rr ON rr.report_id = r.id
                    WHERE {where}
                    GROUP BY rr.category_id, rr.category_label
                    ORDER BY total DESC, records DESC
                    """,
                    parameters,
                ).fetchall()
            ]
            people = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT rr.person_name AS label,
                           COUNT(rr.id) AS records,
                           COALESCE(SUM(rr.total), 0) AS total
                    FROM reports r
                    JOIN report_rows rr ON rr.report_id = r.id
                    WHERE {where} AND rr.person_name <> ''
                    GROUP BY rr.person_name
                    ORDER BY total DESC, records DESC
                    LIMIT 20
                    """,
                    parameters,
                ).fetchall()
            ]
            statuses = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT r.status AS label, COUNT(DISTINCT r.id) AS reports
                    FROM reports r
                    WHERE {where}
                    GROUP BY r.status ORDER BY reports DESC
                    """,
                    parameters,
                ).fetchall()
            ]
            quality_rows = connection.execute(
                f"SELECT r.snapshot_json FROM reports r WHERE {where}",
                parameters,
            ).fetchall()
        reports = int(kpi["reports"] or 0) if kpi else 0
        records = int(kpi["records"] or 0) if kpi else 0
        work_total = int(kpi["work_total"] or 0) if kpi else 0
        result = {
            "start_date": start_date,
            "end_date": end_date,
            "company_id": company_id,
            "reports": reports,
            "records": records,
            "work_total": work_total,
            "average_per_report": round(work_total / reports, 2) if reports else 0,
            "freshness": str(kpi["freshness"] or "") if kpi else "",
            "daily": daily,
            "categories": categories,
            "people": people,
            "statuses": statuses,
        }
        quality_scores: list[int] = []
        for row in quality_rows:
            try:
                snapshot = json.loads(str(row["snapshot_json"] or "{}"))
            except (TypeError, ValueError):
                continue
            quality_scores.append(int(report_quality_score(snapshot)["score"]))
        result["quality"] = {
            "average": round(sum(quality_scores) / len(quality_scores), 1) if quality_scores else 0,
            "minimum": min(quality_scores) if quality_scores else 0,
            "maximum": max(quality_scores) if quality_scores else 0,
            "reports": len(quality_scores),
            "needs_attention": sum(score < 70 for score in quality_scores),
            "model": "report-quality-v2",
        }
        result["anomalies"] = robust_anomalies(daily)
        result["forecast"] = neural_forecast(row.get("total", 0) for row in daily)
        result["scenario"] = scenario_forecast(
            (row.get("total", 0) for row in daily), 0
        )
        result["insights"] = productivity_insights(result)
        return result

    def backup_to(self, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as source:
            with sqlite3.connect(destination, factory=AutoClosingConnection) as target:
                source.backup(target)
        return destination

    def restore_from(self, source_path: Path) -> None:
        with sqlite3.connect(source_path, factory=AutoClosingConnection) as source:
            integrity = source.execute("PRAGMA integrity_check").fetchone()
            if not integrity or str(integrity[0]).lower() != "ok":
                raise ValueError("ملف قاعدة البيانات غير سليم.")
            with self._connect() as target:
                source.backup(target)
        self._initialize()

    def health_check(self) -> dict[str, Any]:
        """Return a read-only summary of database integrity and size."""
        with self._connect() as connection:
            integrity_row = connection.execute("PRAGMA integrity_check").fetchone()
            integrity = str(integrity_row[0] if integrity_row else "unknown")
            foreign_key_issues = [
                tuple(row)
                for row in connection.execute("PRAGMA foreign_key_check").fetchall()
            ]
            journal_row = connection.execute("PRAGMA journal_mode").fetchone()
            report_count = int(
                connection.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
            )
            row_count = int(
                connection.execute("SELECT COUNT(*) FROM report_rows").fetchone()[0]
            )
        return {
            "ok": integrity.lower() == "ok" and not foreign_key_issues,
            "integrity_check": integrity,
            "foreign_key_issues": foreign_key_issues,
            "journal_mode": str(journal_row[0] if journal_row else ""),
            "database_size_bytes": self.path.stat().st_size if self.path.exists() else 0,
            "reports": report_count,
            "rows": row_count,
        }

    def optimize(self, actor: str) -> dict[str, Any]:
        """Optimize a healthy SQLite database and audit the operation."""
        before = self.health_check()
        if not before["ok"]:
            raise ValueError(
                "تعذر بدء الصيانة لأن فحص سلامة قاعدة البيانات لم ينجح."
            )
        connection = sqlite3.connect(self.path, timeout=60, isolation_level=None)
        try:
            connection.execute("PRAGMA busy_timeout = 60000")
            connection.execute("PRAGMA wal_checkpoint(TRUNCATE)").fetchone()
            connection.execute("PRAGMA optimize")
            connection.execute("VACUUM")
        finally:
            connection.close()
        after = self.health_check()
        if not after["ok"]:
            raise ValueError("لم ينجح فحص السلامة بعد صيانة قاعدة البيانات.")
        result = {
            "ok": True,
            "before_size_bytes": int(before["database_size_bytes"]),
            "after_size_bytes": int(after["database_size_bytes"]),
            "reclaimed_bytes": max(
                0,
                int(before["database_size_bytes"])
                - int(after["database_size_bytes"]),
            ),
            "reports": int(after["reports"]),
            "rows": int(after["rows"]),
        }
        self.audit(actor, "database_maintenance", "database", details=result)
        return result

    def migrate_legacy_history(self, records: Any) -> int:
        if not isinstance(records, list):
            return 0
        migrated = 0
        for record in records:
            if not isinstance(record, dict):
                continue
            snapshot = record.get("report")
            if not isinstance(snapshot, dict) or not snapshot.get("rows"):
                continue
            report_id = str(record.get("id") or uuid.uuid4().hex)
            with self._connect() as connection:
                exists = connection.execute(
                    "SELECT 1 FROM reports WHERE id = ?",
                    (report_id,),
                ).fetchone()
            if exists:
                continue
            self.save_snapshot(
                snapshot,
                report_id,
                saved_at=str(record.get("saved_at") or ""),
                actor="ترحيل تلقائي",
            )
            migrated += 1
        return migrated


class BackupManager:
    def __init__(
        self,
        data_dir: Path,
        database: ReportDatabase,
        config_path: Path,
    ):
        self.data_dir = data_dir
        self.database = database
        self.config_path = config_path
        self.backup_dir = data_dir / "backups"
        self.auto_dir = self.backup_dir / "auto"
        self.assets_dir = data_dir / "company_assets"
        self.auto_dir.mkdir(parents=True, exist_ok=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)

    def automatic_backup(self, force: bool = False) -> Path | None:
        today = date.today().isoformat()
        marker = self.auto_dir / "last_backup.txt"
        previous = ""
        try:
            if marker.exists():
                previous = marker.read_text(encoding="utf-8").strip()
        except OSError:
            previous = ""
        if not force and previous == today:
            return None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destination = self.auto_dir / f"reports_auto_{stamp}.db"
        self.database.backup_to(destination)
        marker.write_text(today, encoding="utf-8")
        backups = sorted(
            self.auto_dir.glob("reports_auto_*.db"),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        )
        for old_backup in backups[30:]:
            try:
                old_backup.unlink()
            except OSError:
                pass
        return destination

    def create_archive(self, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temporary_db = self.backup_dir / f"manual_{stamp}.db"
        self.database.backup_to(temporary_db)
        try:
            with zipfile.ZipFile(
                destination,
                "w",
                compression=zipfile.ZIP_DEFLATED,
                compresslevel=6,
            ) as archive:
                archive.write(temporary_db, "reports.db")
                if self.config_path.is_file():
                    archive.write(self.config_path, "config.json")
                if self.assets_dir.is_dir():
                    for asset in self.assets_dir.rglob("*"):
                        if asset.is_file():
                            archive.write(
                                asset,
                                str(Path("company_assets") / asset.relative_to(self.assets_dir)),
                            )
                archive.writestr(
                    "backup_info.json",
                    json.dumps(
                        {
                            "application": APP_NAME,
                            "version": APP_VERSION,
                            "created_at": datetime.now().isoformat(timespec="seconds"),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                )
        finally:
            try:
                temporary_db.unlink()
            except OSError:
                pass
        return destination

    def restore_archive(self, source: Path) -> dict[str, bool]:
        self.automatic_backup(force=True)
        restored = {"database": False, "config": False, "assets": False}
        if source.suffix.lower() in {".db", ".sqlite", ".sqlite3"}:
            self.database.restore_from(source)
            restored["database"] = True
            return restored
        if source.suffix.lower() != ".zip":
            raise ValueError("اختر ملف نسخ احتياطي ZIP أو قاعدة بيانات DB.")
        restore_dir = self.data_dir / f"_restore_{uuid.uuid4().hex}"
        restore_dir.mkdir(parents=True, exist_ok=False)
        try:
            with zipfile.ZipFile(source, "r") as archive:
                for member in archive.infolist():
                    target = (restore_dir / member.filename).resolve()
                    if restore_dir.resolve() not in target.parents and target != restore_dir.resolve():
                        raise ValueError("يحتوي ملف النسخة على مسار غير آمن.")
                archive.extractall(restore_dir)
            restored_db = restore_dir / "reports.db"
            if not restored_db.is_file():
                raise ValueError("لا تحتوي النسخة على قاعدة البيانات.")
            self.database.restore_from(restored_db)
            restored["database"] = True
            restored_config = restore_dir / "config.json"
            if restored_config.is_file():
                restored_config.replace(self.config_path)
                restored["config"] = True
            restored_assets = restore_dir / "company_assets"
            if restored_assets.is_dir():
                self.assets_dir.mkdir(parents=True, exist_ok=True)
                for asset in restored_assets.rglob("*"):
                    if asset.is_file():
                        destination = self.assets_dir / asset.relative_to(restored_assets)
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(asset, destination)
                restored["assets"] = True
        finally:
            shutil.rmtree(restore_dir, ignore_errors=True)
        return restored

    def maintain(self, actor: str) -> dict[str, Any]:
        """Create a safety backup before database optimization."""
        backup = self.automatic_backup(force=True)
        result = self.database.optimize(actor)
        return {**result, "backup": str(backup or "")}


class DraftStore:
    """Atomic per-user recovery storage for the currently edited report."""

    def __init__(self, path: Path):
        self.path = path
        self.history_dir = path.parent / f"{path.stem}_history"

    def _valid_payload(self, payload: Any) -> dict[str, Any] | None:
        if not isinstance(payload, dict) or payload.get("format") != 1:
            return None
        snapshot = payload.get("snapshot")
        if not isinstance(snapshot, dict) or not isinstance(snapshot.get("rows"), list):
            return None
        return payload if snapshot["rows"] else None

    def save(
        self,
        snapshot: dict[str, Any],
        *,
        report_id: str = "",
        is_locked: bool = False,
        raw_entries: list[dict[str, Any]] | None = None,
    ) -> None:
        rows = snapshot.get("rows", [])
        if not isinstance(rows, list) or not rows:
            self.clear()
            return
        payload = {
            "format": 1,
            "app_version": APP_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "report_id": str(report_id),
            "is_locked": bool(is_locked),
            "snapshot": snapshot,
            # Backward-compatible extension: preserve values for fields that
            # are no longer visible in the current report schema.
            "raw_entries": copy.deepcopy(raw_entries) if raw_entries is not None else None,
        }
        atomic_write_json(self.path, payload)
        checkpoint = self.history_dir / (
            f"draft_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_"
            f"{uuid.uuid4().hex[:8]}.json"
        )
        atomic_write_json(checkpoint, payload)
        checkpoints = sorted(
            self.history_dir.glob("draft_*.json"),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        )
        for expired in checkpoints[20:]:
            try:
                expired.unlink(missing_ok=True)
            except OSError:
                pass

    def load(self) -> dict[str, Any] | None:
        candidates = [self.path]
        if self.history_dir.is_dir():
            candidates.extend(
                sorted(
                    self.history_dir.glob("draft_*.json"),
                    key=lambda item: item.stat().st_mtime,
                    reverse=True,
                )
            )
        for candidate in candidates:
            payload = self._valid_payload(read_json(candidate, None))
            if payload is not None:
                return payload
        return None

    def clear(self) -> None:
        try:
            self.path.unlink(missing_ok=True)
        except OSError:
            pass
        if self.history_dir.is_dir():
            for checkpoint in self.history_dir.glob("draft_*.json"):
                try:
                    checkpoint.unlink(missing_ok=True)
                except OSError:
                    pass


def coerce_nonnegative_int(value: Any) -> int:
    if isinstance(value, bool):
        return int(value)
    try:
        parsed = int(str(value).strip() or "0")
    except (TypeError, ValueError):
        return 0
    return max(0, parsed)


def calculate_total(config: dict[str, Any], entry: dict[str, Any]) -> int:
    category = config["categories"].get(entry.get("category"), {})
    values = entry.get("values", {})
    if not isinstance(values, dict):
        values = {}
    return sum(
        coerce_nonnegative_int(values.get(field["id"], 0))
        for field in category.get("fields", [])
        if field.get("kind") == "number"
    )


def arabic_day_name(value: date) -> str:
    return ARABIC_WEEKDAYS[value.weekday()]


def new_entry(category: str, name: str, values: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": uuid.uuid4().hex,
        "category": category,
        "name": name.strip(),
        "values": dict(values),
    }


def _import_header_key(value: Any) -> str:
    return re.sub(r"[\s_\-/\\]+", "", str(value or "").strip().casefold())


def _records_to_import_matrix(records: list[dict[str, Any]]) -> list[list[Any]]:
    headers: list[str] = []
    for record in records:
        for key in record:
            clean_key = str(key or "").strip()
            if clean_key and clean_key not in headers:
                headers.append(clean_key)
    if not headers:
        return []
    return [headers, *[[record.get(header, "") for header in headers] for record in records]]


def _key_value_fragments(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for fragment in re.split(r"[|؛\n]", text):
        separator = ":" if ":" in fragment else "=" if "=" in fragment else ""
        if not separator:
            continue
        label, value = fragment.split(separator, 1)
        label = label.strip()
        value = value.strip()
        if label and value:
            result[label] = value
    return result


def _normalized_profile_mapping(profile: dict[str, Any] | None) -> dict[str, str]:
    if not isinstance(profile, dict):
        return {}
    mapping = profile.get("mapping", profile.get("header_map", {}))
    if not isinstance(mapping, dict):
        return {}
    return {
        _import_header_key(source): str(target).strip()
        for source, target in mapping.items()
        if str(source).strip() and str(target).strip()
    }


def _read_import_table(
    path: Path,
    profile: dict[str, Any] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    if not path.is_file():
        raise ValueError("ملف الاستيراد غير موجود.")
    if path.stat().st_size > 25 * 1024 * 1024:
        raise ValueError("حجم ملف الاستيراد أكبر من الحد المسموح 25 MB.")
    suffix = path.suffix.lower()
    profile_mapping = _normalized_profile_mapping(profile)

    def mapped_key(value: Any) -> str:
        source_key = _import_header_key(value)
        return _import_header_key(profile_mapping.get(source_key, value))
    if suffix in {".csv", ".tsv", ".txt"}:
        content = ""
        for encoding in ("utf-8-sig", "cp1256"):
            try:
                content = path.read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if not content:
            raise ValueError("تعذر قراءة ترميز ملف CSV.")
        try:
            dialect = csv.Sniffer().sniff(content[:8192], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel_tab if suffix == ".tsv" else csv.excel
        matrix = list(csv.reader(content.splitlines(), dialect))
    elif suffix == ".xlsx":
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("استيراد Excel غير متاح في هذا التثبيت.")
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook["السجلات"] if "السجلات" in workbook.sheetnames else workbook.active
            matrix = [
                list(row)
                for index, row in enumerate(sheet.iter_rows(values_only=True), 1)
                if index <= 5001
            ]
        finally:
            workbook.close()
    elif suffix == ".pptx":
        if not PPTX_AVAILABLE:
            raise RuntimeError("استيراد PowerPoint يحتاج مكتبة python-pptx.")
        presentation = Presentation(str(path))
        records: list[dict[str, Any]] = []
        for slide in presentation.slides:
            slide_record: dict[str, Any] = {}
            slide_has_complete_line = False
            for shape in slide.shapes:
                if getattr(shape, "has_table", False):
                    table_matrix = [
                        [cell.text.strip() for cell in row.cells]
                        for row in shape.table.rows
                    ]
                    if table_matrix:
                        table_headers = table_matrix[0]
                        for table_row in table_matrix[1:]:
                            record = {
                                str(table_headers[index]).strip(): value
                                for index, value in enumerate(table_row)
                                if index < len(table_headers)
                                and str(table_headers[index]).strip()
                            }
                            if any(str(value or "").strip() for value in record.values()):
                                records.append(record)
                shape_text = str(getattr(shape, "text", "") or "").strip()
                if shape_text:
                    parsed = _key_value_fragments(shape_text)
                    if parsed:
                        slide_record.update(parsed)
                    for line in shape_text.splitlines():
                        line_record = _key_value_fragments(line)
                        keys = {mapped_key(key) for key in line_record}
                        if keys & {"الفئة", "القسم", "category"} and keys & {
                            "الاسم",
                            "الاسمالجهة",
                            "name",
                            "person",
                        }:
                            records.append(line_record)
                            slide_has_complete_line = True
            slide_keys = {mapped_key(key) for key in slide_record}
            if not slide_has_complete_line and slide_keys & {"الفئة", "القسم", "category"} and slide_keys & {
                "الاسم",
                "الاسمالجهة",
                "name",
                "person",
            }:
                records.append(slide_record)
        matrix = _records_to_import_matrix(records)
    elif suffix == ".json":
        try:
            loaded = json.loads(path.read_text(encoding="utf-8-sig"))
        except (UnicodeError, json.JSONDecodeError) as error:
            raise ValueError("ملف JSON غير صالح أو ترميزه غير مدعوم.") from error
        source_records: Any = loaded
        if isinstance(loaded, dict):
            source_records = next(
                (
                    loaded.get(key)
                    for key in ("rows", "entries", "records", "data")
                    if isinstance(loaded.get(key), list)
                ),
                [],
            )
        records = []
        for source_record in source_records if isinstance(source_records, list) else []:
            if not isinstance(source_record, dict):
                continue
            record = dict(source_record)
            if isinstance(source_record.get("fields"), list):
                record.setdefault(
                    "الفئة",
                    source_record.get("category_label") or source_record.get("category", ""),
                )
                record.setdefault("الاسم/الجهة", source_record.get("name", ""))
                for field in source_record["fields"]:
                    if isinstance(field, dict):
                        label = str(field.get("label") or field.get("id") or "").strip()
                        if label:
                            record[label] = field.get("value", "")
            records.append(record)
        matrix = _records_to_import_matrix(records)
    else:
        raise ValueError(
            "الصيغ المدعومة: Excel XLSX وCSV/TSV/TXT وPowerPoint PPTX وJSON."
        )
    if not matrix:
        raise ValueError("ملف الاستيراد فارغ.")
    header_index = next(
        (
            index
            for index, row in enumerate(matrix[:20])
            if any(
                mapped_key(cell)
                in {"الفئة", "القسم", "category", "الاسم", "الاسمالجهة", "name", "person"}
                for cell in row
            )
        ),
        None,
    )
    if header_index is None:
        header_index = next(
            (
                index
                for index, row in enumerate(matrix[:20])
                if sum(1 for cell in row if str(cell or "").strip()) >= 2
            ),
            None,
        )
    if header_index is None:
        raise ValueError("لم يُعثر على صف عناوين يحتوي الفئة والاسم/الجهة.")
    headers = [
        profile_mapping.get(_import_header_key(value), str(value or "").strip())
        for value in matrix[header_index]
    ]
    return headers, matrix[header_index + 1 :]


def import_entries_from_path(
    path: Path | str,
    config: dict[str, Any],
    profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Import report entries by current category and field labels."""
    headers, rows = _read_import_table(Path(path), profile)
    options = profile.get("options", {}) if isinstance(profile, dict) else {}
    if not isinstance(options, dict):
        options = {}
    categories: dict[str, str] = {}
    for category_id in category_ids(config):
        category = config["categories"][category_id]
        categories[_import_header_key(category_id)] = category_id
        categories[_import_header_key(category.get("label", ""))] = category_id
    category_aliases = ("الفئة", "القسم", "category")
    name_aliases = (
        "الاسمالجهة",
        "الاسم",
        "الجهة",
        "الشخص",
        "name",
        "person",
        *(
            _import_header_key(config["categories"][category_id].get("person_label", ""))
            for category_id in category_ids(config)
        ),
    )
    details_aliases = ("تفاصيلالعمل", "التفاصيل", "details")
    imported: list[dict[str, Any]] = []
    warnings: list[str] = []
    skipped = 0
    for source_index, source_row in enumerate(rows, 2):
        values_by_header = {
            _import_header_key(headers[index]): value
            for index, value in enumerate(source_row)
            if index < len(headers) and headers[index]
        }
        if not any(str(value or "").strip() for value in source_row):
            continue
        category_value = next(
            (
                values_by_header.get(_import_header_key(alias))
                for alias in category_aliases
                if values_by_header.get(_import_header_key(alias)) not in (None, "")
            ),
            "",
        )
        if not str(category_value).strip():
            category_value = str(options.get("default_category") or "")
        category_id = categories.get(_import_header_key(category_value), "")
        if not category_id and not str(category_value).strip():
            matching_categories = []
            for candidate_id in category_ids(config):
                candidate_fields = config["categories"][candidate_id].get("fields", [])
                if any(
                    _import_header_key(field.get("id", "")) in values_by_header
                    or _import_header_key(field.get("label", "")) in values_by_header
                    for field in candidate_fields
                ):
                    matching_categories.append(candidate_id)
            if len(matching_categories) == 1:
                category_id = matching_categories[0]
            elif len(category_ids(config)) == 1:
                category_id = category_ids(config)[0]
        name = str(
            next(
                (
                    values_by_header.get(_import_header_key(alias))
                    for alias in name_aliases
                    if values_by_header.get(_import_header_key(alias)) not in (None, "")
                ),
                "",
            )
            or ""
        ).strip()
        if not name:
            name = str(options.get("default_name") or "").strip()
        if not category_id or not name:
            skipped += 1
            warnings.append(f"الصف {source_index}: الفئة أو الاسم/الجهة غير معروف.")
            continue
        details_text = str(
            next(
                (
                    values_by_header.get(_import_header_key(alias))
                    for alias in details_aliases
                    if values_by_header.get(_import_header_key(alias)) not in (None, "")
                ),
                "",
            )
            or ""
        )
        details: dict[str, str] = {}
        for fragment in details_text.split("|"):
            if ":" in fragment:
                label, value = fragment.split(":", 1)
                details[_import_header_key(label)] = value.strip()
        entry_values: dict[str, Any] = {}
        invalid_number = ""
        for field in config["categories"][category_id].get("fields", []):
            candidates = (
                _import_header_key(field.get("id", "")),
                _import_header_key(field.get("label", "")),
            )
            raw_value: Any = ""
            for candidate in candidates:
                if candidate in values_by_header and values_by_header[candidate] not in (None, ""):
                    raw_value = values_by_header[candidate]
                    break
                if candidate in details:
                    raw_value = details[candidate]
                    break
            if field.get("kind") == "number":
                text_value = str(raw_value or "0").strip()
                if re.fullmatch(r"\d+(?:\.0+)?", text_value):
                    entry_values[field["id"]] = int(float(text_value))
                else:
                    invalid_number = str(field.get("label", field["id"]))
                    break
            else:
                entry_values[field["id"]] = str(raw_value or "").strip()
        if invalid_number:
            skipped += 1
            warnings.append(
                f"الصف {source_index}: قيمة الحقل «{invalid_number}» ليست رقمًا صحيحًا موجبًا."
            )
            continue
        imported.append(new_entry(category_id, name, entry_values))
    return {
        "entries": imported,
        "imported": len(imported),
        "skipped": skipped,
        "warnings": warnings[:50],
        "profile_name": str(profile.get("name") or "") if isinstance(profile, dict) else "",
    }


def analyze_import_path(
    path: Path | str,
    config: dict[str, Any],
    profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    source_headers, source_rows = _read_import_table(Path(path))
    mapping = _normalized_profile_mapping(profile)
    mapped_headers = [
        mapping.get(_import_header_key(header), header)
        for header in source_headers
    ]
    result = import_entries_from_path(path, config, profile)
    preview_rows = [
        {
            mapped_headers[index]: row[index] if index < len(row) else ""
            for index in range(len(mapped_headers))
            if mapped_headers[index]
        }
        for row in source_rows[:5]
    ]
    return {
        "source_headers": source_headers,
        "mapped_headers": mapped_headers,
        "preview_rows": preview_rows,
        "imported": int(result["imported"]),
        "skipped": int(result["skipped"]),
        "warnings": list(result["warnings"]),
        "profile_name": str(result.get("profile_name") or ""),
    }


def export_import_template_to_path(
    config: dict[str, Any],
    output_path: Path | str,
) -> Path:
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("إنشاء قالب Excel غير متاح في هذا التثبيت.")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "الاستيراد"
    field_labels: list[str] = []
    for category_id in category_ids(config):
        for field in config["categories"][category_id].get("fields", []):
            label = str(field.get("label", "")).strip()
            if label and label not in field_labels:
                field_labels.append(label)
    sheet.append(["الفئة", "الاسم/الجهة", *field_labels])
    for category_id in category_ids(config):
        category = config["categories"][category_id]
        sheet.append(
            [
                category.get("label", category_id),
                (category.get("people") or [""])[0],
                *["" for _field in field_labels],
            ]
        )
    _excel_style_table(sheet, "ImportTemplate")
    workbook.save(output)
    return output


def build_report_snapshot(
    config: dict[str, Any],
    entries: list[dict[str, Any]],
    report_date: date,
    code: str,
    page_label: str,
    company_name: str | None = None,
    user_name: str | None = None,
    *,
    company_profile: dict[str, Any] | None = None,
    company_id: str = "",
    template_id: str = "",
    template_name: str = "",
    status: str = "draft",
) -> dict[str, Any]:
    config_copy = normalize_config(copy.deepcopy(config))
    identity = config_copy["identity"]
    effective_company = (
        str(company_name).strip()
        if company_name is not None
        else identity["company_name"]
    )
    effective_user = (
        str(user_name).strip() if user_name is not None else identity["user_name"]
    )
    rows: list[dict[str, Any]] = []
    for entry in entries:
        category_id = entry.get("category")
        category = config_copy["categories"].get(category_id)
        if not category:
            continue
        name = str(entry.get("name", "")).strip()
        if not name:
            continue
        raw_values = entry.get("values", {})
        if not isinstance(raw_values, dict):
            raw_values = {}
        fields: list[dict[str, Any]] = []
        clean_values: dict[str, Any] = {}
        for field in category["fields"]:
            field_id = field["id"]
            kind = field.get("kind", "text")
            if kind == "number":
                value: Any = coerce_nonnegative_int(raw_values.get(field_id, 0))
            elif kind in {"percentage", "currency"}:
                try:
                    value = round(float(raw_values.get(field_id, 0) or 0), 2)
                except (TypeError, ValueError):
                    value = 0.0
            else:
                value = str(raw_values.get(field_id, "")).strip()
            clean_values[field_id] = value
            if field.get("report_visible", True):
                report_field = {**field, "value": value}
                report_field["label"] = str(field.get("report_label") or field.get("label") or field_id)
                fields.append(report_field)
        clean_entry = {
            "category": category_id,
            "name": name,
            "values": clean_values,
        }
        rows.append(
            {
                "category": category_id,
                "category_label": category["label"],
                "name": name,
                "fields": fields,
                "total": calculate_total(config_copy, clean_entry),
            }
        )
    return {
        "version": CONFIG_VERSION,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "report_title": identity["report_title"],
        "report_theme": config_copy.get("report_theme", "executive_teal"),
        "report_orientation": config_copy.get("report_layout", {}).get("orientation", "landscape"),
        "report_footer": identity.get("report_footer", ""),
        "template_id": template_id,
        "template_name": template_name,
        "company_id": company_id,
        "company_name": effective_company,
        "company_profile": copy.deepcopy(company_profile or {}),
        "company_logo_path": str((company_profile or {}).get("logo_path", "")),
        "user_name": effective_user or default_user_name(),
        "status": status if status in REPORT_STATUSES else "draft",
        "date": report_date.isoformat(),
        "day": arabic_day_name(report_date),
        "code": code.strip(),
        "page": page_label.strip() or identity["default_report_number"],
        "config": config_copy,
        "rows": rows,
    }


def parse_iso_date(value: Any) -> date | None:
    try:
        return date.fromisoformat(str(value).strip())
    except (TypeError, ValueError):
        return None


def week_range(reference: date) -> tuple[date, date]:
    start = reference.fromordinal(reference.toordinal() - reference.weekday())
    return start, start.fromordinal(start.toordinal() + 6)


def month_range(reference: date) -> tuple[date, date]:
    start = reference.replace(day=1)
    end = reference.replace(day=calendar.monthrange(reference.year, reference.month)[1])
    return start, end


def entries_from_snapshot(snapshot: dict[str, Any]) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for row in snapshot.get("rows", []):
        if not isinstance(row, dict):
            continue
        category_id = str(row.get("category", "")).strip()
        name = str(row.get("name", "")).strip()
        if not category_id or not name:
            continue
        values = {
            str(field.get("id")): field.get("value", "")
            for field in row.get("fields", [])
            if isinstance(field, dict) and field.get("id")
        }
        entries.append(new_entry(category_id, name, values))
    return entries


def combine_report_snapshots(
    snapshots: list[dict[str, Any]],
    fallback_config: dict[str, Any],
    period_label: str,
) -> dict[str, Any]:
    valid_snapshots = [
        snapshot
        for snapshot in snapshots
        if isinstance(snapshot, dict) and snapshot.get("rows")
    ]
    if not valid_snapshots:
        raise ValueError("لا توجد بيانات محفوظة ضمن الفترة المحددة.")
    valid_snapshots.sort(key=lambda item: str(item.get("date", "")))

    fallback = normalize_config(copy.deepcopy(fallback_config))
    combined_config: dict[str, Any] = {
        "version": CONFIG_VERSION,
        "report_theme": fallback.get("report_theme", "executive_teal"),
        "app_state": copy.deepcopy(fallback.get("app_state", {})),
        "identity": copy.deepcopy(fallback["identity"]),
        "category_order": [],
        "categories": {},
    }
    combined_rows: list[dict[str, Any]] = []
    date_field_id = "__work_date"

    for snapshot in valid_snapshots:
        source_config = normalize_config(snapshot.get("config", fallback))
        report_date = str(snapshot.get("date", ""))
        for source_row in snapshot.get("rows", []):
            if not isinstance(source_row, dict):
                continue
            category_id = str(source_row.get("category", "")).strip()
            if not category_id:
                continue
            source_category = source_config["categories"].get(category_id, {})
            if category_id not in combined_config["categories"]:
                combined_config["category_order"].append(category_id)
                combined_config["categories"][category_id] = {
                    "label": str(
                        source_row.get("category_label")
                        or source_category.get("label")
                        or "قسم محفوظ"
                    ),
                    "person_label": str(
                        source_category.get("person_label", "الاسم")
                    ),
                    "total_label": str(
                        source_category.get("total_label", "الإجمالي")
                    ),
                    "people": [],
                    "fields": [
                        {
                            "id": date_field_id,
                            "label": "تاريخ العمل",
                            "kind": "text",
                            "options": [],
                        }
                    ],
                }
            combined_category = combined_config["categories"][category_id]
            existing_field_ids = {
                field["id"] for field in combined_category["fields"]
            }
            row_fields: list[dict[str, Any]] = [
                {
                    "id": date_field_id,
                    "label": "تاريخ العمل",
                    "kind": "text",
                    "options": [],
                    "value": report_date,
                }
            ]
            for raw_field in source_row.get("fields", []):
                if not isinstance(raw_field, dict):
                    continue
                label = str(raw_field.get("label", "")).strip() or "حقل محفوظ"
                field_id = str(raw_field.get("id", "")).strip()
                if not field_id:
                    field_id = (
                        "field_"
                        + uuid.uuid5(uuid.NAMESPACE_DNS, label).hex[:12]
                    )
                kind = str(raw_field.get("kind", "number")).strip().lower()
                if kind not in SUPPORTED_FIELD_KINDS:
                    kind = "text"
                options = unique_clean_strings(raw_field.get("options", []))
                if field_id not in existing_field_ids:
                    combined_category["fields"].append(
                        {
                            "id": field_id,
                            "label": label,
                            "kind": kind,
                            "options": options,
                        }
                    )
                    existing_field_ids.add(field_id)
                row_fields.append(
                    {
                        "id": field_id,
                        "label": label,
                        "kind": kind,
                        "options": options,
                        "value": raw_field.get("value", ""),
                    }
                )
            combined_rows.append(
                {
                    "category": category_id,
                    "category_label": combined_category["label"],
                    "name": str(source_row.get("name", "")),
                    "fields": row_fields,
                    "total": coerce_nonnegative_int(source_row.get("total", 0)),
                }
            )

    combined_config = normalize_config(combined_config)
    for row in combined_rows:
        category = combined_config["categories"].get(row["category"], {})
        row["category_label"] = category.get("label", row["category_label"])

    first_date = str(valid_snapshots[0].get("date", ""))
    last_date = str(valid_snapshots[-1].get("date", ""))
    latest = valid_snapshots[-1]
    identity = combined_config["identity"]
    company = identity["company_name"] or str(latest.get("company_name", ""))
    user = identity["user_name"] or str(latest.get("user_name", ""))
    return {
        "version": CONFIG_VERSION,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "report_title": f"{identity['report_title']} - {period_label}",
        "report_theme": str(
            latest.get("report_theme")
            or combined_config.get("report_theme", "executive_teal")
        ),
        "company_name": company,
        "company_profile": copy.deepcopy(latest.get("company_profile", {})),
        "company_logo_path": str(latest.get("company_logo_path", "")),
        "user_name": user or default_user_name(),
        "date": first_date if first_date == last_date else f"{first_date} - {last_date}",
        "day": period_label,
        "code": f"{len(valid_snapshots)} تقرير محفوظ",
        "page": "1",
        "config": combined_config,
        "rows": combined_rows,
    }


# Arabic presentation forms are used so Arabic remains joined even when optional
# bidi/reshaping packages are not installed.
ARABIC_FORMS: dict[str, tuple[str, str | None, str | None, str | None]] = {
    "\u0621": ("\uFE80", None, None, None),
    "\u0622": ("\uFE81", "\uFE82", None, None),
    "\u0623": ("\uFE83", "\uFE84", None, None),
    "\u0624": ("\uFE85", "\uFE86", None, None),
    "\u0625": ("\uFE87", "\uFE88", None, None),
    "\u0626": ("\uFE89", "\uFE8A", "\uFE8B", "\uFE8C"),
    "\u0627": ("\uFE8D", "\uFE8E", None, None),
    "\u0628": ("\uFE8F", "\uFE90", "\uFE91", "\uFE92"),
    "\u0629": ("\uFE93", "\uFE94", None, None),
    "\u062A": ("\uFE95", "\uFE96", "\uFE97", "\uFE98"),
    "\u062B": ("\uFE99", "\uFE9A", "\uFE9B", "\uFE9C"),
    "\u062C": ("\uFE9D", "\uFE9E", "\uFE9F", "\uFEA0"),
    "\u062D": ("\uFEA1", "\uFEA2", "\uFEA3", "\uFEA4"),
    "\u062E": ("\uFEA5", "\uFEA6", "\uFEA7", "\uFEA8"),
    "\u062F": ("\uFEA9", "\uFEAA", None, None),
    "\u0630": ("\uFEAB", "\uFEAC", None, None),
    "\u0631": ("\uFEAD", "\uFEAE", None, None),
    "\u0632": ("\uFEAF", "\uFEB0", None, None),
    "\u0633": ("\uFEB1", "\uFEB2", "\uFEB3", "\uFEB4"),
    "\u0634": ("\uFEB5", "\uFEB6", "\uFEB7", "\uFEB8"),
    "\u0635": ("\uFEB9", "\uFEBA", "\uFEBB", "\uFEBC"),
    "\u0636": ("\uFEBD", "\uFEBE", "\uFEBF", "\uFEC0"),
    "\u0637": ("\uFEC1", "\uFEC2", "\uFEC3", "\uFEC4"),
    "\u0638": ("\uFEC5", "\uFEC6", "\uFEC7", "\uFEC8"),
    "\u0639": ("\uFEC9", "\uFECA", "\uFECB", "\uFECC"),
    "\u063A": ("\uFECD", "\uFECE", "\uFECF", "\uFED0"),
    "\u0641": ("\uFED1", "\uFED2", "\uFED3", "\uFED4"),
    "\u0642": ("\uFED5", "\uFED6", "\uFED7", "\uFED8"),
    "\u0643": ("\uFED9", "\uFEDA", "\uFEDB", "\uFEDC"),
    "\u0644": ("\uFEDD", "\uFEDE", "\uFEDF", "\uFEE0"),
    "\u0645": ("\uFEE1", "\uFEE2", "\uFEE3", "\uFEE4"),
    "\u0646": ("\uFEE5", "\uFEE6", "\uFEE7", "\uFEE8"),
    "\u0647": ("\uFEE9", "\uFEEA", "\uFEEB", "\uFEEC"),
    "\u0648": ("\uFEED", "\uFEEE", None, None),
    "\u0649": ("\uFEEF", "\uFEF0", None, None),
    "\u064A": ("\uFEF1", "\uFEF2", "\uFEF3", "\uFEF4"),
}


def _previous_arabic_index(chars: list[str], index: int) -> int | None:
    candidate = index - 1
    while candidate >= 0 and "\u064B" <= chars[candidate] <= "\u065F":
        candidate -= 1
    return candidate if candidate >= 0 and chars[candidate] in ARABIC_FORMS else None


def _next_arabic_index(chars: list[str], index: int) -> int | None:
    candidate = index + 1
    while candidate < len(chars) and "\u064B" <= chars[candidate] <= "\u065F":
        candidate += 1
    return candidate if candidate < len(chars) and chars[candidate] in ARABIC_FORMS else None


def shape_arabic(text: str) -> str:
    chars = list(str(text))
    output: list[str] = []
    for index, char in enumerate(chars):
        forms = ARABIC_FORMS.get(char)
        if not forms:
            output.append(char)
            continue
        previous_index = _previous_arabic_index(chars, index)
        next_index = _next_arabic_index(chars, index)
        joins_previous = bool(
            previous_index is not None
            and forms[1]
            and ARABIC_FORMS[chars[previous_index]][2]
        )
        joins_next = bool(
            next_index is not None
            and forms[2]
            and ARABIC_FORMS[chars[next_index]][1]
        )
        if joins_previous and joins_next and forms[3]:
            output.append(forms[3])
        elif joins_previous and forms[1]:
            output.append(forms[1])
        elif joins_next and forms[2]:
            output.append(forms[2])
        else:
            output.append(forms[0])
    return "".join(output)


def contains_arabic(text: Any) -> bool:
    return bool(re.search(r"[\u0600-\u06FF\uFE70-\uFEFF]", str(text)))


def rtl_visual(text: Any) -> str:
    logical = str(text)
    if not contains_arabic(logical):
        return logical
    visual = shape_arabic(logical)[::-1]
    protected_ranges: dict[str, str] = {}

    def protect_date_range(match: re.Match[str]) -> str:
        marker = chr(0xE000 + len(protected_ranges))
        protected_ranges[marker] = match.group(0)[::-1]
        return marker

    visual = re.sub(
        r"\d{2}-\d{2}-\d{4}\s*-\s*\d{2}-\d{2}-\d{4}",
        protect_date_range,
        visual,
    )
    visual = re.sub(
        r"[A-Za-z](?:[A-Za-z0-9._/@:+\-]| (?=[A-Za-z]))*",
        lambda match: match.group(0)[::-1],
        visual,
    )
    visual = re.sub(
        r"[0-9][0-9._/@:+\-]*",
        lambda match: match.group(0)[::-1],
        visual,
    )
    for marker, value in protected_ranges.items():
        visual = visual.replace(marker, value)
    return visual


def resolve_font_paths() -> tuple[Path, Path]:
    local_regular = os.environ.get("DAILY_WORK_REPORTS_FONT") or os.environ.get(
        "KITCHENS_ERP_FONT"
    )
    local_bold = os.environ.get("DAILY_WORK_REPORTS_BOLD_FONT") or os.environ.get(
        "KITCHENS_ERP_BOLD_FONT"
    )
    regular_candidates = [
        local_regular,
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\tahoma.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
    ]
    bold_candidates = [
        local_bold,
        r"C:\Windows\Fonts\segoeuib.ttf",
        r"C:\Windows\Fonts\tahomabd.ttf",
        r"C:\Windows\Fonts\arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Bold.ttf",
    ]
    regular = next((Path(p) for p in regular_candidates if p and Path(p).exists()), None)
    bold = next((Path(p) for p in bold_candidates if p and Path(p).exists()), None)
    if regular is None:
        raise RuntimeError(
            "لم يتم العثور على خط يدعم العربية. عيّن DAILY_WORK_REPORTS_FONT لمسار خط TTF."
        )
    return regular, bold or regular


_PDF_FONTS_READY = False


def register_pdf_fonts() -> tuple[str, str]:
    global _PDF_FONTS_READY
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("مكتبة reportlab غير مثبتة.")
    regular_path, bold_path = resolve_font_paths()
    if not _PDF_FONTS_READY:
        pdfmetrics.registerFont(TTFont("KitchensArabic", str(regular_path)))
        pdfmetrics.registerFont(TTFont("KitchensArabicBold", str(bold_path)))
        _PDF_FONTS_READY = True
    return "KitchensArabic", "KitchensArabicBold"


def _snapshot_rows(snapshot: dict[str, Any], category_id: str) -> list[dict[str, Any]]:
    return [
        row
        for row in snapshot.get("rows", [])
        if row.get("category") == category_id
    ]


def _field_value_map(row: dict[str, Any]) -> dict[str, Any]:
    return {
        str(field.get("id")): field.get("value", "")
        for field in row.get("fields", [])
        if isinstance(field, dict)
    }


def _logical_table_rows(
    snapshot: dict[str, Any],
    category_id: str,
) -> tuple[list[dict[str, str]], list[list[Any]], list[Any]]:
    category = snapshot["config"]["categories"][category_id]
    fields = category["fields"]
    rows: list[list[Any]] = []
    field_totals = {field["id"]: 0 for field in fields if field["kind"] == "number"}
    grand_total = 0
    for index, row in enumerate(_snapshot_rows(snapshot, category_id), 1):
        values = _field_value_map(row)
        logical_row: list[Any] = [index, row.get("name", "")]
        for field in fields:
            value = values.get(field["id"], 0 if field["kind"] == "number" else "")
            if field["kind"] == "number":
                value = coerce_nonnegative_int(value)
                field_totals[field["id"]] += value
            logical_row.append(value)
        total = coerce_nonnegative_int(row.get("total", 0))
        grand_total += total
        logical_row.append(total)
        rows.append(logical_row)

    total_row: list[Any] = ["", category.get("total_label", "الإجمالي العام")]
    for field in fields:
        total_row.append(field_totals.get(field["id"], ""))
    total_row.append(grand_total)
    return fields, rows, total_row


def _pdf_paragraph(
    value: Any,
    style: ParagraphStyle,
    *,
    max_line_chars: int = 48,
) -> Paragraph:
    text = str(value)
    if contains_arabic(text) and len(text) > max_line_chars:
        words = text.split()
        lines: list[str] = []
        current = ""
        for word in words:
            candidate = f"{current} {word}".strip()
            if current and len(candidate) > max_line_chars:
                lines.append(current)
                current = word
            else:
                current = candidate
        if current:
            lines.append(current)
        visual = "<br/>".join(escape(rtl_visual(line)) for line in lines)
    else:
        visual = escape(rtl_visual(text) if contains_arabic(text) else text)
    return Paragraph(visual, style)


def export_pdf_to_path(snapshot: dict[str, Any], output_path: Path | str) -> Path:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ثبّت مكتبة reportlab أولًا: pip install reportlab")
    regular_font, bold_font = register_pdf_fonts()
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    theme = REPORT_THEMES.get(
        str(snapshot.get("report_theme", "executive_teal")),
        REPORT_THEMES["executive_teal"],
    )
    pdf_primary = colors.HexColor(theme["primary"])
    pdf_primary_2 = colors.HexColor(theme["primary_2"])
    pdf_primary_soft = colors.HexColor(theme["soft"])
    pdf_gold = colors.HexColor(theme["accent"])
    pdf_gold_dark = colors.HexColor(theme["accent_dark"])
    pdf_gold_soft = colors.HexColor(theme["accent_soft"])
    pdf_ink = colors.HexColor("#243638")
    pdf_muted = colors.HexColor("#667779")
    pdf_border = colors.HexColor(theme["border"])
    pdf_alt_row = colors.HexColor(theme["alt_row"])
    pdf_white = colors.white

    page_size = A4 if str(snapshot.get("report_orientation", "landscape")) == "portrait" else landscape(A4)
    page_width, page_height = page_size
    content_left = 1.0 * cm
    content_width = page_width - 2.0 * cm
    ordered_categories = category_ids(snapshot["config"])
    active_categories = [
        category_id
        for category_id in ordered_categories
        if _snapshot_rows(snapshot, category_id)
    ]
    total_records = sum(
        len(_snapshot_rows(snapshot, category_id))
        for category_id in active_categories
    )
    total_work = sum(
        coerce_nonnegative_int(row.get("total", 0))
        for row in snapshot.get("rows", [])
        if isinstance(row, dict)
    )

    document = SimpleDocTemplate(
        str(output),
        pagesize=page_size,
        leftMargin=content_left,
        rightMargin=content_left,
        topMargin=4.18 * cm,
        bottomMargin=1.68 * cm,
        title=f"{snapshot.get('report_title', APP_NAME)} - {snapshot.get('company_name', '')}",
        author=snapshot.get("user_name", "") or APP_NAME,
    )

    def fitted_font_size(
        text: str,
        font_name: str,
        max_size: float,
        min_size: float,
        max_width: float,
    ) -> float:
        size = max_size
        while size > min_size and pdfmetrics.stringWidth(text, font_name, size) > max_width:
            size -= 0.4
        return max(min_size, size)

    def draw_page(canvas: reportlab_canvas.Canvas, doc: SimpleDocTemplate) -> None:
        canvas.saveState()
        canvas.setStrokeColor(pdf_border)
        canvas.setLineWidth(1.2)
        canvas.roundRect(
            0.55 * cm,
            0.50 * cm,
            page_width - 1.10 * cm,
            page_height - 1.0 * cm,
            7,
        )

        header_x = 0.55 * cm
        header_y = page_height - 2.20 * cm
        header_width = page_width - 1.10 * cm
        header_height = 1.70 * cm
        canvas.setFillColor(pdf_primary)
        canvas.roundRect(
            header_x,
            header_y,
            header_width,
            header_height,
            7,
            fill=1,
            stroke=0,
        )
        canvas.setFillColor(pdf_gold)
        canvas.rect(
            page_width - 0.73 * cm,
            header_y,
            0.18 * cm,
            header_height,
            fill=1,
            stroke=0,
        )

        title_text = rtl_visual(snapshot.get("report_title", APP_NAME))
        title_size = fitted_font_size(
            title_text,
            bold_font,
            18.5,
            13.0,
            page_width - 8.0 * cm,
        )
        canvas.setFillColor(pdf_white)
        canvas.setFont(bold_font, title_size)
        canvas.drawRightString(
            page_width - 1.05 * cm,
            page_height - 1.42 * cm,
            title_text,
        )

        company = snapshot.get("company_name", "") or "بدون اسم شركة"
        company_profile = snapshot.get("company_profile", {})
        if not isinstance(company_profile, dict):
            company_profile = {}
        logo_path = Path(str(snapshot.get("company_logo_path", "")))
        company_text_x = 1.0 * cm
        if logo_path.is_file():
            try:
                canvas.drawImage(
                    str(logo_path),
                    0.94 * cm,
                    page_height - 1.92 * cm,
                    width=1.22 * cm,
                    height=1.22 * cm,
                    preserveAspectRatio=True,
                    anchor="c",
                    mask="auto",
                )
                company_text_x = 2.34 * cm
            except (OSError, ValueError):
                company_text_x = 1.0 * cm
        company_text = rtl_visual(company)
        company_size = fitted_font_size(
            company_text,
            bold_font,
            12.0,
            8.5,
            6.2 * cm - company_text_x,
        )
        canvas.setFillColor(pdf_gold_soft)
        canvas.setFont(regular_font, 7.6)
        canvas.drawString(
            company_text_x,
            page_height - 1.08 * cm,
            rtl_visual("اسم الشركة"),
        )
        canvas.setFont(bold_font, company_size)
        canvas.drawString(
            company_text_x,
            page_height - 1.58 * cm,
            company_text,
        )
        company_detail = " | ".join(
            value
            for value in (
                str(company_profile.get("address", "")).strip(),
                str(company_profile.get("phone", "")).strip(),
            )
            if value
        )
        if company_detail:
            canvas.setFont(regular_font, 6.6)
            canvas.setFillColor(pdf_gold_soft)
            canvas.drawString(
                company_text_x,
                page_height - 1.87 * cm,
                rtl_visual(company_detail[:75]),
            )

        meta_y = page_height - 3.48 * cm
        meta_items = [
            ("المستخدم", snapshot.get("user_name", "") or "-"),
            ("رقم التقرير", snapshot.get("page", "") or "-"),
            ("الرمز", snapshot.get("code", "") or "-"),
            (
                "التاريخ والفترة",
                f"{snapshot.get('date', '')} | {snapshot.get('day', '')}",
            ),
        ]
        gap = 0.22 * cm
        box_width = (content_width - gap * 3) / 4
        box_height = 0.92 * cm
        for index, (label, value) in enumerate(meta_items):
            x = page_width - 1.0 * cm - (index + 1) * box_width - index * gap
            canvas.setFillColor(pdf_primary_soft)
            canvas.roundRect(
                x,
                meta_y,
                box_width,
                box_height,
                5,
                fill=1,
                stroke=0,
            )
            canvas.setFillColor(pdf_gold_dark)
            canvas.rect(
                x + box_width - 0.10 * cm,
                meta_y,
                0.10 * cm,
                box_height,
                fill=1,
                stroke=0,
            )
            canvas.setFillColor(pdf_muted)
            canvas.setFont(regular_font, 7.2)
            canvas.drawRightString(
                x + box_width - 0.24 * cm,
                meta_y + 0.59 * cm,
                rtl_visual(label),
            )
            value_text = rtl_visual(value)
            value_size = fitted_font_size(
                value_text,
                bold_font,
                9.3,
                7.1,
                box_width - 0.42 * cm,
            )
            canvas.setFillColor(pdf_primary)
            canvas.setFont(bold_font, value_size)
            canvas.drawRightString(
                x + box_width - 0.24 * cm,
                meta_y + 0.22 * cm,
                value_text,
            )

        canvas.setStrokeColor(pdf_gold)
        canvas.setLineWidth(1.1)
        canvas.line(1.0 * cm, 1.12 * cm, page_width - 1.0 * cm, 1.12 * cm)
        canvas.setFillColor(pdf_muted)
        canvas.setFont(regular_font, 7.5)
        created_at = snapshot.get("created_at", "").replace("T", " ")
        left_footer = (
            f"v{APP_VERSION}  |  {rtl_visual('صفحة')} {doc.page}"
            f"  |  {created_at}"
            f"  |  R{snapshot.get('revision_number', 0) or '-'}"
            f"  |  {str(snapshot.get('verification_code', ''))[:16].upper() or '-'}"
        )
        canvas.drawString(1.0 * cm, 0.78 * cm, left_footer)
        custom_footer = str(snapshot.get("report_footer", "")).strip()
        canvas.drawCentredString(
            page_width / 2,
            0.78 * cm,
            rtl_visual(custom_footer or f"مستخدم البرنامج: {snapshot.get('user_name', '') or '-'}"),
        )
        canvas.drawRightString(
            page_width - 1.0 * cm,
            0.78 * cm,
            rtl_visual(f"الشركة: {company}"),
        )
        canvas.restoreState()

    normal_style = ParagraphStyle(
        "ArabicNormal",
        fontName=regular_font,
        fontSize=10.2,
        leading=14.2,
        alignment=1,
        textColor=pdf_ink,
    )
    header_style = ParagraphStyle(
        "ArabicHeader",
        parent=normal_style,
        fontName=bold_font,
        fontSize=9.8,
        leading=12.5,
        textColor=pdf_white,
    )
    total_style = ParagraphStyle(
        "ArabicTotal",
        parent=normal_style,
        fontName=bold_font,
        fontSize=10.4,
        textColor=pdf_white,
    )
    section_style = ParagraphStyle(
        "ArabicSection",
        fontName=bold_font,
        fontSize=12.8,
        leading=16,
        alignment=2,
        textColor=pdf_primary,
    )
    section_count_style = ParagraphStyle(
        "ArabicSectionCount",
        fontName=regular_font,
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=pdf_gold_dark,
    )
    summary_heading_style = ParagraphStyle(
        "ArabicSummaryHeading",
        fontName=bold_font,
        fontSize=10.5,
        leading=13,
        alignment=2,
        textColor=pdf_primary,
        spaceAfter=5,
    )
    metric_value_style = ParagraphStyle(
        "ArabicMetricValue",
        fontName=bold_font,
        fontSize=14,
        leading=16,
        alignment=1,
        textColor=pdf_primary,
    )
    metric_label_style = ParagraphStyle(
        "ArabicMetricLabel",
        fontName=regular_font,
        fontSize=8.3,
        leading=10,
        alignment=1,
        textColor=pdf_muted,
    )
    empty_style = ParagraphStyle(
        "ArabicEmpty",
        parent=normal_style,
        alignment=2,
        textColor=pdf_muted,
        spaceAfter=8,
    )

    story: list[Any] = []
    summary_items: list[tuple[str, Any]] = [
        ("إجمالي السجلات", total_records),
        ("إجمالي الأعمال", total_work),
    ]
    if snapshot.get("revision_number"):
        summary_items.extend(
            [
                ("إصدار التقرير", snapshot.get("revision_number")),
                ("رمز التحقق", str(snapshot.get("verification_code", ""))[:16].upper()),
            ]
        )
    for category_id in active_categories:
        category = snapshot["config"]["categories"][category_id]
        count = len(_snapshot_rows(snapshot, category_id))
        summary_items.append((category["label"], count))

    story.append(Paragraph(rtl_visual("ملخص التقرير"), summary_heading_style))
    for row_start in range(0, len(summary_items), 4):
        item_row = summary_items[row_start : row_start + 4]
        cards: list[Any] = []
        for label, value in reversed(item_row):
            card = Table(
                [
                    [_pdf_paragraph(value, metric_value_style)],
                    [_pdf_paragraph(label, metric_label_style, max_line_chars=28)],
                ],
                colWidths=[content_width / len(item_row) - 0.12 * cm],
                rowHeights=[0.52 * cm, 0.40 * cm],
            )
            card.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), pdf_primary_soft),
                        ("BOX", (0, 0), (-1, -1), 0.55, pdf_border),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 2),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ]
                )
            )
            cards.append(card)
        summary_row = Table(
            [cards],
            colWidths=[content_width / len(item_row)] * len(item_row),
            hAlign="CENTER",
        )
        summary_row.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(summary_row)
    story.append(Spacer(1, 0.12 * cm))

    for category_id in ordered_categories:
        category = snapshot["config"]["categories"][category_id]
        fields, logical_rows, total_row = _logical_table_rows(snapshot, category_id)
        if not logical_rows:
            continue

        logical_headers = [
            "#",
            category.get("person_label", "الاسم"),
            *[field["label"] for field in fields],
            category.get("total_label", "الإجمالي"),
        ]
        logical_weights = [0.55, 2.25]
        logical_weights.extend(2.3 if field["kind"] == "text" else 1.25 for field in fields)
        logical_weights.append(1.1)
        weight_sum = sum(logical_weights)
        logical_widths = [content_width * weight / weight_sum for weight in logical_weights]

        has_text_field = any(field["kind"] == "text" for field in fields)
        column_count = len(logical_headers)
        table_font_size = 10.2 if column_count <= 7 else 9.2 if column_count <= 9 else 8.2
        table_leading = table_font_size + 3.7
        category_normal_style = ParagraphStyle(
            f"ArabicNormal_{category_id}",
            parent=normal_style,
            fontSize=table_font_size,
            leading=table_leading,
        )
        category_header_style = ParagraphStyle(
            f"ArabicHeader_{category_id}",
            parent=header_style,
            fontSize=max(8.0, table_font_size - 0.2),
            leading=max(10.6, table_leading - 1.2),
        )
        category_total_style = ParagraphStyle(
            f"ArabicTotal_{category_id}",
            parent=total_style,
            fontSize=table_font_size,
            leading=table_leading,
        )
        rows_per_page = 6 if has_text_field else 8
        chunks = [
            logical_rows[index : index + rows_per_page]
            for index in range(0, len(logical_rows), rows_per_page)
        ]
        for chunk_index, chunk in enumerate(chunks):
            section_title = category["label"]
            if chunk_index:
                section_title = f"{section_title} - تابع"
            start_record = chunk_index * rows_per_page + 1
            end_record = start_record + len(chunk) - 1
            record_range = (
                f"{len(logical_rows)} سجل"
                if len(chunks) == 1
                else f"{start_record}-{end_record} من {len(logical_rows)}"
            )
            section_band = Table(
                [
                    [
                        _pdf_paragraph(record_range, section_count_style),
                        _pdf_paragraph(section_title, section_style),
                    ]
                ],
                colWidths=[3.0 * cm, content_width - 3.0 * cm],
            )
            section_band.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), pdf_gold_soft),
                        ("LINEBELOW", (0, 0), (-1, -1), 0.8, pdf_gold),
                        ("LINEAFTER", (1, 0), (1, 0), 4.0, pdf_gold),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )
            table_data: list[list[Any]] = [
                [
                    _pdf_paragraph(value, category_header_style, max_line_chars=24)
                    for value in reversed(logical_headers)
                ]
            ]
            for logical_row in chunk:
                table_data.append(
                    [
                        _pdf_paragraph(value, category_normal_style)
                        for value in reversed(logical_row)
                    ]
                )
            is_final_chunk = chunk_index == len(chunks) - 1
            if is_final_chunk:
                table_data.append(
                    [
                        _pdf_paragraph(value, category_total_style)
                        for value in reversed(total_row)
                    ]
                )

            table = Table(
                table_data,
                colWidths=list(reversed(logical_widths)),
                repeatRows=1,
                hAlign="CENTER",
                splitByRow=0,
            )
            style_commands: list[tuple[Any, ...]] = [
                ("BACKGROUND", (0, 0), (-1, 0), pdf_primary_2),
                ("TEXTCOLOR", (0, 0), (-1, 0), pdf_white),
                ("BOX", (0, 0), (-1, -1), 0.9, pdf_primary_2),
                ("INNERGRID", (0, 0), (-1, -1), 0.30, pdf_border),
                ("LINEBELOW", (0, 0), (-1, 0), 1.0, pdf_primary),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
            body_end = len(table_data) - 1 if is_final_chunk else len(table_data)
            for row_index in range(1, body_end):
                background = pdf_white if row_index % 2 else pdf_alt_row
                style_commands.append(
                    (
                        "BACKGROUND",
                        (0, row_index),
                        (-1, row_index),
                        background,
                    )
                )
            if is_final_chunk:
                style_commands.extend(
                    [
                        (
                            "BACKGROUND",
                            (0, -1),
                            (-1, -1),
                            pdf_gold_dark,
                        ),
                        ("TEXTCOLOR", (0, -1), (-1, -1), pdf_white),
                        ("LINEABOVE", (0, -1), (-1, -1), 1.0, pdf_gold_dark),
                    ]
                )
            table.setStyle(TableStyle(style_commands))
            story.extend(
                [
                    KeepTogether(
                        [
                            section_band,
                            Spacer(1, 0.08 * cm),
                            table,
                        ]
                    ),
                    Spacer(1, 0.12 * cm),
                ]
            )
    document.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    if not output.exists() or output.stat().st_size < 1000:
        raise RuntimeError("تعذر إنشاء ملف PDF صالح.")
    return output


class ImageReportRenderer:
    WIDTH = 2480
    HEIGHT = 1754
    LEFT = 90
    RIGHT = 90
    CONTENT_TOP = 430
    CONTENT_BOTTOM = 1565

    def __init__(self, snapshot: dict[str, Any]):
        if not PILLOW_AVAILABLE:
            raise RuntimeError("ثبّت مكتبة Pillow أولًا: pip install Pillow")
        self.snapshot = snapshot
        self.theme = REPORT_THEMES.get(
            str(snapshot.get("report_theme", "executive_teal")),
            REPORT_THEMES["executive_teal"],
        )
        regular_path, bold_path = resolve_font_paths()
        self.fonts = {
            "brand": ImageFont.truetype(str(bold_path), 50),
            "title": ImageFont.truetype(str(bold_path), 52),
            "section": ImageFont.truetype(str(bold_path), 38),
            "header": ImageFont.truetype(str(bold_path), 30),
            "body": ImageFont.truetype(str(regular_path), 31),
            "body_bold": ImageFont.truetype(str(bold_path), 32),
            "small": ImageFont.truetype(str(regular_path), 23),
            "small_bold": ImageFont.truetype(str(bold_path), 23),
        }
        self.pages: list[Image.Image] = []
        self.draws: list[ImageDraw.ImageDraw] = []
        self.current_image: Image.Image | None = None
        self.draw: ImageDraw.ImageDraw | None = None
        self.y = self.CONTENT_TOP

    def _start_page(self, section_label: str | None = None) -> None:
        image = Image.new("RGB", (self.WIDTH, self.HEIGHT), Palette.WHITE)
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle(
            [18, 18, self.WIDTH - 18, self.HEIGHT - 18],
            radius=22,
            outline=self.theme["soft"],
            width=10,
        )
        draw.rounded_rectangle(
            [36, 36, self.WIDTH - 36, self.HEIGHT - 36],
            radius=18,
            outline=self.theme["accent"],
            width=4,
        )
        draw.rectangle(
            [36, 36, self.WIDTH - 36, 255],
            fill=self.theme["primary"],
        )
        company_x = self.LEFT
        logo_path = Path(str(self.snapshot.get("company_logo_path", "")))
        if logo_path.is_file():
            try:
                logo = Image.open(logo_path).convert("RGBA")
                logo.thumbnail((145, 145), Image.Resampling.LANCZOS)
                image.paste(logo, (self.LEFT, 73), logo)
                company_x += 170
            except (OSError, ValueError):
                company_x = self.LEFT
        draw.text(
            (company_x, 91),
            rtl_visual(self.snapshot.get("company_name", "") or "بدون اسم شركة"),
            fill=self.theme["accent_soft"],
            font=self.fonts["brand"],
        )
        draw.text(
            (self.WIDTH - self.RIGHT, 92),
            rtl_visual(self.snapshot.get("report_title", APP_NAME)),
            fill=Palette.WHITE,
            font=self.fonts["title"],
            anchor="ra",
        )

        meta_y = 288
        meta_width = (self.WIDTH - self.LEFT - self.RIGHT - 75) // 4
        meta_values = [
            rtl_visual(
                f"التاريخ: {self.snapshot.get('date', '')} | "
                f"{self.snapshot.get('day', '')}"
            ),
            rtl_visual(f"الرمز: {self.snapshot.get('code', '') or '-'}"),
            rtl_visual(f"رقم التقرير: {self.snapshot.get('page', '')}"),
            rtl_visual(f"المستخدم: {self.snapshot.get('user_name', '') or '-'}"),
        ]
        for index, value in enumerate(meta_values):
            x1 = self.LEFT + index * (meta_width + 25)
            draw.rounded_rectangle(
                [x1, meta_y, x1 + meta_width, meta_y + 82],
                radius=12,
                fill=self.theme["soft"],
                outline=self.theme["border"],
                width=2,
            )
            draw.text(
                (x1 + meta_width // 2, meta_y + 36),
                value,
                fill=self.theme["primary"],
                font=self.fonts["small"],
                anchor="mm",
            )

        self.pages.append(image)
        self.draws.append(draw)
        self.current_image = image
        self.draw = draw
        self.y = self.CONTENT_TOP
        if section_label:
            self._draw_section_title(section_label)

    def _draw_section_title(self, label: str) -> None:
        assert self.draw is not None
        self.draw.rounded_rectangle(
            [self.LEFT, self.y, self.WIDTH - self.RIGHT, self.y + 72],
            radius=14,
            fill=self.theme["primary_2"],
        )
        self.draw.text(
            (self.WIDTH - self.RIGHT - 28, self.y + 36),
            rtl_visual(label),
            fill=Palette.WHITE,
            font=self.fonts["section"],
            anchor="rm",
        )
        self.y += 90

    def _text_width(self, text: str, font: ImageFont.FreeTypeFont) -> int:
        assert self.draw is not None
        box = self.draw.textbbox((0, 0), text, font=font)
        return box[2] - box[0]

    def _wrap_text(
        self,
        value: Any,
        font: ImageFont.FreeTypeFont,
        max_width: int,
        max_lines: int = 3,
    ) -> list[str]:
        text = str(value)
        if not text:
            return [""]
        words = text.split()
        if not words:
            return [""]
        lines: list[str] = []
        current = ""
        for word in words:
            candidate = f"{current} {word}".strip()
            candidate_visual = rtl_visual(candidate) if contains_arabic(candidate) else candidate
            if current and self._text_width(candidate_visual, font) > max_width:
                lines.append(current)
                current = word
            else:
                current = candidate
        if current:
            lines.append(current)
        if len(lines) > max_lines:
            lines = lines[:max_lines]
            lines[-1] = lines[-1].rstrip() + "..."
        return [rtl_visual(line) if contains_arabic(line) else line for line in lines]

    def _column_widths(
        self,
        fields: list[dict[str, str]],
    ) -> list[int]:
        weights = [0.6, 2.2]
        weights.extend(2.4 if field["kind"] == "text" else 1.25 for field in fields)
        weights.append(1.0)
        available = self.WIDTH - self.LEFT - self.RIGHT
        total_weight = sum(weights)
        widths = [int(available * weight / total_weight) for weight in weights]
        widths[-1] += available - sum(widths)
        return widths

    def _draw_table_header(
        self,
        category: dict[str, Any],
        fields: list[dict[str, str]],
        logical_widths: list[int],
    ) -> None:
        assert self.draw is not None
        logical_headers = [
            "#",
            category.get("person_label", "الاسم"),
            *[field["label"] for field in fields],
            category.get("total_label", "الإجمالي"),
        ]
        x = self.LEFT
        height = 86
        for value, width in zip(reversed(logical_headers), reversed(logical_widths)):
            self.draw.rectangle(
                [x, self.y, x + width, self.y + height],
                fill=self.theme["primary_2"],
                outline=self.theme["soft"],
                width=2,
            )
            lines = self._wrap_text(
                value,
                self.fonts["header"],
                max(30, width - 18),
                max_lines=2,
            )
            first_y = self.y + height // 2 - (len(lines) - 1) * 16
            for line_index, line in enumerate(lines):
                self.draw.text(
                    (x + width // 2, first_y + line_index * 32),
                    line,
                    fill=Palette.WHITE,
                    font=self.fonts["header"],
                    anchor="mm",
                )
            x += width
        self.y += height

    def _row_height(
        self,
        logical_values: list[Any],
        logical_widths: list[int],
    ) -> int:
        line_count = 1
        for value, width in zip(logical_values, logical_widths):
            lines = self._wrap_text(value, self.fonts["body"], max(30, width - 24))
            line_count = max(line_count, len(lines))
        return max(72, 22 + line_count * 36)

    def _draw_table_row(
        self,
        logical_values: list[Any],
        logical_widths: list[int],
        *,
        background: str,
        total_row: bool = False,
    ) -> None:
        assert self.draw is not None
        height = self._row_height(logical_values, logical_widths)
        x = self.LEFT
        values = list(reversed(logical_values))
        widths = list(reversed(logical_widths))
        for value, width in zip(values, widths):
            self.draw.rectangle(
                [x, self.y, x + width, self.y + height],
                fill=background,
                outline=self.theme["border"] if not total_row else self.theme["accent_soft"],
                width=2,
            )
            font = self.fonts["body_bold"] if total_row else self.fonts["body"]
            color = Palette.WHITE if total_row else Palette.INK
            lines = self._wrap_text(value, font, max(30, width - 24))
            first_y = self.y + height // 2 - (len(lines) - 1) * 18
            for line_index, line in enumerate(lines):
                self.draw.text(
                    (x + width // 2, first_y + line_index * 36),
                    line,
                    fill=color,
                    font=font,
                    anchor="mm",
                )
            x += width
        self.y += height

    def render(self) -> list[Image.Image]:
        any_rows = False
        for category_id in category_ids(self.snapshot["config"]):
            category = self.snapshot["config"]["categories"][category_id]
            fields, rows, total_row = _logical_table_rows(self.snapshot, category_id)
            if not rows:
                continue
            any_rows = True
            self._start_page(category["label"])
            widths = self._column_widths(fields)
            self._draw_table_header(category, fields, widths)
            for row_index, logical_row in enumerate(rows, 1):
                needed = self._row_height(logical_row, widths)
                reserve_for_total = (
                    self._row_height(total_row, widths)
                    if row_index == len(rows)
                    else 0
                )
                if self.y + needed + reserve_for_total > self.CONTENT_BOTTOM:
                    self._start_page(category["label"])
                    self._draw_table_header(category, fields, widths)
                self._draw_table_row(
                    logical_row,
                    widths,
                    background=Palette.WHITE if row_index % 2 else Palette.ALT_ROW,
                )
            total_height = self._row_height(total_row, widths)
            if self.y + total_height > self.CONTENT_BOTTOM:
                self._start_page(category["label"])
                self._draw_table_header(category, fields, widths)
            self._draw_table_row(
                total_row,
                widths,
                background=Palette.ACCENT,
                total_row=True,
            )

        if not any_rows:
            self._start_page()
            assert self.draw is not None
            self.draw.text(
                (self.WIDTH // 2, 750),
                rtl_visual("لا توجد بيانات للتصدير"),
                fill=Palette.MUTED,
                font=self.fonts["title"],
                anchor="mm",
            )

        total_pages = len(self.pages)
        for index, draw in enumerate(self.draws, 1):
            draw.line(
                [(self.LEFT, 1640), (self.WIDTH - self.RIGHT, 1640)],
                fill=Palette.ACCENT,
                width=4,
            )
            company = self.snapshot.get("company_name", "") or "بدون اسم شركة"
            stamp = (
                f"{rtl_visual('وقت الإصدار')}: {self.snapshot.get('created_at', '').replace('T', ' ')}"
                f" | {rtl_visual('المستخدم')}: {rtl_visual(self.snapshot.get('user_name', '') or '-')}"
                f" | {rtl_visual('الشركة')}: {rtl_visual(company)}"
            )
            draw.text(
                (self.LEFT, 1674),
                stamp,
                fill=Palette.MUTED,
                font=self.fonts["small"],
                anchor="lm",
            )
            draw.text(
                (self.WIDTH - self.RIGHT, 1674),
                f"{index} / {total_pages}",
                fill=Palette.MUTED,
                font=self.fonts["small"],
                anchor="rm",
            )
        return self.pages


def export_images_to_path(
    snapshot: dict[str, Any],
    output_path: Path | str,
) -> list[Path]:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    pages = ImageReportRenderer(snapshot).render()
    saved: list[Path] = []
    if len(pages) == 1:
        pages[0].save(output, format="PNG", dpi=(300, 300), optimize=True)
        saved.append(output)
    else:
        for index, page in enumerate(pages, 1):
            page_path = output.with_name(f"{output.stem}_{index:02d}.png")
            page.save(page_path, format="PNG", dpi=(300, 300), optimize=True)
            saved.append(page_path)
    return saved


def export_csv_to_path(snapshot: dict[str, Any], output_path: Path | str) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(
            [
                snapshot.get("report_title", APP_NAME),
                snapshot.get("company_name", ""),
            ]
        )
        writer.writerow(
            [
                "التاريخ",
                snapshot.get("date", ""),
                "اليوم",
                snapshot.get("day", ""),
                "الرمز",
                snapshot.get("code", ""),
                "المستخدم",
                snapshot.get("user_name", ""),
                "وقت الإصدار",
                snapshot.get("created_at", ""),
            ]
        )
        writer.writerow([])
        for category_id in category_ids(snapshot["config"]):
            category = snapshot["config"]["categories"][category_id]
            fields, rows, total_row = _logical_table_rows(snapshot, category_id)
            writer.writerow([category["label"]])
            writer.writerow(
                [
                    "#",
                    category.get("person_label", "الاسم"),
                    *[field["label"] for field in fields],
                    category.get("total_label", "الإجمالي"),
                ]
            )
            writer.writerows(rows)
            writer.writerow(total_row)
            writer.writerow([])
    return output


def _excel_prepare_sheet(sheet: Any, title: str) -> None:
    sheet.sheet_view.rightToLeft = True
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.oddHeader.center.text = f"&B{title}"
    sheet.oddFooter.left.text = f"{APP_NAME} - v{APP_VERSION}"
    sheet.oddFooter.right.text = "صفحة &P من &N"


def _excel_datetime(value: Any) -> datetime | str:
    try:
        return datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return str(value)


def _excel_style_table(sheet: Any, table_name: str) -> None:
    if sheet.max_row < 1 or sheet.max_column < 1:
        return
    primary_fill = PatternFill("solid", fgColor="214E52")
    accent_fill = PatternFill("solid", fgColor="C39A5B")
    alternate_fill = PatternFill("solid", fgColor="F3F7F6")
    white_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, color="243638")
    border = Border(
        bottom=Side(style="thin", color="CFDDDA"),
    )
    for cell in sheet[1]:
        cell.fill = primary_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 28
    for row_index in range(2, sheet.max_row + 1):
        for cell in sheet[row_index]:
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(
                horizontal="right" if isinstance(cell.value, str) else "center",
                vertical="center",
                wrap_text=True,
            )
            if row_index % 2 == 0:
                cell.fill = alternate_fill
    for column_index in range(1, sheet.max_column + 1):
        values = [
            str(sheet.cell(row=row_index, column=column_index).value or "")
            for row_index in range(1, min(sheet.max_row, 250) + 1)
        ]
        width = min(60, max(12, max((len(value) for value in values), default=0) + 3))
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    if sheet.max_row >= 2:
        table = ExcelTable(
            displayName=table_name,
            ref=f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    sheet.sheet_properties.tabColor = accent_fill.fgColor.rgb
    _excel_prepare_sheet(sheet, sheet.title)


def export_snapshot_excel_to_path(
    snapshot: dict[str, Any],
    output_path: Path | str,
) -> Path:
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("تصدير Excel يحتاج مكتبة openpyxl.")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "الملخص"
    total_work = sum(
        coerce_nonnegative_int(row.get("total", 0))
        for row in snapshot.get("rows", [])
        if isinstance(row, dict)
    )
    summary_rows = [
        ("المؤشر", "القيمة"),
        ("عنوان التقرير", snapshot.get("report_title", APP_NAME)),
        ("الشركة", snapshot.get("company_name", "")),
        ("المستخدم", snapshot.get("user_name", "")),
        ("التاريخ", parse_iso_date(snapshot.get("date")) or snapshot.get("date", "")),
        ("الرمز", snapshot.get("code", "")),
        ("الحالة", REPORT_STATUSES.get(str(snapshot.get("status", "draft")), "مسودة")),
        ("عدد السجلات", len(snapshot.get("rows", []))),
        ("إجمالي الأعمال", total_work),
        ("إصدار التقرير", snapshot.get("revision_number", "")),
        ("رمز التحقق", str(snapshot.get("verification_code", "")).upper()),
        ("وقت الإصدار", _excel_datetime(snapshot.get("created_at", ""))),
    ]
    for row in summary_rows:
        summary.append(row)
    _excel_style_table(summary, "ReportSummary")
    summary["B5"].number_format = "yyyy-mm-dd"
    summary[f"B{len(summary_rows)}"].number_format = "yyyy-mm-dd hh:mm"

    records = workbook.create_sheet("السجلات")
    records.append(
        [
            "التاريخ",
            "الشركة",
            "المستخدم",
            "الرمز",
            "الفئة",
            "الاسم/الجهة",
            "تفاصيل العمل",
            "الإجمالي",
        ]
    )
    for row in snapshot.get("rows", []):
        if not isinstance(row, dict):
            continue
        details = " | ".join(
            f"{field.get('label', '')}: {field.get('value', '')}"
            for field in row.get("fields", [])
            if isinstance(field, dict)
        )
        records.append(
            [
                parse_iso_date(snapshot.get("date")) or snapshot.get("date", ""),
                snapshot.get("company_name", ""),
                snapshot.get("user_name", ""),
                snapshot.get("code", ""),
                row.get("category_label", ""),
                row.get("name", ""),
                details,
                coerce_nonnegative_int(row.get("total", 0)),
            ]
        )
    _excel_style_table(records, "ReportRecords")
    for cell in records["A"][1:]:
        cell.number_format = "yyyy-mm-dd"
    records.column_dimensions["G"].width = 55

    categories = workbook.create_sheet("الفئات")
    categories.append(["الفئة", "عدد السجلات", "إجمالي الأعمال"])
    category_values: dict[str, list[int]] = {}
    for row in snapshot.get("rows", []):
        if not isinstance(row, dict):
            continue
        label = str(row.get("category_label", ""))
        values = category_values.setdefault(label, [0, 0])
        values[0] += 1
        values[1] += coerce_nonnegative_int(row.get("total", 0))
    for label, values in sorted(
        category_values.items(),
        key=lambda item: item[1][1],
        reverse=True,
    ):
        categories.append([label, values[0], values[1]])
    _excel_style_table(categories, "ReportCategories")
    if categories.max_row >= 2:
        chart = BarChart()
        chart.title = "حجم الأعمال حسب الفئة"
        chart.y_axis.title = "إجمالي الأعمال"
        chart.x_axis.title = "الفئة"
        chart.add_data(
            Reference(categories, min_col=3, min_row=1, max_row=categories.max_row),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(categories, min_col=1, min_row=2, max_row=categories.max_row)
        )
        chart.height = 8
        chart.width = 16
        categories.add_chart(chart, "E2")
    workbook.save(output)
    return output


def export_analytics_excel_to_path(
    analytics: dict[str, Any],
    company: dict[str, Any] | None,
    output_path: Path | str,
) -> Path:
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("تصدير Excel يحتاج مكتبة openpyxl.")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "الملخص الإداري"
    summary.append(["المؤشر", "القيمة"])
    for label, value in (
        ("الشركة", (company or {}).get("name", "كل الشركات")),
        (
            "من تاريخ",
            parse_iso_date(analytics.get("start_date"))
            or analytics.get("start_date", ""),
        ),
        (
            "إلى تاريخ",
            parse_iso_date(analytics.get("end_date"))
            or analytics.get("end_date", ""),
        ),
        ("عدد التقارير", analytics.get("reports", 0)),
        ("عدد السجلات", analytics.get("records", 0)),
        ("إجمالي الأعمال", analytics.get("work_total", 0)),
        ("المتوسط لكل تقرير", analytics.get("average_per_report", 0)),
        ("آخر تحديث للبيانات", _excel_datetime(analytics.get("freshness", ""))),
    ):
        summary.append([label, value])
    _excel_style_table(summary, "ManagementSummary")
    summary["B3"].number_format = "yyyy-mm-dd"
    summary["B4"].number_format = "yyyy-mm-dd"
    summary["B9"].number_format = "yyyy-mm-dd hh:mm"

    def add_data_sheet(
        title: str,
        headers: list[str],
        rows: list[dict[str, Any]],
        fields: list[str],
        table_name: str,
    ) -> Any:
        sheet = workbook.create_sheet(title)
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(field, "") for field in fields])
        _excel_style_table(sheet, table_name)
        return sheet

    daily = add_data_sheet(
        "الاتجاه اليومي",
        ["التاريخ", "التقارير", "السجلات", "إجمالي الأعمال"],
        analytics.get("daily", []),
        ["label", "reports", "records", "total"],
        "DailyTrend",
    )
    for cell in daily["A"][1:]:
        parsed_date = parse_iso_date(cell.value)
        if parsed_date:
            cell.value = parsed_date
            cell.number_format = "yyyy-mm-dd"
    categories = add_data_sheet(
        "مقارنة الفئات",
        ["الفئة", "السجلات", "إجمالي الأعمال"],
        analytics.get("categories", []),
        ["label", "records", "total"],
        "CategoryComparison",
    )
    people = add_data_sheet(
        "مقارنة الأسماء",
        ["الاسم/الجهة", "السجلات", "إجمالي الأعمال"],
        analytics.get("people", []),
        ["label", "records", "total"],
        "PeopleComparison",
    )
    add_data_sheet(
        "حالات التقارير",
        ["الحالة", "عدد التقارير"],
        [
            {
                "label": REPORT_STATUSES.get(str(row.get("label", "")), row.get("label", "")),
                "reports": row.get("reports", 0),
            }
            for row in analytics.get("statuses", [])
        ],
        ["label", "reports"],
        "StatusComparison",
    )
    if daily.max_row >= 2:
        chart = LineChart()
        chart.title = "الإنتاجية خلال الفترة"
        chart.y_axis.title = "إجمالي الأعمال"
        chart.x_axis.title = "التاريخ"
        chart.add_data(
            Reference(daily, min_col=4, min_row=1, max_row=daily.max_row),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(daily, min_col=1, min_row=2, max_row=daily.max_row)
        )
        chart.height = 8
        chart.width = 17
        daily.add_chart(chart, "F2")
    if categories.max_row >= 2:
        chart = BarChart()
        chart.title = "حجم الأعمال حسب الفئة"
        chart.add_data(
            Reference(categories, min_col=3, min_row=1, max_row=categories.max_row),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(categories, min_col=1, min_row=2, max_row=categories.max_row)
        )
        chart.height = 8
        chart.width = 17
        categories.add_chart(chart, "E2")
    if people.max_row >= 2:
        chart = BarChart()
        chart.title = "حجم الأعمال حسب الاسم أو الجهة"
        chart.add_data(
            Reference(people, min_col=3, min_row=1, max_row=people.max_row),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(people, min_col=1, min_row=2, max_row=people.max_row)
        )
        chart.height = 8
        chart.width = 17
        people.add_chart(chart, "E2")
    workbook.save(output)
    return output


def export_management_pdf_to_path(
    analytics: dict[str, Any],
    company: dict[str, Any] | None,
    output_path: Path | str,
) -> Path:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("تصدير PDF يحتاج مكتبة reportlab.")
    regular_font, bold_font = register_pdf_fonts()
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    page_width, page_height = landscape(A4)
    pdf = reportlab_canvas.Canvas(str(output), pagesize=(page_width, page_height))
    primary = colors.HexColor("#214E52")
    accent = colors.HexColor("#C39A5B")
    soft = colors.HexColor("#E9F2F0")
    ink = colors.HexColor("#243638")
    muted = colors.HexColor("#667779")
    border = colors.HexColor("#CFDDDA")
    company_name = str((company or {}).get("name", "كل الشركات"))
    logo_path = Path(str((company or {}).get("logo_path", "")))

    pdf.setFillColor(primary)
    pdf.roundRect(
        0.7 * cm,
        page_height - 2.4 * cm,
        page_width - 1.4 * cm,
        1.65 * cm,
        8,
        fill=1,
        stroke=0,
    )
    pdf.setFillColor(accent)
    pdf.rect(
        page_width - 0.92 * cm,
        page_height - 2.4 * cm,
        0.22 * cm,
        1.65 * cm,
        fill=1,
        stroke=0,
    )
    pdf.setFillColor(colors.white)
    pdf.setFont(bold_font, 19)
    pdf.drawRightString(
        page_width - 1.2 * cm,
        page_height - 1.52 * cm,
        rtl_visual("التقرير الإداري وتحليل الإنتاجية"),
    )
    logo_offset = 0
    if logo_path.is_file():
        try:
            pdf.drawImage(
                str(logo_path),
                1.0 * cm,
                page_height - 2.18 * cm,
                width=1.25 * cm,
                height=1.25 * cm,
                preserveAspectRatio=True,
                mask="auto",
            )
            logo_offset = 1.45 * cm
        except (OSError, ValueError):
            logo_offset = 0
    pdf.setFont(bold_font, 10)
    pdf.drawString(1.0 * cm + logo_offset, page_height - 1.42 * cm, rtl_visual(company_name))
    pdf.setFont(regular_font, 8)
    pdf.drawString(
        1.0 * cm + logo_offset,
        page_height - 1.87 * cm,
        f"{analytics.get('start_date', '')} - {analytics.get('end_date', '')}",
    )

    cards = [
        ("عدد التقارير", analytics.get("reports", 0)),
        ("عدد السجلات", analytics.get("records", 0)),
        ("إجمالي الأعمال", analytics.get("work_total", 0)),
        ("المتوسط / تقرير", analytics.get("average_per_report", 0)),
    ]
    card_y = page_height - 4.05 * cm
    card_gap = 0.25 * cm
    card_width = (page_width - 2.0 * cm - card_gap * 3) / 4
    for index, (label, value) in enumerate(cards):
        x = page_width - 1.0 * cm - (index + 1) * card_width - index * card_gap
        pdf.setFillColor(soft)
        pdf.roundRect(
            x,
            card_y,
            card_width,
            1.15 * cm,
            7,
            fill=1,
            stroke=0,
        )
        pdf.setFillColor(primary)
        pdf.setFont(bold_font, 16)
        pdf.drawCentredString(x + card_width / 2, card_y + 0.55 * cm, str(value))
        pdf.setFillColor(muted)
        pdf.setFont(regular_font, 8)
        pdf.drawCentredString(
            x + card_width / 2,
            card_y + 0.20 * cm,
            rtl_visual(label),
        )

    def draw_bar_panel(
        title: str,
        rows: list[dict[str, Any]],
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> None:
        pdf.setStrokeColor(border)
        pdf.setFillColor(colors.white)
        pdf.roundRect(x, y, width, height, 7, fill=1, stroke=1)
        pdf.setFillColor(primary)
        pdf.setFont(bold_font, 11)
        pdf.drawRightString(x + width - 0.35 * cm, y + height - 0.55 * cm, rtl_visual(title))
        shown = rows[:7]
        max_value = max(
            [coerce_nonnegative_int(row.get("total", 0)) for row in shown] or [1]
        )
        row_height = (height - 1.05 * cm) / max(1, len(shown))
        for index, row in enumerate(shown):
            row_y = y + height - 1.0 * cm - (index + 1) * row_height
            label = str(row.get("label", ""))
            value = coerce_nonnegative_int(row.get("total", 0))
            pdf.setFillColor(soft if index % 2 == 0 else colors.white)
            pdf.rect(
                x + 0.2 * cm,
                row_y,
                width - 0.4 * cm,
                row_height,
                fill=1,
                stroke=0,
            )
            bar_width = (width * 0.42) * (value / max_value if max_value else 0)
            bar_height = min(0.40 * cm, max(4, row_height - 0.28 * cm))
            bar_y = row_y + (row_height - bar_height) / 2
            pdf.setFillColor(accent)
            pdf.roundRect(
                x + 0.35 * cm,
                bar_y,
                bar_width,
                bar_height,
                3,
                fill=1,
                stroke=0,
            )
            pdf.setFillColor(ink)
            pdf.setFont(regular_font, 7.5)
            pdf.drawRightString(
                x + width - 0.35 * cm,
                row_y + row_height / 2 - 2,
                rtl_visual(label[:38]),
            )
            pdf.setFont(bold_font, 7.5)
            pdf.drawString(x + 0.35 * cm + bar_width + 4, row_y + row_height / 2 - 2, str(value))

    panel_y = 2.05 * cm
    panel_height = card_y - panel_y - 0.38 * cm
    panel_width = (page_width - 2.25 * cm) / 2
    draw_bar_panel(
        "مقارنة الفئات",
        analytics.get("categories", []),
        page_width - 1.0 * cm - panel_width,
        panel_y,
        panel_width,
        panel_height,
    )
    draw_bar_panel(
        "الأسماء والجهات الأعلى إنتاجية",
        analytics.get("people", []),
        1.0 * cm,
        panel_y,
        panel_width,
        panel_height,
    )
    pdf.setStrokeColor(accent)
    pdf.line(1.0 * cm, 1.42 * cm, page_width - 1.0 * cm, 1.42 * cm)
    pdf.setFillColor(muted)
    pdf.setFont(regular_font, 7.5)
    footer = (
        f"{rtl_visual(ARABIC_APP_NAME)} | Daily Work Report v{APP_VERSION} | "
        f"{datetime.now().isoformat(timespec='minutes').replace('T', ' ')} | "
        f"{rtl_visual('آخر تحديث للبيانات')}: {analytics.get('freshness', '-')}"
    )
    pdf.drawString(1.0 * cm, 0.92 * cm, footer)
    pdf.save()
    if not output.is_file() or output.stat().st_size < 1000:
        raise RuntimeError("تعذر إنشاء التقرير الإداري.")
    return output


class LuxuryCalendar:
    def __init__(
        self,
        parent: tk.Misc,
        selected_date: date | None = None,
        on_select: Any = None,
    ):
        self.parent = parent
        self.on_select = on_select
        self.selected_date = selected_date or date.today()
        self.current_month = self.selected_date.replace(day=1)
        self.month_label: tk.Label | None = None
        self.days_frame: tk.Frame | None = None
        self._build()

    def _build(self) -> None:
        header = tk.Frame(self.parent, bg=Palette.NAVY, height=48)
        header.pack(fill="x", padx=6, pady=(6, 2))
        header.pack_propagate(False)

        tk.Button(
            header,
            text="◀",
            command=self.previous_month,
            bg=Palette.NAVY,
            fg=Palette.GOLD,
            activebackground=Palette.BLUE,
            activeforeground=Palette.WHITE,
            relief="flat",
            font=("Arial", 12, "bold"),
            width=3,
        ).pack(side="left", padx=3, pady=5)
        tk.Button(
            header,
            text="▶",
            command=self.next_month,
            bg=Palette.NAVY,
            fg=Palette.GOLD,
            activebackground=Palette.BLUE,
            activeforeground=Palette.WHITE,
            relief="flat",
            font=("Arial", 12, "bold"),
            width=3,
        ).pack(side="right", padx=3, pady=5)
        self.month_label = tk.Label(
            header,
            bg=Palette.NAVY,
            fg=Palette.GOLD,
            font=("Arial", 12, "bold"),
        )
        self.month_label.pack(expand=True, pady=9)

        weekdays = tk.Frame(self.parent, bg=Palette.BLUE)
        weekdays.pack(fill="x", padx=6)
        saturday_first = ("سبت", "أحد", "إثنين", "ثلاثاء", "أربعاء", "خميس", "جمعة")
        for column, day_name in enumerate(saturday_first):
            weekdays.grid_columnconfigure(column, weight=1, uniform="weekday")
            tk.Label(
                weekdays,
                text=day_name,
                bg=Palette.BLUE,
                fg=Palette.GOLD,
                font=("Arial", 9, "bold"),
                pady=5,
            ).grid(row=0, column=column, sticky="nsew")

        self.days_frame = tk.Frame(self.parent, bg=Palette.WHITE)
        self.days_frame.pack(fill="both", expand=True, padx=6)

        tk.Button(
            self.parent,
            text="اليوم",
            command=self.go_to_today,
            bg=Palette.PALE_BLUE,
            fg=Palette.NAVY,
            activebackground=Palette.BLUE,
            activeforeground=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            pady=5,
        ).pack(fill="x", padx=6, pady=(4, 6))
        self.render()

    def render(self) -> None:
        if self.days_frame is None or self.month_label is None:
            return
        for widget in self.days_frame.winfo_children():
            widget.destroy()
        self.month_label.config(
            text=f"{ARABIC_MONTHS[self.current_month.month - 1]} {self.current_month.year}"
        )

        month_rows = calendar.Calendar(firstweekday=calendar.SATURDAY).monthdayscalendar(
            self.current_month.year,
            self.current_month.month,
        )
        while len(month_rows) < 6:
            month_rows.append([0] * 7)

        today = date.today()
        for row_index, week in enumerate(month_rows):
            self.days_frame.grid_rowconfigure(row_index, weight=1, uniform="calendar_row")
            for column, day_number in enumerate(week):
                self.days_frame.grid_columnconfigure(column, weight=1, uniform="calendar_col")
                if not day_number:
                    tk.Label(self.days_frame, text="", bg=Palette.WHITE).grid(
                        row=row_index,
                        column=column,
                        sticky="nsew",
                        padx=1,
                        pady=1,
                    )
                    continue
                day_date = date(
                    self.current_month.year,
                    self.current_month.month,
                    day_number,
                )
                background = Palette.WHITE
                foreground = Palette.INK
                font = ("Arial", 10)
                if day_date == today:
                    background = Palette.PALE_BLUE
                    foreground = Palette.BLUE
                    font = ("Arial", 10, "bold")
                if day_date == self.selected_date:
                    background = Palette.ORANGE
                    foreground = Palette.WHITE
                    font = ("Arial", 10, "bold")
                tk.Button(
                    self.days_frame,
                    text=str(day_number),
                    command=lambda chosen=day_date: self.select_date(chosen),
                    bg=background,
                    fg=foreground,
                    activebackground=Palette.BLUE,
                    activeforeground=Palette.WHITE,
                    relief="flat",
                    font=font,
                    bd=0,
                ).grid(
                    row=row_index,
                    column=column,
                    sticky="nsew",
                    padx=1,
                    pady=1,
                )

    def select_date(self, chosen: date) -> None:
        self.selected_date = chosen
        self.current_month = chosen.replace(day=1)
        self.render()
        if callable(self.on_select):
            self.on_select(chosen)

    def set_date(self, chosen: date) -> None:
        self.select_date(chosen)

    def previous_month(self) -> None:
        year = self.current_month.year
        month = self.current_month.month
        if year <= 1 and month == 1:
            return
        self.current_month = (
            date(year - 1, 12, 1) if month == 1 else date(year, month - 1, 1)
        )
        self.render()

    def next_month(self) -> None:
        year = self.current_month.year
        month = self.current_month.month
        if year >= 9999 and month == 12:
            return
        self.current_month = (
            date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        )
        self.render()

    def go_to_today(self) -> None:
        self.select_date(date.today())


class ConfigurationManager(tk.Toplevel):
    def __init__(self, app: "KitchensERP"):
        super().__init__(app.root)
        self.app = app
        self.draft = normalize_config(copy.deepcopy(app.config))
        self._dragged_field_id: str | None = None
        self.category_by_label = {
            self.draft["categories"][category_id]["label"]: category_id
            for category_id in category_ids(self.draft)
        }
        self.title("تخصيص التقرير والأقسام")
        self.geometry("1080x740")
        self.minsize(940, 650)
        self.configure(bg=Palette.PALE)
        self.transient(app.root)
        self.grab_set()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Control-s>", lambda _event: self.apply_changes())

    @property
    def category_id(self) -> str:
        fallback = category_ids(self.draft)[0]
        return self.category_by_label.get(self.category_var.get(), fallback)

    @property
    def category(self) -> dict[str, Any]:
        return self.draft["categories"][self.category_id]

    def _build(self) -> None:
        tk.Label(
            self,
            text="تخصيص التقرير والأقسام",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 19, "bold"),
            pady=14,
        ).pack(fill="x")

        identity_bar = tk.Frame(self, bg=Palette.WHITE, bd=1, relief="solid")
        identity_bar.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(
            identity_bar,
            text="عنوان التقرير:",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(10, 4), pady=9)
        self.report_title_var = tk.StringVar(
            value=self.draft["identity"]["report_title"]
        )
        ttk.Entry(
            identity_bar,
            textvariable=self.report_title_var,
            justify="right",
            width=40,
        ).pack(side="right", padx=4, pady=8)
        tk.Label(
            identity_bar,
            text="غيّر اسم القسم ومسمى الشخص/الجهة والقائمة والحقول؛ ويظهر التغيير مباشرة في شاشة الإدخال والتقرير.",
            bg=Palette.WHITE,
            fg=Palette.MUTED,
            font=("Arial", 9),
        ).pack(side="left", padx=10)

        category_bar = tk.Frame(self, bg=Palette.PALE)
        category_bar.pack(fill="x", padx=16, pady=8)
        tk.Label(
            category_bar,
            text="الفئة:",
            bg=Palette.PALE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
        ).pack(side="right", padx=6)
        first_category_id = category_ids(self.draft)[0]
        self.category_var = tk.StringVar(
            value=self.draft["categories"][first_category_id]["label"]
        )
        self.category_combo = ttk.Combobox(
            category_bar,
            textvariable=self.category_var,
            values=list(self.category_by_label),
            state="readonly",
            width=24,
            justify="right",
        )
        self.category_combo.pack(side="right")
        self.category_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh())
        self._small_button(category_bar, "إضافة قسم", self.add_category, Palette.GREEN)
        self._small_button(
            category_bar,
            "تخصيص القسم",
            self.edit_category_labels,
            Palette.BLUE,
        )
        self._small_button(category_bar, "حذف القسم", self.delete_category, Palette.RED)
        self._small_button(
            category_bar,
            "←",
            lambda: self.move_category(1),
            Palette.NAVY,
        )
        self._small_button(
            category_bar,
            "→",
            lambda: self.move_category(-1),
            Palette.NAVY,
        )
        tk.Button(
            category_bar,
            text="استعادة إعدادات الفئة",
            command=self.restore_category,
            bg=Palette.PALE_BLUE,
            fg=Palette.NAVY,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=10,
            pady=5,
        ).pack(side="left")

        body = tk.Frame(self, bg=Palette.PALE)
        body.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        self.people_box = tk.LabelFrame(
            body,
            text="القائمة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=10,
            pady=10,
        )
        people_box = self.people_box
        people_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        people_list_area = tk.Frame(people_box, bg=Palette.WHITE)
        people_list_area.pack(fill="both", expand=True)
        self.people_list = tk.Listbox(
            people_list_area,
            font=("Arial", 11),
            justify="right",
            selectmode="browse",
            activestyle="none",
            bd=1,
            relief="solid",
        )
        people_scroll = ttk.Scrollbar(
            people_list_area,
            orient="vertical",
            command=self.people_list.yview,
        )
        self.people_list.configure(yscrollcommand=people_scroll.set)
        self.people_list.pack(side="left", fill="both", expand=True)
        people_scroll.pack(side="right", fill="y")

        people_buttons = tk.Frame(people_box, bg=Palette.WHITE)
        people_buttons.pack(fill="x", pady=(8, 0))
        self._small_button(people_buttons, "إضافة", self.add_person, Palette.GREEN)
        self._small_button(people_buttons, "تعديل", self.rename_person, Palette.BLUE)
        self._small_button(people_buttons, "حذف", self.delete_person, Palette.RED)
        self._small_button(people_buttons, "↑", lambda: self.move_person(-1), Palette.NAVY)
        self._small_button(people_buttons, "↓", lambda: self.move_person(1), Palette.NAVY)

        fields_box = tk.LabelFrame(
            body,
            text="الحقول والأعمدة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=10,
            pady=10,
        )
        fields_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        fields_box.grid_columnconfigure(0, weight=1)
        fields_box.grid_rowconfigure(0, weight=1)

        self.fields_tree = ttk.Treeview(
            fields_box,
            columns=("label", "kind", "options"),
            show="headings",
            height=12,
            selectmode="browse",
        )
        self.fields_tree.heading("label", text="العنوان")
        self.fields_tree.heading("kind", text="النوع")
        self.fields_tree.heading("options", text="خيارات القائمة")
        self.fields_tree.column("label", width=210, anchor="e")
        self.fields_tree.column("kind", width=110, anchor="center")
        self.fields_tree.column("options", width=220, anchor="e")
        fields_scroll = ttk.Scrollbar(
            fields_box,
            orient="vertical",
            command=self.fields_tree.yview,
        )
        self.fields_tree.configure(yscrollcommand=fields_scroll.set)
        self.fields_tree.grid(row=0, column=0, sticky="nsew")
        fields_scroll.grid(row=0, column=1, sticky="ns")
        self.fields_tree.bind("<<TreeviewSelect>>", self.on_field_select)

        order_box = tk.Frame(fields_box, bg=Palette.PALE_BLUE, bd=1, relief="solid")
        order_box.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(9, 0))
        tk.Label(
            order_box,
            text="ترتيب الأعمدة في التقرير — اسحب أي عنوان يمينًا أو يسارًا:",
            bg=Palette.PALE_BLUE,
            fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
            anchor="e",
        ).pack(fill="x", padx=7, pady=(6, 2))
        self.column_order_strip = tk.Frame(order_box, bg=Palette.PALE_BLUE)
        self.column_order_strip.pack(fill="x", padx=7, pady=(2, 4))
        self.column_order_hint_var = tk.StringVar(
            value="يُعرض أول حقل في أقصى اليمين، ويُحفظ الترتيب في كل التقارير والتصديرات."
        )
        tk.Label(
            order_box,
            textvariable=self.column_order_hint_var,
            bg=Palette.PALE_BLUE,
            fg=Palette.MUTED,
            font=("Arial", 8),
            anchor="e",
        ).pack(fill="x", padx=7, pady=(0, 5))
        self.bind("<ButtonRelease-1>", self._finish_column_drag, add="+")

        edit_frame = tk.Frame(fields_box, bg=Palette.WHITE)
        edit_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(9, 0))
        edit_frame.grid_columnconfigure(1, weight=1)
        tk.Label(
            edit_frame,
            text="العنوان:",
            bg=Palette.WHITE,
            fg=Palette.INK,
            font=("Arial", 10, "bold"),
        ).grid(row=0, column=2, padx=5, pady=4, sticky="e")
        self.field_label_var = tk.StringVar()
        ttk.Entry(
            edit_frame,
            textvariable=self.field_label_var,
            justify="right",
        ).grid(row=0, column=1, padx=5, pady=4, sticky="ew")
        self.field_kind_var = tk.StringVar(value=FIELD_KIND_LABELS["number"])
        ttk.Combobox(
            edit_frame,
            textvariable=self.field_kind_var,
            values=tuple(FIELD_KIND_LABELS.values()),
            state="readonly",
            width=18,
            justify="right",
        ).grid(row=0, column=0, padx=5, pady=4)
        tk.Label(
            edit_frame,
            text="خيارات القائمة (افصل بفاصلة):",
            bg=Palette.WHITE,
            fg=Palette.INK,
            font=("Arial", 9, "bold"),
        ).grid(row=1, column=2, padx=5, pady=4, sticky="e")
        self.field_options_var = tk.StringVar()
        ttk.Entry(
            edit_frame,
            textvariable=self.field_options_var,
            justify="right",
        ).grid(row=1, column=0, columnspan=2, padx=5, pady=4, sticky="ew")

        field_buttons = tk.Frame(fields_box, bg=Palette.WHITE)
        field_buttons.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self._small_button(field_buttons, "إضافة حقل", self.add_field, Palette.GREEN)
        self._small_button(field_buttons, "حفظ التعديل", self.update_field, Palette.BLUE)
        self._small_button(field_buttons, "حذف الحقل", self.delete_field, Palette.RED)
        self._small_button(field_buttons, "↑", lambda: self.move_field(-1), Palette.NAVY)
        self._small_button(field_buttons, "↓", lambda: self.move_field(1), Palette.NAVY)
        self._small_button(
            field_buttons,
            "الترتيب الافتراضي",
            self.restore_field_order,
            Palette.MUTED,
        )

        footer = tk.Frame(self, bg=Palette.PALE)
        footer.pack(fill="x", padx=16, pady=(0, 14))
        tk.Button(
            footer,
            text="حفظ التغييرات",
            command=self.apply_changes,
            bg=Palette.GREEN,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=8,
        ).pack(side="right", padx=4)
        tk.Button(
            footer,
            text="إلغاء",
            command=self.destroy,
            bg=Palette.MUTED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=8,
        ).pack(side="right", padx=4)
        self.refresh()

    def _small_button(
        self,
        parent: tk.Misc,
        text: str,
        command: Any,
        color: str,
    ) -> None:
        tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=9,
            pady=5,
        ).pack(side="right", padx=2)

    def refresh_category_selector(self, selected_id: str | None = None) -> None:
        ordered = category_ids(self.draft)
        self.category_by_label = {
            self.draft["categories"][category_id]["label"]: category_id
            for category_id in ordered
        }
        self.category_combo.configure(values=list(self.category_by_label))
        target_id = selected_id if selected_id in ordered else ordered[0]
        self.category_var.set(self.draft["categories"][target_id]["label"])
        self.refresh()

    def add_category(self) -> None:
        value = simpledialog.askstring(
            "إضافة قسم",
            "اكتب اسم القسم الجديد، مثل: المتابعة أو المبيعات:",
            parent=self,
        )
        if value is None:
            return
        label = value.strip()
        if not label:
            messagebox.showwarning("تنبيه", "لا يمكن إضافة قسم بلا اسم.", parent=self)
            return
        if label.casefold() in {
            category["label"].casefold()
            for category in self.draft["categories"].values()
        }:
            messagebox.showwarning("تنبيه", "اسم القسم موجود مسبقًا.", parent=self)
            return
        category_id = f"category_{uuid.uuid4().hex[:12]}"
        self.draft["categories"][category_id] = {
            "label": label,
            "person_label": "الاسم",
            "total_label": "الإجمالي",
            "people": [],
            "fields": [],
        }
        self.draft["category_order"].append(category_id)
        self.refresh_category_selector(category_id)

    def edit_category_labels(self) -> None:
        category = self.category
        label = simpledialog.askstring(
            "اسم القسم",
            "اسم القسم الظاهر في الشاشة والتقرير:",
            initialvalue=category["label"],
            parent=self,
        )
        if label is None:
            return
        clean_label = label.strip()
        if not clean_label:
            messagebox.showwarning("تنبيه", "اسم القسم مطلوب.", parent=self)
            return
        if clean_label.casefold() in {
            item["label"].casefold()
            for category_id, item in self.draft["categories"].items()
            if category_id != self.category_id
        }:
            messagebox.showwarning("تنبيه", "اسم القسم موجود مسبقًا.", parent=self)
            return
        person_label = simpledialog.askstring(
            "مسمى الاسم",
            "مسمى الشخص/الجهة، مثل: الرسام، الفني، الموظف، العميل:",
            initialvalue=category.get("person_label", "الاسم"),
            parent=self,
        )
        if person_label is None:
            return
        total_label = simpledialog.askstring(
            "مسمى الإجمالي",
            "مسمى عمود الإجمالي:",
            initialvalue=category.get("total_label", "الإجمالي"),
            parent=self,
        )
        if total_label is None:
            return
        selected_id = self.category_id
        category["label"] = clean_label
        category["person_label"] = person_label.strip() or "الاسم"
        category["total_label"] = total_label.strip() or "الإجمالي"
        self.refresh_category_selector(selected_id)

    def delete_category(self) -> None:
        ordered = category_ids(self.draft)
        if len(ordered) <= 1:
            messagebox.showwarning(
                "تنبيه",
                "يجب الإبقاء على قسم واحد على الأقل.",
                parent=self,
            )
            return
        selected_id = self.category_id
        label = self.category["label"]
        if not messagebox.askyesno(
            "تأكيد حذف القسم",
            f"حذف قسم «{label}» مع أسمائه وحقوله؟\n"
            "ستُحذف إدخالاته من التقرير الحالي عند حفظ الإعدادات.",
            parent=self,
        ):
            return
        position = ordered.index(selected_id)
        del self.draft["categories"][selected_id]
        self.draft["category_order"] = [
            category_id
            for category_id in self.draft["category_order"]
            if category_id != selected_id
        ]
        remaining = category_ids(self.draft)
        self.refresh_category_selector(remaining[min(position, len(remaining) - 1)])

    def move_category(self, direction: int) -> None:
        selected_id = self.category_id
        ordered = category_ids(self.draft)
        index = ordered.index(selected_id)
        destination = index + direction
        if destination < 0 or destination >= len(ordered):
            return
        ordered[index], ordered[destination] = ordered[destination], ordered[index]
        self.draft["category_order"] = ordered
        self.refresh_category_selector(selected_id)

    def refresh(self, selected_field_id: str | None = None) -> None:
        self.people_box.configure(
            text=f"قائمة {self.category.get('person_label', 'الأسماء')}"
        )
        self.people_list.delete(0, "end")
        for person in self.category["people"]:
            self.people_list.insert("end", person)
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for field in self.category["fields"]:
            kind_label = FIELD_KIND_LABELS.get(field["kind"], FIELD_KIND_LABELS["number"])
            self.fields_tree.insert(
                "",
                "end",
                iid=field["id"],
                values=(
                    field["label"],
                    kind_label,
                    "، ".join(field.get("options", [])),
                ),
            )
        self._render_column_order_strip()
        self.field_label_var.set("")
        self.field_kind_var.set(FIELD_KIND_LABELS["number"])
        self.field_options_var.set("")
        if selected_field_id and self.fields_tree.exists(selected_field_id):
            self.fields_tree.selection_set(selected_field_id)
            self.fields_tree.focus(selected_field_id)
            self.fields_tree.see(selected_field_id)
            self.on_field_select()

    def _render_column_order_strip(self) -> None:
        for widget in self.column_order_strip.winfo_children():
            widget.destroy()
        fields = self.category["fields"]
        if not fields:
            tk.Label(
                self.column_order_strip,
                text="لا توجد أعمدة قابلة للترتيب.",
                bg=Palette.PALE_BLUE,
                fg=Palette.MUTED,
                font=("Arial", 9),
            ).pack(side="right", padx=4, pady=4)
            return
        for field in fields:
            chip = tk.Label(
                self.column_order_strip,
                text=field["label"],
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                activebackground=Palette.GOLD,
                bd=1,
                relief="raised",
                font=("Arial", 9, "bold"),
                padx=9,
                pady=5,
                cursor="hand2",
            )
            setattr(chip, "field_id", field["id"])
            chip.bind(
                "<ButtonPress-1>",
                lambda _event, field_id=field["id"]: self._begin_column_drag(field_id),
            )
            chip.pack(side="right", padx=3, pady=3)

    def _begin_column_drag(self, field_id: str) -> None:
        self._dragged_field_id = field_id
        self.configure(cursor="fleur")
        self.column_order_hint_var.set(
            "استمر بالسحب ثم أفلت فوق عنوان آخر لاختيار موضع العمود الجديد."
        )

    def _finish_column_drag(self, event: Any) -> None:
        source_id = self._dragged_field_id
        if not source_id:
            return
        self._dragged_field_id = None
        self.configure(cursor="")
        target_widget = self.winfo_containing(event.x_root, event.y_root)
        while target_widget is not None and not getattr(target_widget, "field_id", None):
            if target_widget is self:
                target_widget = None
                break
            target_widget = getattr(target_widget, "master", None)
        target_id = str(getattr(target_widget, "field_id", ""))
        if not target_id or target_id == source_id:
            self.column_order_hint_var.set(
                "لم يتغير الترتيب. اسحب العنوان وأفلته فوق عمود آخر."
            )
            return
        local_x = event.x_root - target_widget.winfo_rootx()
        place_after = local_x < (target_widget.winfo_width() / 2)
        if move_field_by_drop(
            self.category["fields"],
            source_id,
            target_id,
            place_after=place_after,
        ):
            self.column_order_hint_var.set(
                "تم تحديث الترتيب من اليمين إلى اليسار؛ اضغط «حفظ التغييرات» لاعتماده."
            )
            self.refresh(source_id)

    def restore_field_order(self) -> None:
        if self.category_id not in BUILTIN_FIELD_ORDER:
            messagebox.showinfo(
                "قسم مخصص",
                "لا يوجد ترتيب افتراضي لهذا القسم المخصص.",
                parent=self,
            )
            return
        selected_id = self.selected_field_id()
        self.category["fields"] = canonical_builtin_fields(
            self.category_id,
            self.category["fields"],
            restore_missing=True,
        )
        self.column_order_hint_var.set(
            "تمت استعادة ترتيب الأعمدة القياسي مع إبقاء الحقول المخصصة."
        )
        self.refresh(selected_id)

    def selected_person_index(self) -> int | None:
        selection = self.people_list.curselection()
        return int(selection[0]) if selection else None

    def add_person(self) -> None:
        value = simpledialog.askstring(
            "إضافة اسم",
            f"أدخل اسمًا جديدًا ضمن {self.category['label']}:",
            parent=self,
        )
        if value is None:
            return
        name = value.strip()
        if not name:
            messagebox.showwarning("تنبيه", "لا يمكن إضافة اسم فارغ.", parent=self)
            return
        if name.casefold() in {person.casefold() for person in self.category["people"]}:
            messagebox.showwarning("تنبيه", "الاسم موجود مسبقًا.", parent=self)
            return
        self.category["people"].append(name)
        self.refresh()

    def rename_person(self) -> None:
        index = self.selected_person_index()
        if index is None:
            messagebox.showwarning("تنبيه", "اختر اسمًا أولًا.", parent=self)
            return
        old_name = self.category["people"][index]
        value = simpledialog.askstring(
            "تعديل الاسم",
            "الاسم الجديد:",
            initialvalue=old_name,
            parent=self,
        )
        if value is None:
            return
        new_name = value.strip()
        if not new_name:
            messagebox.showwarning("تنبيه", "لا يمكن حفظ اسم فارغ.", parent=self)
            return
        duplicates = {
            name.casefold()
            for person_index, name in enumerate(self.category["people"])
            if person_index != index
        }
        if new_name.casefold() in duplicates:
            messagebox.showwarning("تنبيه", "الاسم موجود مسبقًا.", parent=self)
            return
        self.category["people"][index] = new_name
        self.refresh()

    def delete_person(self) -> None:
        index = self.selected_person_index()
        if index is None:
            messagebox.showwarning("تنبيه", "اختر اسمًا أولًا.", parent=self)
            return
        name = self.category["people"][index]
        confirmed = messagebox.askyesno(
            "تأكيد الحذف",
            f"حذف «{name}» من القائمة؟\nلن تُحذف السجلات القديمة المرتبطة به.",
            parent=self,
        )
        if confirmed:
            self.category["people"].pop(index)
            self.refresh()

    def move_person(self, direction: int) -> None:
        index = self.selected_person_index()
        if index is None:
            return
        people = self.category["people"]
        destination = index + direction
        if destination < 0 or destination >= len(people):
            return
        people[index], people[destination] = people[destination], people[index]
        self.refresh()
        self.people_list.selection_set(destination)
        self.people_list.activate(destination)
        self.people_list.see(destination)

    def selected_field_id(self) -> str | None:
        selection = self.fields_tree.selection()
        return selection[0] if selection else None

    def on_field_select(self, _event: Any = None) -> None:
        field_id = self.selected_field_id()
        if not field_id:
            return
        field = next(
            (item for item in self.category["fields"] if item["id"] == field_id),
            None,
        )
        if field:
            self.field_label_var.set(field["label"])
            self.field_kind_var.set(FIELD_KIND_LABELS.get(field["kind"], FIELD_KIND_LABELS["number"]))
            self.field_options_var.set("، ".join(field.get("options", [])))

    def _edited_field_values(self) -> tuple[str, str, list[str]] | None:
        label = self.field_label_var.get().strip()
        if not label:
            messagebox.showwarning("تنبيه", "أدخل عنوان الحقل.", parent=self)
            return None
        kind = FIELD_KIND_BY_LABEL.get(self.field_kind_var.get(), "number")
        options = unique_clean_strings(
            [
                item
                for item in re.split(r"[,،;\n]+", self.field_options_var.get())
                if item.strip()
            ]
        )
        if kind == "choice" and not options:
            messagebox.showwarning(
                "تنبيه",
                "أدخل خيارًا واحدًا على الأقل لحقل قائمة الاختيار.",
                parent=self,
            )
            return None
        if kind != "choice":
            options = []
        return label, kind, options

    def add_field(self) -> None:
        edited = self._edited_field_values()
        if not edited:
            return
        label, kind, options = edited
        if label.casefold() in {
            field["label"].casefold() for field in self.category["fields"]
        }:
            messagebox.showwarning("تنبيه", "عنوان الحقل موجود مسبقًا.", parent=self)
            return
        self.category["fields"].append(
            {
                "id": f"field_{uuid.uuid4().hex[:12]}",
                "label": label,
                "kind": kind,
                "options": options,
            }
        )
        self.refresh()

    def update_field(self) -> None:
        field_id = self.selected_field_id()
        if not field_id:
            messagebox.showwarning("تنبيه", "اختر حقلًا أولًا.", parent=self)
            return
        edited = self._edited_field_values()
        if not edited:
            return
        label, kind, options = edited
        if label.casefold() in {
            field["label"].casefold()
            for field in self.category["fields"]
            if field["id"] != field_id
        }:
            messagebox.showwarning("تنبيه", "عنوان الحقل موجود مسبقًا.", parent=self)
            return
        for field in self.category["fields"]:
            if field["id"] == field_id:
                field["label"] = label
                field["kind"] = kind
                field["options"] = options
                break
        self.refresh()

    def delete_field(self) -> None:
        field_id = self.selected_field_id()
        if not field_id:
            messagebox.showwarning("تنبيه", "اختر حقلًا أولًا.", parent=self)
            return
        field = next(
            item for item in self.category["fields"] if item["id"] == field_id
        )
        confirmed = messagebox.askyesno(
            "تأكيد حذف الحقل",
            f"حذف عمود «{field['label']}»؟\nستُزال قيمه من الإدخالات الحالية.",
            parent=self,
        )
        if confirmed:
            self.category["fields"] = [
                item for item in self.category["fields"] if item["id"] != field_id
            ]
            self.refresh()

    def move_field(self, direction: int) -> None:
        field_id = self.selected_field_id()
        if not field_id:
            return
        fields = self.category["fields"]
        index = next(
            (position for position, field in enumerate(fields) if field["id"] == field_id),
            None,
        )
        if index is None:
            return
        destination = index + direction
        if destination < 0 or destination >= len(fields):
            return
        fields[index], fields[destination] = fields[destination], fields[index]
        self.refresh()
        self.fields_tree.selection_set(field_id)
        self.fields_tree.focus(field_id)

    def restore_category(self) -> None:
        if self.category_id not in DEFAULT_CONFIG["categories"]:
            messagebox.showinfo(
                "قسم مخصص",
                "هذا قسم مخصص ولا توجد له إعدادات افتراضية.",
                parent=self,
            )
            return
        confirmed = messagebox.askyesno(
            "استعادة الإعدادات",
            "استعادة الأسماء والحقول الافتراضية لهذه الفئة؟",
            parent=self,
        )
        if confirmed:
            self.draft["categories"][self.category_id] = copy.deepcopy(
                DEFAULT_CONFIG["categories"][self.category_id]
            )
            self.refresh()

    def apply_changes(self) -> None:
        self.draft["identity"]["report_title"] = (
            self.report_title_var.get().strip() or APP_NAME
        )
        if self.app.apply_configuration(self.draft) is not False:
            self.destroy()


class PeopleManager(tk.Toplevel):
    """إدارة مرنة للأشخاص/العمال/الفنيين داخل أي قسم."""

    def __init__(self, app: "KitchensERP", category_id: str | None = None):
        super().__init__(app.root)
        self.app = app
        self.title("إدارة الأشخاص والقوائم")
        self.geometry("760x560")
        self.minsize(650, 500)
        self.configure(bg=Palette.PALE)
        self.transient(app.root)
        self.grab_set()
        ordered = category_ids(app.config)
        self.category_id = category_id if category_id in ordered else ordered[0]
        self._build()
        self.refresh_categories()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())

    @property
    def category(self) -> dict[str, Any]:
        return self.app.config["categories"][self.category_id]

    def _build(self) -> None:
        header = tk.Frame(self, bg=Palette.NAVY, height=76)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="إدارة الأشخاص والقوائم", bg=Palette.NAVY, fg=Palette.WHITE,
                 font=("Arial", 19, "bold")).pack(pady=(11, 0))
        tk.Label(header, text="أضف أو عدّل أو احذف أو رتّب أي موظف / عامل / فني / اسم مخصص",
                 bg=Palette.NAVY, fg=Palette.ACCENT_LIGHT,
                 font=("Arial", 9, "bold")).pack()

        top = tk.Frame(self, bg=Palette.WHITE, bd=1, relief="solid")
        top.pack(fill="x", padx=16, pady=12)
        tk.Label(top, text="القسم", bg=Palette.WHITE, fg=Palette.NAVY,
                 font=("Arial", 10, "bold")).pack(side="right", padx=(8, 4), pady=10)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(top, textvariable=self.category_var, state="readonly",
                                           justify="right", width=32)
        self.category_combo.pack(side="right", padx=4, pady=8)
        self.category_combo.bind("<<ComboboxSelected>>", self.on_category_change)
        self.person_label_var = tk.StringVar()
        tk.Label(top, textvariable=self.person_label_var, bg=Palette.PALE_BLUE, fg=Palette.NAVY,
                 font=("Arial", 10, "bold"), padx=12, pady=6).pack(side="left", padx=8)

        body = tk.Frame(self, bg=Palette.WHITE, bd=1, relief="solid")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        list_frame = tk.Frame(body, bg=Palette.WHITE)
        list_frame.pack(fill="both", expand=True, padx=12, pady=12)
        self.people_list = tk.Listbox(list_frame, font=("Arial", 12), justify="right",
                                      selectmode="browse", activestyle="none", bd=1, relief="solid")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.people_list.yview)
        self.people_list.configure(yscrollcommand=scroll.set)
        self.people_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.people_list.bind("<Double-1>", lambda _event: self.rename_person())
        self.people_list.bind("<Delete>", lambda _event: self.delete_person())

        buttons = tk.Frame(body, bg=Palette.WHITE)
        buttons.pack(fill="x", padx=12, pady=(0, 12))
        self._button(buttons, "+ إضافة", self.add_person, Palette.GREEN)
        self._button(buttons, "تعديل", self.rename_person, Palette.BLUE)
        self._button(buttons, "حذف", self.delete_person, Palette.RED)
        self._button(buttons, "↑ أعلى", lambda: self.move_person(-1), Palette.NAVY)
        self._button(buttons, "↓ أسفل", lambda: self.move_person(1), Palette.INFO)
        self._button(buttons, "إغلاق", self.destroy, Palette.MUTED, side="left")

    def _button(self, parent: tk.Misc, text: str, command: Any, color: str, side: str = "right") -> None:
        button = tk.Button(parent, text=text, command=command, bg=color, fg=Palette.WHITE,
                           activebackground=Palette.PRIMARY_DARK, activeforeground=Palette.WHITE,
                           relief="flat", bd=0, font=("Arial", 10, "bold"), padx=14, pady=8,
                           cursor="hand2")
        button.pack(side=side, padx=3)
        button.bind("<Enter>", lambda _event, item=button: item.configure(relief="raised"))
        button.bind("<Leave>", lambda _event, item=button: item.configure(relief="flat"))

    def refresh_categories(self) -> None:
        ordered = category_ids(self.app.config)
        self.label_to_id = {self.app.config["categories"][category_id]["label"]: category_id
                            for category_id in ordered}
        self.category_combo.configure(values=list(self.label_to_id))
        if self.category_id not in ordered:
            self.category_id = ordered[0]
        self.category_var.set(self.app.config["categories"][self.category_id]["label"])
        self.refresh_people()

    def on_category_change(self, _event: Any = None) -> None:
        self.category_id = self.label_to_id.get(self.category_var.get(), self.category_id)
        self.refresh_people()

    def refresh_people(self, select_index: int | None = None) -> None:
        self.person_label_var.set(f"القائمة: {self.category.get('person_label', 'الأسماء')}")
        self.people_list.delete(0, "end")
        for name in self.category.get("people", []):
            self.people_list.insert("end", name)
        people = self.category.get("people", [])
        if select_index is not None and 0 <= select_index < len(people):
            self.people_list.selection_set(select_index)
            self.people_list.activate(select_index)
            self.people_list.see(select_index)

    def selected_index(self) -> int | None:
        selection = self.people_list.curselection()
        return int(selection[0]) if selection else None

    def _commit(self) -> None:
        self.app.save_config()
        self.app._refresh_name_values()
        self.app._schedule_draft_save()

    def add_person(self) -> None:
        value = simpledialog.askstring("إضافة", f"أدخل اسمًا جديدًا ضمن {self.category['label']}:", parent=self)
        if value is None:
            return
        name = value.strip()
        if not name:
            messagebox.showwarning("تنبيه", "لا يمكن إضافة اسم فارغ.", parent=self)
            return
        people = self.category.setdefault("people", [])
        if name.casefold() in {person.casefold() for person in people}:
            messagebox.showwarning("تنبيه", "الاسم موجود مسبقًا.", parent=self)
            return
        people.append(name)
        self._commit()
        self.refresh_people(len(people) - 1)

    def rename_person(self) -> None:
        index = self.selected_index()
        if index is None:
            messagebox.showwarning("تنبيه", "اختر اسمًا أولًا.", parent=self)
            return
        people = self.category["people"]
        value = simpledialog.askstring("تعديل", "الاسم الجديد:", initialvalue=people[index], parent=self)
        if value is None:
            return
        new_name = value.strip()
        if not new_name:
            messagebox.showwarning("تنبيه", "لا يمكن حفظ اسم فارغ.", parent=self)
            return
        if new_name.casefold() in {person.casefold() for i, person in enumerate(people) if i != index}:
            messagebox.showwarning("تنبيه", "الاسم موجود مسبقًا.", parent=self)
            return
        people[index] = new_name
        self._commit()
        self.refresh_people(index)

    def delete_person(self) -> None:
        index = self.selected_index()
        if index is None:
            messagebox.showwarning("تنبيه", "اختر اسمًا أولًا.", parent=self)
            return
        people = self.category["people"]
        name = people[index]
        confirmed = messagebox.askyesno(
            "تأكيد الحذف",
            f"حذف «{name}» من القائمة؟\nلن تُحذف السجلات القديمة المرتبطة بهذا الاسم.",
            parent=self,
        )
        if not confirmed:
            return
        people.pop(index)
        self._commit()
        self.refresh_people(min(index, len(people) - 1) if people else None)

    def move_person(self, direction: int) -> None:
        index = self.selected_index()
        if index is None:
            return
        people = self.category["people"]
        destination = index + direction
        if destination < 0 or destination >= len(people):
            return
        people[index], people[destination] = people[destination], people[index]
        self._commit()
        self.refresh_people(destination)


class StructureManager(tk.Toplevel):
    """إدارة الأعمدة والقوائم المنسدلة لكل قسم مع الحفظ الفوري."""

    KIND_LABELS = FIELD_KIND_LABELS

    def __init__(self, app: "KitchensERP", category_id: str | None = None):
        super().__init__(app.root)
        self.app = app
        self.draft = normalize_config(copy.deepcopy(app.config))
        ordered = category_ids(self.draft)
        self.category_id = category_id if category_id in ordered else ordered[0]
        self.title("إدارة القوائم والأعمدة")
        self.geometry("1000x650")
        self.minsize(860, 580)
        self.configure(bg=Palette.PALE)
        self.transient(app.root)
        self.grab_set()
        self._build()
        self.refresh_categories()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())

    @property
    def category(self) -> dict[str, Any]:
        return self.draft["categories"][self.category_id]

    def _button(self, parent: tk.Misc, text: str, command: Any, color: str, *, side: str = "right") -> tk.Button:
        button = tk.Button(
            parent, text=text, command=command, bg=color, fg=Palette.WHITE,
            activebackground=Palette.PRIMARY_DARK, activeforeground=Palette.WHITE,
            relief="flat", bd=0, font=("Arial", 9, "bold"), padx=12, pady=7,
            cursor="hand2",
        )
        button.pack(side=side, padx=3, pady=3)
        return button

    def _build(self) -> None:
        header = tk.Frame(self, bg=Palette.NAVY, height=82)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header, text="إدارة القوائم والأعمدة", bg=Palette.NAVY, fg=Palette.WHITE,
            font=("Arial", 19, "bold"),
        ).pack(pady=(11, 0))
        tk.Label(
            header,
            text="أضف أو احذف أو أعد تسمية وترتيب الأعمدة وخيارات القوائم حسب حاجة كل قسم",
            bg=Palette.NAVY, fg=Palette.ACCENT_LIGHT, font=("Arial", 9, "bold"),
        ).pack()

        top = tk.Frame(self, bg=Palette.WHITE, bd=1, relief="solid")
        top.pack(fill="x", padx=16, pady=12)
        tk.Label(top, text="القسم", bg=Palette.WHITE, fg=Palette.NAVY,
                 font=("Arial", 10, "bold")).pack(side="right", padx=(8, 4), pady=10)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(
            top, textvariable=self.category_var, state="readonly", justify="right", width=30,
        )
        self.category_combo.pack(side="right", padx=4, pady=8)
        self.category_combo.bind("<<ComboboxSelected>>", self.on_category_change)
        self.summary_var = tk.StringVar()
        tk.Label(top, textvariable=self.summary_var, bg=Palette.PALE_BLUE, fg=Palette.NAVY,
                 font=("Arial", 9, "bold"), padx=10, pady=6).pack(side="left", padx=8)

        body = tk.Frame(self, bg=Palette.PALE)
        body.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        options_box = tk.LabelFrame(
            body, text="خيارات القائمة المنسدلة", bg=Palette.WHITE, fg=Palette.NAVY,
            font=("Arial", 11, "bold"), padx=9, pady=9,
        )
        options_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        list_area = tk.Frame(options_box, bg=Palette.WHITE)
        list_area.pack(fill="both", expand=True)
        self.options_list = tk.Listbox(
            list_area, font=("Arial", 11), justify="right", selectmode="browse",
            activestyle="none", bd=1, relief="solid",
        )
        options_scroll = ttk.Scrollbar(list_area, orient="vertical", command=self.options_list.yview)
        self.options_list.configure(yscrollcommand=options_scroll.set)
        self.options_list.pack(side="left", fill="both", expand=True)
        options_scroll.pack(side="right", fill="y")
        self.options_list.bind("<Double-1>", lambda _event: self.rename_option())
        self.options_list.bind("<Delete>", lambda _event: self.delete_option())
        option_actions = tk.Frame(options_box, bg=Palette.WHITE)
        option_actions.pack(fill="x", pady=(8, 0))
        self._button(option_actions, "+ خيار", self.add_option, Palette.GREEN)
        self._button(option_actions, "تعديل", self.rename_option, Palette.BLUE)
        self._button(option_actions, "حذف", self.delete_option, Palette.RED)
        self._button(option_actions, "↑", lambda: self.move_option(-1), Palette.NAVY)
        self._button(option_actions, "↓", lambda: self.move_option(1), Palette.INFO)
        self.option_hint_var = tk.StringVar(value="اختر عمودًا من نوع «قائمة منسدلة» لإدارة خياراته.")
        tk.Label(options_box, textvariable=self.option_hint_var, bg=Palette.WHITE, fg=Palette.MUTED,
                 justify="right", wraplength=300, font=("Arial", 9)).pack(fill="x", pady=(8, 0))

        fields_box = tk.LabelFrame(
            body, text="الأعمدة / الحقول", bg=Palette.WHITE, fg=Palette.NAVY,
            font=("Arial", 11, "bold"), padx=9, pady=9,
        )
        fields_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        fields_box.grid_columnconfigure(0, weight=1)
        fields_box.grid_rowconfigure(0, weight=1)
        self.fields_tree = ttk.Treeview(
            fields_box, columns=("label", "kind", "count"), show="headings",
            selectmode="browse", height=14,
        )
        self.fields_tree.heading("label", text="اسم العمود")
        self.fields_tree.heading("kind", text="النوع")
        self.fields_tree.heading("count", text="عدد الخيارات")
        self.fields_tree.column("label", width=280, anchor="e")
        self.fields_tree.column("kind", width=130, anchor="center")
        self.fields_tree.column("count", width=90, anchor="center")
        field_scroll = ttk.Scrollbar(fields_box, orient="vertical", command=self.fields_tree.yview)
        self.fields_tree.configure(yscrollcommand=field_scroll.set)
        self.fields_tree.grid(row=0, column=0, sticky="nsew")
        field_scroll.grid(row=0, column=1, sticky="ns")
        self.fields_tree.bind("<<TreeviewSelect>>", lambda _event: self.refresh_options())
        self.fields_tree.bind("<Double-1>", lambda _event: self.rename_field())
        self.fields_tree.bind("<Delete>", lambda _event: self.delete_field())

        field_actions = tk.Frame(fields_box, bg=Palette.WHITE)
        field_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        add_menu = tk.Menubutton(
            field_actions, text="+ إضافة عمود ▾", bg=Palette.GREEN, fg=Palette.WHITE,
            activebackground=Palette.PRIMARY_DARK, activeforeground=Palette.WHITE,
            relief="flat", bd=0, font=("Arial", 9, "bold"), padx=12, pady=7,
            cursor="hand2",
        )
        menu = tk.Menu(add_menu, tearoff=False)
        for kind, label in FIELD_KIND_LABELS.items():
            menu.add_command(label=label, command=lambda selected=kind: self.add_field(selected))
        add_menu.configure(menu=menu)
        add_menu.pack(side="right", padx=3, pady=3)
        self._button(field_actions, "تعديل الاسم", self.rename_field, Palette.BLUE)
        self._button(field_actions, "إلزامي / اختياري", self.toggle_required, Palette.ACCENT)
        self._button(field_actions, "حذف العمود", self.delete_field, Palette.RED)
        self._button(field_actions, "↑ أعلى", lambda: self.move_field(-1), Palette.NAVY)
        self._button(field_actions, "↓ أسفل", lambda: self.move_field(1), Palette.INFO)
        self._button(field_actions, "التخصيص المتقدم", self.open_advanced, Palette.ACCENT, side="left")
        self._button(field_actions, "إغلاق", self.destroy, Palette.MUTED, side="left")

    def refresh_categories(self) -> None:
        ordered = category_ids(self.draft)
        self.label_to_id = {
            self.draft["categories"][category_id]["label"]: category_id for category_id in ordered
        }
        self.category_combo.configure(values=list(self.label_to_id))
        if self.category_id not in ordered:
            self.category_id = ordered[0]
        self.category_var.set(self.draft["categories"][self.category_id]["label"])
        self.refresh_fields()

    def on_category_change(self, _event: Any = None) -> None:
        self.category_id = self.label_to_id.get(self.category_var.get(), self.category_id)
        self.refresh_fields()

    def selected_field(self) -> dict[str, Any] | None:
        selection = self.fields_tree.selection()
        if not selection:
            return None
        field_id = selection[0]
        return next((field for field in self.category.get("fields", []) if field["id"] == field_id), None)

    def selected_field_index(self) -> int | None:
        field = self.selected_field()
        if field is None:
            return None
        return next((i for i, item in enumerate(self.category["fields"]) if item["id"] == field["id"]), None)

    def refresh_fields(self, select_id: str | None = None) -> None:
        self.fields_tree.delete(*self.fields_tree.get_children())
        fields = self.category.get("fields", [])
        for field in fields:
            options = field.get("options", []) if field.get("kind") == "choice" else []
            self.fields_tree.insert(
                "", "end", iid=field["id"],
                values=(("★ " if field.get("required") else "") + field["label"], self.KIND_LABELS.get(field["kind"], field["kind"]), len(options)),
            )
        self.summary_var.set(f"{len(fields)} عمودًا — {sum(1 for f in fields if f.get('kind') == 'choice')} قائمة")
        if select_id and self.fields_tree.exists(select_id):
            self.fields_tree.selection_set(select_id)
            self.fields_tree.focus(select_id)
            self.fields_tree.see(select_id)
        elif fields:
            self.fields_tree.selection_set(fields[0]["id"])
            self.fields_tree.focus(fields[0]["id"])
        self.refresh_options()

    def refresh_options(self, select_index: int | None = None) -> None:
        self.options_list.delete(0, "end")
        field = self.selected_field()
        enabled = bool(field and field.get("kind") == "choice")
        self.options_list.configure(state="normal" if enabled else "disabled")
        if not enabled:
            self.option_hint_var.set("اختر عمودًا من نوع «قائمة منسدلة» لإدارة خياراته.")
            return
        options = field.setdefault("options", [])
        for option in options:
            self.options_list.insert("end", option)
        self.option_hint_var.set(f"القائمة: {field['label']} — {len(options)} خيار")
        if select_index is not None and 0 <= select_index < len(options):
            self.options_list.selection_set(select_index)
            self.options_list.activate(select_index)
            self.options_list.see(select_index)

    def _commit(self, select_id: str | None = None) -> None:
        self.app.apply_configuration(copy.deepcopy(self.draft))
        self.draft = normalize_config(copy.deepcopy(self.app.config))
        self.refresh_categories()
        if select_id:
            self.refresh_fields(select_id)

    def add_field(self, kind: str) -> None:
        label = simpledialog.askstring("إضافة عمود", "اسم العمود الجديد:", parent=self)
        if label is None:
            return
        label = label.strip()
        if not label:
            messagebox.showwarning("تنبيه", "اسم العمود لا يمكن أن يكون فارغًا.", parent=self)
            return
        if label.casefold() in {str(field.get("label", "")).casefold() for field in self.category.get("fields", [])}:
            messagebox.showwarning("تنبيه", "يوجد عمود بهذا الاسم مسبقًا.", parent=self)
            return
        options: list[str] = []
        if kind == "choice":
            raw = simpledialog.askstring(
                "خيارات القائمة", "أدخل الخيارات مفصولة بفاصلة (يمكن تعديلها لاحقًا):", parent=self,
            )
            if raw is None:
                return
            options = unique_clean_strings(item for item in re.split(r"[,،;\n]+", raw) if item.strip())
            if not options:
                messagebox.showwarning("تنبيه", "أدخل خيارًا واحدًا على الأقل.", parent=self)
                return
        field_id = f"field_{uuid.uuid4().hex[:12]}"
        self.category.setdefault("fields", []).append(
            {"id": field_id, "label": label, "kind": kind, "options": options}
        )
        self._commit(field_id)

    def rename_field(self) -> None:
        field = self.selected_field()
        if field is None:
            messagebox.showwarning("تنبيه", "اختر عمودًا أولًا.", parent=self)
            return
        value = simpledialog.askstring(
            "تعديل اسم العمود", "الاسم الجديد:", initialvalue=field["label"], parent=self,
        )
        if value is None:
            return
        label = value.strip()
        if not label:
            messagebox.showwarning("تنبيه", "اسم العمود لا يمكن أن يكون فارغًا.", parent=self)
            return
        duplicates = {
            str(item.get("label", "")).casefold() for item in self.category.get("fields", [])
            if item["id"] != field["id"]
        }
        if label.casefold() in duplicates:
            messagebox.showwarning("تنبيه", "يوجد عمود بهذا الاسم مسبقًا.", parent=self)
            return
        field["label"] = label
        self._commit(field["id"])

    def toggle_required(self) -> None:
        field = self.selected_field()
        if field is None:
            messagebox.showwarning("تنبيه", "اختر عمودًا أولًا.", parent=self)
            return
        field["required"] = not bool(field.get("required", False))
        self._commit(field["id"])

    def delete_field(self) -> None:
        field = self.selected_field()
        if field is None:
            messagebox.showwarning("تنبيه", "اختر عمودًا أولًا.", parent=self)
            return
        if not messagebox.askyesno(
            "تأكيد حذف العمود",
            f"حذف عمود «{field['label']}»؟\nسيتم حذف قيم هذا العمود من السجلات الحالية في شاشة العمل.",
            parent=self,
        ):
            return
        self.category["fields"] = [item for item in self.category.get("fields", []) if item["id"] != field["id"]]
        self._commit()

    def move_field(self, direction: int) -> None:
        index = self.selected_field_index()
        if index is None:
            return
        fields = self.category.get("fields", [])
        destination = index + direction
        if destination < 0 or destination >= len(fields):
            return
        fields[index], fields[destination] = fields[destination], fields[index]
        field_id = fields[destination]["id"]
        self._commit(field_id)

    def selected_option_index(self) -> int | None:
        selection = self.options_list.curselection()
        return int(selection[0]) if selection else None

    def _choice_field(self) -> dict[str, Any] | None:
        field = self.selected_field()
        return field if field and field.get("kind") == "choice" else None

    def add_option(self) -> None:
        field = self._choice_field()
        if field is None:
            messagebox.showwarning("تنبيه", "اختر قائمة منسدلة أولًا.", parent=self)
            return
        value = simpledialog.askstring("إضافة خيار", "الخيار الجديد:", parent=self)
        if value is None:
            return
        option = value.strip()
        if not option:
            return
        options = field.setdefault("options", [])
        if option.casefold() in {item.casefold() for item in options}:
            messagebox.showwarning("تنبيه", "الخيار موجود مسبقًا.", parent=self)
            return
        options.append(option)
        field_id = field["id"]
        self._commit(field_id)
        self.refresh_options(len(options) - 1)

    def rename_option(self) -> None:
        field = self._choice_field()
        index = self.selected_option_index()
        if field is None or index is None:
            messagebox.showwarning("تنبيه", "اختر خيارًا أولًا.", parent=self)
            return
        options = field["options"]
        value = simpledialog.askstring("تعديل الخيار", "النص الجديد:", initialvalue=options[index], parent=self)
        if value is None:
            return
        option = value.strip()
        if not option:
            return
        if option.casefold() in {item.casefold() for i, item in enumerate(options) if i != index}:
            messagebox.showwarning("تنبيه", "الخيار موجود مسبقًا.", parent=self)
            return
        options[index] = option
        field_id = field["id"]
        self._commit(field_id)
        self.refresh_options(index)

    def delete_option(self) -> None:
        field = self._choice_field()
        index = self.selected_option_index()
        if field is None or index is None:
            messagebox.showwarning("تنبيه", "اختر خيارًا أولًا.", parent=self)
            return
        options = field["options"]
        option = options[index]
        if len(options) <= 1:
            messagebox.showwarning(
                "تنبيه",
                "لا يمكن حذف آخر خيار من القائمة المنسدلة. أضف خيارًا جديدًا أولًا أو احذف العمود كاملًا.",
                parent=self,
            )
            return
        if not messagebox.askyesno("تأكيد الحذف", f"حذف خيار «{option}»؟", parent=self):
            return
        options.pop(index)
        field_id = field["id"]
        self._commit(field_id)
        self.refresh_options(min(index, len(options) - 1) if options else None)

    def move_option(self, direction: int) -> None:
        field = self._choice_field()
        index = self.selected_option_index()
        if field is None or index is None:
            return
        options = field["options"]
        destination = index + direction
        if destination < 0 or destination >= len(options):
            return
        options[index], options[destination] = options[destination], options[index]
        field_id = field["id"]
        self._commit(field_id)
        self.refresh_options(destination)

    def open_advanced(self) -> None:
        self.destroy()
        ConfigurationManager(self.app)


class ReportDesigner(tk.Toplevel):
    def __init__(self, app: "KitchensERP"):
        super().__init__(app.root)
        self.app = app
        self.draft = normalize_config(copy.deepcopy(app.config))
        self.title("مصمم التقرير")
        self.geometry("820x620")
        self.minsize(760, 560)
        self.configure(bg=Palette.PALE)
        self.transient(app.root)
        self.grab_set()
        self.category_map = {self.draft["categories"][cid]["label"]: cid for cid in category_ids(self.draft)}
        self.category_var = tk.StringVar(value=next(iter(self.category_map)))
        self.visible_var = tk.BooleanVar(value=True)
        self.report_label_var = tk.StringVar()
        self.orientation_var = tk.StringVar(value="عمودي" if self.draft.get("report_layout", {}).get("orientation") == "portrait" else "أفقي")
        self.footer_var = tk.StringVar(value=str(self.draft.get("identity", {}).get("report_footer", "")))
        self._build()
        self.refresh()
        self.bind("<Escape>", lambda _e: self.destroy())

    @property
    def category_id(self) -> str:
        return self.category_map.get(self.category_var.get(), category_ids(self.draft)[0])

    @property
    def fields(self) -> list[dict[str, Any]]:
        return self.draft["categories"][self.category_id]["fields"]

    def _build(self) -> None:
        tk.Label(self, text="مصمم التقرير", bg=Palette.NAVY, fg=Palette.WHITE, font=("Arial", 19, "bold"), pady=13).pack(fill="x")
        top=tk.Frame(self,bg=Palette.WHITE,padx=10,pady=10); top.pack(fill="x",padx=14,pady=12)
        tk.Label(top,text="القسم / الكيان",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",10,"bold")).pack(side="right",padx=6)
        combo=ttk.Combobox(top,textvariable=self.category_var,values=list(self.category_map),state="readonly",justify="right",width=30); combo.pack(side="right"); combo.bind("<<ComboboxSelected>>",lambda _e:self.refresh())
        tk.Label(top,text="يمكن إخفاء أعمدة من المخرج أو تغيير عنوانها في التقرير دون حذف بياناتها.",bg=Palette.WHITE,fg=Palette.MUTED,font=("Arial",9)).pack(side="left",padx=8)
        body=tk.Frame(self,bg=Palette.PALE); body.pack(fill="both",expand=True,padx=14,pady=(0,10)); body.grid_columnconfigure(0,weight=1); body.grid_columnconfigure(1,weight=1); body.grid_rowconfigure(0,weight=1)
        left=tk.LabelFrame(body,text="أعمدة التقرير",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",11,"bold"),padx=8,pady=8); left.grid(row=0,column=1,sticky="nsew",padx=(6,0))
        self.tree=ttk.Treeview(left,columns=("name","visible","report"),show="headings",selectmode="browse"); self.tree.heading("name",text="العمود"); self.tree.heading("visible",text="يظهر"); self.tree.heading("report",text="عنوان التقرير"); self.tree.column("name",width=190,anchor="e"); self.tree.column("visible",width=65,anchor="center"); self.tree.column("report",width=200,anchor="e"); self.tree.pack(fill="both",expand=True); self.tree.bind("<<TreeviewSelect>>",lambda _e:self.load_selected())
        actions=tk.Frame(left,bg=Palette.WHITE); actions.pack(fill="x",pady=(8,0))
        tk.Button(actions,text="↑",command=lambda:self.move(-1),bg=Palette.NAVY,fg=Palette.WHITE,relief="flat",padx=12,pady=6).pack(side="right",padx=2)
        tk.Button(actions,text="↓",command=lambda:self.move(1),bg=Palette.INFO,fg=Palette.WHITE,relief="flat",padx=12,pady=6).pack(side="right",padx=2)
        right=tk.LabelFrame(body,text="خصائص العمود",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",11,"bold"),padx=12,pady=12); right.grid(row=0,column=0,sticky="nsew",padx=(0,6))
        ttk.Checkbutton(right,text="إظهار هذا العمود في التقرير والتصدير",variable=self.visible_var).pack(anchor="e",pady=8)
        tk.Label(right,text="عنوان العمود في التقرير",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",10,"bold")).pack(anchor="e",pady=(12,4))
        ttk.Entry(right,textvariable=self.report_label_var,justify="right").pack(fill="x")
        tk.Button(right,text="تطبيق على العمود",command=self.apply_field,bg=Palette.BLUE,fg=Palette.WHITE,relief="flat",font=("Arial",10,"bold"),padx=14,pady=7).pack(anchor="e",pady=14)
        tk.Label(right,text="اتجاه صفحة PDF",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",10,"bold")).pack(anchor="e",pady=(8,4))
        ttk.Combobox(right,textvariable=self.orientation_var,values=("أفقي","عمودي"),state="readonly",justify="right").pack(fill="x")
        tk.Label(right,text="تذييل التقرير المخصص",bg=Palette.WHITE,fg=Palette.NAVY,font=("Arial",10,"bold")).pack(anchor="e",pady=(10,4))
        ttk.Entry(right,textvariable=self.footer_var,justify="right").pack(fill="x")
        tk.Label(right,text="الترتيب هنا يصبح ترتيب الأعمدة في القالب الحالي. يمكنك بعد ذلك حفظه من «إدارة القوالب» لإعادة استخدامه.",bg=Palette.WHITE,fg=Palette.MUTED,wraplength=320,justify="right",font=("Arial",9)).pack(fill="x",pady=10)
        footer=tk.Frame(self,bg=Palette.PALE); footer.pack(fill="x",padx=14,pady=(0,14))
        tk.Button(footer,text="حفظ المصمم",command=self.save,bg=Palette.GREEN,fg=Palette.WHITE,relief="flat",font=("Arial",11,"bold"),padx=18,pady=8).pack(side="right",padx=4)
        tk.Button(footer,text="إلغاء",command=self.destroy,bg=Palette.MUTED,fg=Palette.WHITE,relief="flat",font=("Arial",10,"bold"),padx=18,pady=8).pack(side="right",padx=4)

    def refresh(self, select_id: str | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        for field in self.fields:
            self.tree.insert("","end",iid=field["id"],values=(("★ " if field.get("required") else "") + field["label"],"نعم" if field.get("report_visible",True) else "لا",field.get("report_label") or field["label"]))
        target=select_id if select_id and self.tree.exists(select_id) else (self.fields[0]["id"] if self.fields else "")
        if target:
            self.tree.selection_set(target); self.tree.focus(target); self.load_selected()

    def selected_field(self) -> dict[str, Any] | None:
        sel=self.tree.selection(); return next((f for f in self.fields if f["id"]==sel[0]),None) if sel else None

    def load_selected(self) -> None:
        field=self.selected_field()
        if not field: return
        self.visible_var.set(bool(field.get("report_visible",True)))
        self.report_label_var.set(str(field.get("report_label") or field["label"]))

    def apply_field(self) -> None:
        field=self.selected_field()
        if not field: return
        label=self.report_label_var.get().strip()
        field["report_visible"]=bool(self.visible_var.get())
        field["report_label"]="" if label==field["label"] else label
        self.refresh(field["id"])

    def move(self, direction: int) -> None:
        field=self.selected_field()
        if not field: return
        index=next(i for i,f in enumerate(self.fields) if f["id"]==field["id"]); dest=index+direction
        if 0 <= dest < len(self.fields):
            self.fields[index],self.fields[dest]=self.fields[dest],self.fields[index]; self.refresh(field["id"])

    def save(self) -> None:
        self.apply_field()
        self.draft.setdefault("report_layout", {})["orientation"] = "portrait" if self.orientation_var.get() == "عمودي" else "landscape"
        self.draft.setdefault("identity", {})["report_footer"] = self.footer_var.get().strip()
        self.app.apply_configuration(self.draft)
        messagebox.showinfo("تم الحفظ","تم تطبيق تصميم التقرير على القالب الحالي.",parent=self)
        self.destroy()


class AppearanceManager(tk.Toplevel):
    def __init__(self, app: "KitchensERP"):
        super().__init__(app.root)
        self.app = app
        self.title("المظهر وكثافة الواجهة")
        self.geometry("520x330")
        self.resizable(False, False)
        self.configure(bg=Palette.PALE)
        self.transient(app.root)
        self.grab_set()
        current = app.config.get("app_state", {}).get("appearance", {})
        self.mode_var = tk.StringVar(value=APPEARANCE_MODES.get(str(current.get("mode", "light")), "فاتح"))
        self.density_var = tk.StringVar(value=DENSITY_LABELS.get(str(current.get("density", "comfortable")), "مريح"))
        self.font_var = tk.DoubleVar(value=float(current.get("font_scale", 1.0)))
        tk.Label(self, text="المظهر", bg=Palette.NAVY, fg=Palette.WHITE, font=("Arial", 18, "bold"), pady=12).pack(fill="x")
        body = tk.Frame(self, bg=Palette.WHITE, padx=18, pady=18)
        body.pack(fill="both", expand=True, padx=14, pady=14)
        for label, variable, values in (
            ("الوضع", self.mode_var, tuple(APPEARANCE_MODES.values())),
            ("كثافة الواجهة", self.density_var, tuple(DENSITY_LABELS.values())),
        ):
            row = tk.Frame(body, bg=Palette.WHITE)
            row.pack(fill="x", pady=7)
            tk.Label(row, text=label, bg=Palette.WHITE, fg=Palette.NAVY, font=("Arial", 10, "bold")).pack(side="right")
            ttk.Combobox(row, textvariable=variable, values=values, state="readonly", justify="right", width=20).pack(side="left")
        row = tk.Frame(body, bg=Palette.WHITE)
        row.pack(fill="x", pady=7)
        tk.Label(row, text="حجم الخط", bg=Palette.WHITE, fg=Palette.NAVY, font=("Arial", 10, "bold")).pack(side="right")
        ttk.Combobox(row, textvariable=self.font_var, values=(0.85, 1.0, 1.15, 1.3), state="readonly", width=20).pack(side="left")
        buttons = tk.Frame(body, bg=Palette.WHITE)
        buttons.pack(fill="x", pady=(18,0))
        tk.Button(buttons, text="حفظ وتطبيق", command=self.save, bg=Palette.GREEN, fg=Palette.WHITE, relief="flat", font=("Arial", 10, "bold"), padx=16, pady=7).pack(side="right", padx=4)
        tk.Button(buttons, text="إغلاق", command=self.destroy, bg=Palette.MUTED, fg=Palette.WHITE, relief="flat", font=("Arial", 10, "bold"), padx=16, pady=7).pack(side="right", padx=4)
        tk.Label(body, text="الوضع الداكن يُحفظ فورًا ويُطبّق بالكامل عند التشغيل التالي؛ حجم الخط وكثافة الجداول يطبقان مباشرة.", bg=Palette.WHITE, fg=Palette.MUTED, wraplength=450, justify="right", font=("Arial", 9)).pack(fill="x", pady=(12,0))

    def save(self) -> None:
        reverse_modes = {label: key for key, label in APPEARANCE_MODES.items()}
        reverse_density = {label: key for key, label in DENSITY_LABELS.items()}
        appearance = {
            "mode": reverse_modes.get(self.mode_var.get(), "light"),
            "density": reverse_density.get(self.density_var.get(), "comfortable"),
            "font_scale": float(self.font_var.get()),
        }
        self.app.config.setdefault("app_state", {})["appearance"] = appearance
        Palette.use_mode(appearance["mode"])
        self.app.save_config()
        self.app.apply_appearance_settings()
        messagebox.showinfo("تم الحفظ", "تم حفظ إعدادات المظهر. الوضع اللوني الكامل يظهر بعد إعادة تشغيل البرنامج.", parent=self)


class KitchensERP:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.data_dir = app_data_dir()
        self.config_path = self.data_dir / "config.json"
        self.history_path = self.data_dir / "history.json"
        self.database_path = self.data_dir / "reports.db"
        stored_config = read_json(self.config_path, DEFAULT_CONFIG)
        self.config = normalize_config(stored_config)
        Palette.use_mode(self.config.get("app_state", {}).get("appearance", {}).get("mode", "light"))
        if stored_config != self.config:
            try:
                atomic_write_json(self.config_path, self.config)
            except OSError:
                pass
        loaded_history = read_json(self.history_path, [])
        self.database = ReportDatabase(self.database_path)
        self.database.migrate_legacy_history(loaded_history)
        default_company_id, default_template_id = self.database.ensure_defaults(self.config)
        company_ids = {
            str(company["id"]) for company in self.database.list_companies()
        }
        template_ids = {
            str(template["id"]) for template in self.database.list_templates()
        }
        saved_company_id = str(self.config.get("app_state", {}).get("company_id", ""))
        saved_template_id = str(self.config.get("app_state", {}).get("template_id", ""))
        self.current_company_id = (
            saved_company_id if saved_company_id in company_ids else default_company_id
        )
        self.current_template_id = (
            saved_template_id if saved_template_id in template_ids else default_template_id
        )
        self.current_user = self.database.bootstrap_login() or self._prompt_for_login()
        draft_identity = str(
            self.current_user.get("id")
            or self.current_user.get("username")
            or default_user_name()
        )
        draft_key = hashlib.sha256(draft_identity.encode("utf-8")).hexdigest()[:16]
        self.draft_store = DraftStore(self.data_dir / "drafts" / f"draft_{draft_key}.json")
        self._draft_save_job: str | None = None
        self.backup_manager = BackupManager(
            self.data_dir,
            self.database,
            self.config_path,
        )
        self.last_auto_backup: Path | None = None
        try:
            self.last_auto_backup = self.backup_manager.automatic_backup()
        except (OSError, sqlite3.Error):
            self.last_auto_backup = None
        self.entries: list[dict[str, Any]] = []
        self._undo_stack: list[list[dict[str, Any]]] = []
        self._redo_stack: list[list[dict[str, Any]]] = []
        self._history_limit = 50
        self.selected_entry_id: str | None = None
        self.loaded_report_id: str | None = None
        self.loaded_report_locked = False
        self.selected_date = date.today()
        self.field_vars: dict[str, tk.StringVar] = {}
        self.category_id_by_label = self._category_label_map()
        self._reload_business_choices()

        self.root.title(
            f"{self.config['identity']['report_title']} - الإصدار {APP_VERSION}"
        )
        self.root.geometry("1420x900")
        self.root.minsize(1030, 690)
        self.root.configure(bg=Palette.PALE)
        icon_path = next(
            (
                candidate
                for candidate in (
                    bundled_resource_path("daily_work_reports_icon.ico"),
                    bundled_resource_path("أيقونة_التطبيق_13.0.ico"),
                    bundled_resource_path("أيقونة_التطبيق_12.0.ico"),
                    bundled_resource_path("أيقونة_التطبيق_11.0.ico"),
                    bundled_resource_path("أيقونة_التطبيق_10.0.ico"),
                    bundled_resource_path("أيقونة_التطبيق_9.0.ico"),
                )
                if candidate.is_file()
            ),
            None,
        )
        if icon_path is not None:
            try:
                self.root.iconbitmap(str(icon_path))
            except tk.TclError:
                pass
        try:
            self.root.state("zoomed")
        except tk.TclError:
            pass
        self._configure_styles()
        self._build_ui()
        self.apply_appearance_settings()
        self._install_keyboard_shortcuts()
        self._install_draft_tracking()
        self.root.after_idle(self._offer_draft_recovery)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _prompt_for_login(self) -> dict[str, Any]:
        for _attempt in range(3):
            username = simpledialog.askstring(
                "تسجيل الدخول",
                "اسم المستخدم:",
                parent=self.root,
            )
            if username is None:
                break
            password = simpledialog.askstring(
                "تسجيل الدخول",
                "كلمة المرور:",
                parent=self.root,
                show="●",
            )
            if password is None:
                break
            user = self.database.authenticate(username, password)
            if user:
                return user
            messagebox.showerror(
                "تعذر تسجيل الدخول",
                "اسم المستخدم أو كلمة المرور غير صحيحة.",
                parent=self.root,
            )
        self.root.destroy()
        raise RuntimeError("تم إلغاء تسجيل الدخول.")

    @property
    def actor_name(self) -> str:
        return str(
            self.current_user.get("display_name")
            or self.current_user.get("username")
            or default_user_name()
        )

    def has_permission(self, permission: str) -> bool:
        role = str(self.current_user.get("role", "viewer"))
        permissions = ROLE_PERMISSIONS.get(role, set())
        return role == "admin" or "*" in permissions or permission in permissions

    def require_permission(self, permission: str) -> bool:
        if self.has_permission(permission):
            return True
        messagebox.showwarning(
            "صلاحية غير كافية",
            "لا يملك حسابك صلاحية تنفيذ هذا الإجراء.\n"
            f"الدور الحالي: {ROLE_LABELS.get(str(self.current_user.get('role')), 'مشاهد')}",
            parent=self.root,
        )
        return False

    def _category_label_map(self) -> dict[str, str]:
        return {
            self.config["categories"][category_id]["label"]: category_id
            for category_id in category_ids(self.config)
        }

    def _reload_business_choices(self) -> None:
        self.company_records = self.database.list_companies(active_only=True)
        self.company_label_to_id: dict[str, str] = {}
        self.company_id_to_label: dict[str, str] = {}
        used_company_labels: set[str] = set()
        for company in self.company_records:
            label = str(company["name"])
            if label in used_company_labels:
                label = f"{label} ({str(company['id'])[:6]})"
            used_company_labels.add(label)
            self.company_label_to_id[label] = str(company["id"])
            self.company_id_to_label[str(company["id"])] = label
        self.template_records = self.database.list_templates()
        self.template_label_to_id: dict[str, str] = {}
        self.template_id_to_label: dict[str, str] = {}
        used_template_labels: set[str] = set()
        for template in self.template_records:
            label = str(template["name"])
            if label in used_template_labels:
                label = f"{label} ({str(template['id'])[:6]})"
            used_template_labels.add(label)
            self.template_label_to_id[label] = str(template["id"])
            self.template_id_to_label[str(template["id"])] = label
        if hasattr(self, "company_combo"):
            self.company_combo.configure(values=list(self.company_label_to_id))
            self.company_var.set(
                self.company_id_to_label.get(self.current_company_id, "")
            )
        if hasattr(self, "template_combo"):
            self.template_combo.configure(values=list(self.template_label_to_id))
            self.template_var.set(
                self.template_id_to_label.get(self.current_template_id, "")
            )

    def current_company(self) -> dict[str, Any]:
        return self.database.get_company(self.current_company_id) or {
            "id": self.current_company_id,
            "name": self.config["identity"].get("company_name", ""),
            "logo_path": "",
            "address": "",
            "phone": "",
            "tax_number": "",
        }

    def current_template(self) -> dict[str, Any] | None:
        return next(
            (
                template
                for template in getattr(self, "template_records", [])
                if str(template["id"]) == self.current_template_id
            ),
            None,
        )

    def on_company_selected(self, _event: Any = None) -> None:
        company_id = self.company_label_to_id.get(self.company_var.get())
        if not company_id:
            return
        self.current_company_id = company_id
        company = self.current_company()
        self.config["app_state"]["company_id"] = company_id
        self.config["identity"]["company_name"] = str(company.get("name", ""))
        self.header_company_label.config(
            text=str(company.get("name", "")) or "نظام تقارير مرن وقابل للتخصيص"
        )
        self.save_config()
        if hasattr(self, "database_status_var"):
            self.database_status_var.set(
                f"الشركة: {company.get('name', '')} | "
                f"{self.database.count(company_id)} تقرير"
            )
        self._schedule_draft_save()

    def on_template_selected(self, _event: Any = None) -> None:
        template_id = self.template_label_to_id.get(self.template_var.get())
        if not template_id or template_id == self.current_template_id:
            return
        template = next(
            (
                item
                for item in self.template_records
                if str(item["id"]) == template_id
            ),
            None,
        )
        if template is None:
            return
        if self.entries and not messagebox.askyesno(
            "تبديل القالب",
            "سيتم تطبيق مسميات وحقول القالب الجديد على شاشة الإدخال.\n"
            "قد تُحذف من الشاشة سجلات لا توجد فئاتها في القالب الجديد. هل تريد المتابعة؟",
            parent=self.root,
        ):
            self.template_var.set(
                self.template_id_to_label.get(self.current_template_id, "")
            )
            return
        self.current_template_id = template_id
        self.config["app_state"]["template_id"] = template_id
        self.apply_configuration(copy.deepcopy(template["config"]))
        self.database.audit(
            self.actor_name,
            "apply",
            "template",
            template_id,
            {"name": template.get("name", "")},
        )

    def apply_appearance_settings(self) -> None:
        appearance = self.config.get("app_state", {}).get("appearance", {})
        density = str(appearance.get("density", "comfortable"))
        scale = float(appearance.get("font_scale", 1.0))
        rowheight = {"compact": 26, "comfortable": 32, "spacious": 38}.get(density, 32)
        style = ttk.Style(self.root)
        style.configure("Treeview", font=("Arial", max(8, round(10 * scale))), rowheight=rowheight)
        style.configure("Treeview.Heading", font=("Arial", max(8, round(10 * scale)), "bold"))
        style.configure("TCombobox", font=("Arial", max(8, round(10 * scale))))
        if hasattr(self, "entries_tree"):
            self.entries_tree.configure(style="Treeview")

    def show_report_designer(self) -> None:
        if self.require_permission("manage"):
            ReportDesigner(self)

    def show_appearance_manager(self) -> None:
        if self.require_permission("manage"):
            AppearanceManager(self)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            "Treeview",
            font=("Arial", 10),
            rowheight=32,
            background=Palette.WHITE,
            fieldbackground=Palette.WHITE,
            foreground=Palette.INK,
        )
        style.configure(
            "Treeview.Heading",
            font=("Arial", 10, "bold"),
            background=Palette.NAVY,
            foreground=Palette.WHITE,
            padding=7,
        )
        style.map(
            "Treeview",
            background=[("selected", Palette.BLUE_2)],
            foreground=[("selected", Palette.WHITE)],
        )
        style.configure("TCombobox", padding=6, font=("Arial", 10))
        style.configure("TEntry", padding=7, font=("Arial", 10))
        style.configure("TButton", padding=(10, 6), font=("Arial", 9, "bold"))

    def _build_menu_bar(self) -> None:
        menu_bar = tk.Menu(self.root, tearoff=False)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="تقرير جديد", command=self.clear_report, accelerator="Ctrl+N")
        file_menu.add_command(label="حفظ التقرير", command=self.save_to_history, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="استيراد بيانات", command=self.import_entries)
        file_menu.add_command(label="قالب الاستيراد", command=self.export_import_template)
        file_menu.add_separator()
        file_menu.add_command(label="خروج", command=self.on_close, accelerator="Alt+F4")
        menu_bar.add_cascade(label="ملف", menu=file_menu)

        edit_menu = tk.Menu(menu_bar, tearoff=False)
        edit_menu.add_command(label="إدخال / تحديث السجل", command=self.add_or_update_entry, accelerator="Enter")
        edit_menu.add_command(label="تراجع فعلي", command=self.undo_entries, accelerator="Ctrl+Z")
        edit_menu.add_command(label="إعادة", command=self.redo_entries, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="تراجع عن الإدخال الحالي", command=self.clear_editor, accelerator="Esc")
        edit_menu.add_separator()
        edit_menu.add_command(label="حذف السجل المحدد", command=self.delete_selected_entry, accelerator="Delete")
        menu_bar.add_cascade(label="تحرير", menu=edit_menu)

        export_menu = tk.Menu(menu_bar, tearoff=False)
        export_menu.add_command(label="PDF", command=self.export_pdf)
        export_menu.add_command(label="صورة", command=self.export_images)
        export_menu.add_command(label="Excel", command=self.export_excel)
        export_menu.add_command(label="CSV", command=self.export_csv)
        menu_bar.add_cascade(label="تصدير", menu=export_menu)

        reports_menu = tk.Menu(menu_bar, tearoff=False)
        reports_menu.add_command(label="السجل والتقارير الدورية", command=self.show_history)
        reports_menu.add_command(label="لوحة المؤشرات", command=self.show_dashboard)
        reports_menu.add_command(label="المحذوفات / الأرشيف", command=self.show_trash)
        menu_bar.add_cascade(label="التقارير", menu=reports_menu)

        customize_menu = tk.Menu(menu_bar, tearoff=False)
        customize_menu.add_command(label="مركز التخصيص الجديد", command=self.open_structure_studio)
        customize_menu.add_command(label="تخصيص التقرير والأقسام", command=self.open_configuration_manager)
        customize_menu.add_command(label="مصمم التقرير", command=self.show_report_designer)
        customize_menu.add_command(label="إدارة القوالب", command=self.show_template_manager)
        customize_menu.add_command(label="المظهر وكثافة الواجهة", command=self.show_appearance_manager)
        customize_menu.add_command(label="إدارة الشركات", command=self.show_company_manager)
        menu_bar.add_cascade(label="تخصيص", menu=customize_menu)

        manage_menu = tk.Menu(menu_bar, tearoff=False)
        manage_menu.add_command(label="مركز إدارة الأقسام والموظفين والحقول", command=self.open_structure_studio)
        manage_menu.add_separator()
        manage_menu.add_command(label="إدارة الأشخاص", command=self.open_people_manager)
        manage_menu.add_command(label="إدارة القوائم والأعمدة", command=self.open_structure_manager)
        manage_menu.add_separator()
        manage_menu.add_command(label="المستخدمون وسجل التدقيق", command=self.show_administration)
        manage_menu.add_command(label="النسخ الاحتياطي والاستعادة", command=self.show_backup_manager)
        manage_menu.add_command(label="ملفات تعريف الاستيراد", command=self.show_import_profiles_manager)
        menu_bar.add_cascade(label="إدارة", menu=manage_menu)

        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="اختصارات لوحة المفاتيح", command=self.show_keyboard_shortcuts)
        menu_bar.add_cascade(label="مساعدة", menu=help_menu)

        self.main_menu = menu_bar
        self.root.configure(menu=menu_bar)

    def _dropdown_button(
        self,
        parent: tk.Misc,
        text: str,
        items: list[tuple[str, Any] | None],
        color: str,
    ) -> tk.Menubutton:
        button = tk.Menubutton(
            parent,
            text=f"{text} ▾",
            bg=color,
            fg=Palette.WHITE,
            activebackground=Palette.NAVY,
            activeforeground=Palette.GOLD,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=12,
            pady=7,
            cursor="hand2",
        )
        menu = tk.Menu(
            button,
            tearoff=False,
            font=("Arial", 10),
            activebackground=Palette.PRIMARY_LIGHT,
            activeforeground=Palette.NAVY,
            bd=0,
        )
        for item in items:
            if item is None:
                menu.add_separator()
            else:
                label, command = item
                menu.add_command(label=label, command=command)
        button.configure(menu=menu)
        button.pack(side="right", padx=3, pady=5)
        return button

    def open_modern_home(self) -> None:
        """Quick navigation window; all operations use the existing business logic."""
        ModernHome(self)

    def _build_compact_actions(self) -> None:
        """Two clear action levels, preserving the original command surface."""
        actions = tk.Frame(self.root, bg=Palette.PALE, bd=0)
        self.actions_frame = actions
        actions.pack(fill="x", padx=12, pady=(4, 6))
        primary = tk.Frame(actions, bg=Palette.WHITE, highlightthickness=1,
                           highlightbackground=Palette.BORDER)
        primary.pack(fill="x", pady=(0, 4))
        def action(label, callback, color):
            button = tk.Button(primary, text=label, command=callback, bg=color,
                               fg="#FFFFFF", activebackground=Palette.NAVY,
                               activeforeground="#FFFFFF", relief="flat", bd=0,
                               font=("Arial", 10, "bold"), padx=13, pady=10,
                               cursor="hand2")
            button.pack(side="right", padx=4, pady=8)
            return button
        action("＋ إدخال سجل", self.add_or_update_entry, Palette.GREEN)
        action("حفظ التقرير", self.save_to_history, Palette.PRIMARY)
        action("تقرير جديد", self.clear_report, Palette.INFO)
        action("إدارة الموظفين والحقول", self.open_structure_studio, Palette.NAVY)
        action("سجل التقارير", self.show_history, Palette.PRIMARY)
        action("المؤشرات", self.show_dashboard, Palette.INFO)
        action("الرئيسية", self.open_modern_home, Palette.ACCENT)

        secondary = tk.Frame(actions, bg=Palette.PRIMARY_LIGHT, highlightthickness=1,
                             highlightbackground=Palette.BORDER)
        secondary.pack(fill="x")
        self._dropdown_button(secondary, "ملف", [
            ("تقرير جديد", self.clear_report),
            ("حفظ التقرير", self.save_to_history),
            None,
            ("استيراد بيانات", self.import_entries),
            ("قالب الاستيراد", self.export_import_template),
            None,
            ("خروج", self.on_close)], Palette.PRIMARY_DARK)
        self._dropdown_button(secondary, "إدخال", [
            ("إضافة / تحديث السجل", self.add_or_update_entry),
            ("تحديث السجل المحدد", self.update_selected_entry),
            ("إلغاء التحديد", self.clear_editor),
            None,
            ("تراجع", self.undo_entries),
            ("إعادة", self.redo_entries),
            ("حذف السجل المحدد", self.delete_selected_entry)], Palette.BLUE)
        self._dropdown_button(secondary, "تصدير", [
            ("تصدير PDF", self.export_pdf),
            ("تصدير صورة", self.export_images),
            ("تصدير Excel", self.export_excel),
            ("تصدير CSV", self.export_csv)], Palette.RED)
        self._dropdown_button(secondary, "التقارير", [
            ("سجل التقارير", self.show_history),
            ("لوحة المؤشرات", self.show_dashboard),
            ("المحذوفات / الأرشيف", self.show_trash)], Palette.NAVY)
        self._dropdown_button(secondary, "تخصيص", [
            ("مركز الموظفين والأقسام والحقول", self.open_structure_studio),
            ("مصمم التقرير", self.show_report_designer),
            ("إدارة القوالب", self.show_template_manager),
            ("إدارة الشركات", self.show_company_manager),
            ("المظهر وحجم الخط", self.show_appearance_manager)], Palette.ACCENT)
        self._dropdown_button(secondary, "إدارة", [
            ("مركز الإدارة الموحد", self.open_structure_studio),
            ("المستخدمون والصلاحيات", self.show_administration),
            ("النسخ الاحتياطي والاستعادة", self.show_backup_manager),
            ("ملفات تعريف الاستيراد", self.show_import_profiles_manager),
            ("استيراد مرن", self.import_entries)], Palette.INFO)
        self.quality_var = tk.StringVar(value="جودة التقرير: 0/100")
        self.database_status_var = tk.StringVar(
            value=f"قاعدة البيانات: {self.database.count()} تقرير")
        status = tk.Frame(secondary, bg=Palette.PRIMARY_LIGHT)
        status.pack(side="left", fill="x", expand=True, padx=8)
        tk.Label(status, textvariable=self.database_status_var,
                 fg=Palette.PRIMARY_DARK, bg=Palette.PRIMARY_LIGHT,
                 font=("Arial", 9, "bold")).pack(side="left", padx=6, pady=9)
        tk.Label(status, textvariable=self.quality_var,
                 fg=Palette.PRIMARY_DARK, bg=Palette.PRIMARY_LIGHT,
                 font=("Arial", 9, "bold")).pack(side="left", padx=6, pady=9)

    def _install_keyboard_shortcuts(self) -> None:
        self.root.bind("<Return>", self._on_enter_shortcut, add="+")
        self.root.bind("<KP_Enter>", self._on_enter_shortcut, add="+")
        self.root.bind("<Escape>", self._on_escape_shortcut, add="+")
        self.root.bind("<Delete>", lambda _event: self.delete_selected_entry(), add="+")
        self.root.bind("<Control-s>", lambda _event: self.save_to_history(), add="+")
        self.root.bind("<Control-n>", lambda _event: self.clear_report(), add="+")
        self.root.bind("<Control-z>", lambda _event: self.undo_entries(), add="+")
        self.root.bind("<Control-y>", lambda _event: self.redo_entries(), add="+")
        self.root.bind("<F2>", lambda _event: self.update_selected_entry(), add="+")
        self.root.bind("<Control-f>", lambda _event: self.show_history(), add="+")

    def _on_enter_shortcut(self, event: Any) -> str | None:
        focus = self.root.focus_get()
        if isinstance(focus, (tk.Button, tk.Menubutton, ttk.Button)):
            return None
        self.add_or_update_entry()
        return "break"

    def _on_escape_shortcut(self, _event: Any) -> str:
        self.clear_editor()
        return "break"

    def show_keyboard_shortcuts(self) -> None:
        messagebox.showinfo(
            "اختصارات لوحة المفاتيح",
            "Enter  : إدخال سجل جديد أو تحديث السجل المحدد\n"
            "Esc    : إلغاء التحديد والتراجع عن الإدخال الحالي\n"
            "Delete : حذف السجل المحدد\n"
            "Ctrl+S : حفظ التقرير في قاعدة البيانات\n"
            "Ctrl+N : تقرير جديد\n"
            "Ctrl+Z : تراجع عن آخر تعديل للسجلات\n"
            "Ctrl+Y : إعادة آخر تعديل\n"
            "Ctrl+F : البحث في سجل التقارير\n"
            "F2     : تعديل السجل المحدد",
            parent=self.root,
        )

    def _build_ui(self) -> None:
        self._build_menu_bar()
        # Compact modern application header: keeps the report area large and
        # puts the navigation/action menus directly underneath it.
        header = tk.Frame(self.root, bg=Palette.NAVY, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_box = tk.Frame(header, bg=Palette.NAVY)
        title_box.pack(side="right", fill="y", padx=16)
        self.header_title_label = tk.Label(
            title_box,
            text=self.config["identity"]["report_title"],
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 19, "bold"),
        )
        self.header_title_label.pack(anchor="e", pady=(9, 0))
        self.header_company_label = tk.Label(
            title_box,
            text=self.config["identity"]["company_name"] or "إدارة تقارير العمل اليومية",
            bg=Palette.NAVY,
            fg=Palette.ACCENT_LIGHT,
            font=("Arial", 10, "bold"),
        )
        self.header_company_label.pack(anchor="e", pady=(1, 7))
        tk.Label(
            header,
            text=f"v{APP_VERSION}",
            bg=Palette.PRIMARY,
            fg=Palette.WHITE,
            font=("Arial", 9, "bold"),
            padx=12,
            pady=5,
        ).pack(side="left", padx=16)

        self._build_compact_actions()

        report_info = tk.Frame(self.root, bg=Palette.WHITE, bd=1, relief="solid")
        report_info.pack(fill="x", padx=12, pady=(10, 6))
        identity_row = tk.Frame(report_info, bg=Palette.WHITE)
        identity_row.pack(fill="x")
        tk.Label(
            identity_row,
            text="اسم الشركة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(6, 2), pady=8)
        self.company_var = tk.StringVar(
            value=self.company_id_to_label.get(self.current_company_id, "")
        )
        self.company_combo = ttk.Combobox(
            identity_row,
            textvariable=self.company_var,
            values=list(self.company_label_to_id),
            state="readonly",
            width=25,
            justify="right",
        )
        self.company_combo.pack(side="right", padx=4, pady=8)
        self.company_combo.bind("<<ComboboxSelected>>", self.on_company_selected)
        tk.Button(
            identity_row,
            text="إدارة الشركات",
            command=self.show_company_manager,
            bg=Palette.PRIMARY,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=8,
            pady=5,
        ).pack(side="right", padx=2)
        tk.Label(
            identity_row,
            text="قالب التقرير",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(14, 2), pady=8)
        self.template_var = tk.StringVar(
            value=self.template_id_to_label.get(self.current_template_id, "")
        )
        self.template_combo = ttk.Combobox(
            identity_row,
            textvariable=self.template_var,
            values=list(self.template_label_to_id),
            state="readonly",
            width=27,
            justify="right",
        )
        self.template_combo.pack(side="right", padx=4, pady=8)
        self.template_combo.bind("<<ComboboxSelected>>", self.on_template_selected)
        tk.Button(
            identity_row,
            text="إدارة القوالب",
            command=self.show_template_manager,
            bg=Palette.ACCENT,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=8,
            pady=5,
        ).pack(side="right", padx=2)
        tk.Label(
            identity_row,
            text="مستخدم البرنامج",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(18, 2), pady=8)
        self.user_var = tk.StringVar(
            value=self.actor_name
        )
        ttk.Entry(
            identity_row,
            textvariable=self.user_var,
            width=18,
            justify="right",
            state="readonly",
        ).pack(side="right", padx=4, pady=8)
        tk.Label(
            identity_row,
            text=ROLE_LABELS.get(str(self.current_user.get("role")), "مشاهد"),
            bg=Palette.WHITE,
            fg=Palette.MUTED,
            font=("Arial", 9),
        ).pack(side="left", padx=8)

        details_row = tk.Frame(report_info, bg=Palette.WHITE)
        details_row.pack(fill="x", pady=(0, 4))
        tk.Label(
            details_row,
            text="الرمز",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(6, 2), pady=5)
        self.code_var = tk.StringVar(
            value=self.config["identity"]["default_code"]
        )
        ttk.Entry(
            details_row,
            textvariable=self.code_var,
            width=12,
            justify="center",
        ).pack(side="right", padx=4, pady=5)
        tk.Label(
            details_row,
            text="حالة التقرير",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(18, 2), pady=5)
        self.status_var = tk.StringVar(value=REPORT_STATUSES["draft"])
        self.status_combo = ttk.Combobox(
            details_row,
            textvariable=self.status_var,
            values=list(REPORT_STATUSES.values()),
            state="readonly",
            width=12,
            justify="right",
        )
        self.status_combo.pack(side="right", padx=4, pady=5)
        tk.Label(
            details_row,
            text="رقم التقرير",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(18, 2), pady=5)
        self.page_var = tk.StringVar(
            value=self.config["identity"]["default_report_number"]
        )
        ttk.Entry(
            details_row,
            textvariable=self.page_var,
            width=8,
            justify="center",
        ).pack(side="right", padx=4, pady=5)
        self.date_summary_label = tk.Label(
            details_row,
            text="",
            bg=Palette.PALE_BLUE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=18,
            pady=7,
        )
        self.date_summary_label.pack(side="left", padx=8, pady=5)
        self._update_date_summary()

        main = tk.Frame(self.root, bg=Palette.PALE)
        main.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        calendar_box = tk.LabelFrame(
            main,
            text="التقويم",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            width=330,
            padx=3,
            pady=3,
        )
        calendar_box.pack(side="right", fill="y", padx=(6, 0))
        calendar_box.pack_propagate(False)
        self.calendar = LuxuryCalendar(
            calendar_box,
            self.selected_date,
            self.on_date_selected,
        )

        work_area = tk.Frame(main, bg=Palette.PALE)
        work_area.pack(side="left", fill="both", expand=True)
        self._build_entry_editor(work_area)
        self._build_entries_table(work_area)

    def _build_entry_editor(self, parent: tk.Misc) -> None:
        editor = tk.LabelFrame(
            parent,
            text="إدخال الأعمال",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=10,
            pady=8,
        )
        editor.pack(fill="x", pady=(0, 6))

        selector_row = tk.Frame(editor, bg=Palette.WHITE)
        selector_row.pack(fill="x")
        tk.Label(
            selector_row,
            text="الفئة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(8, 3))
        first_category_id = category_ids(self.config)[0]
        first_label = self.config["categories"][first_category_id]["label"]
        self.category_var = tk.StringVar(value=first_label)
        self.category_combo = ttk.Combobox(
            selector_row,
            textvariable=self.category_var,
            values=list(self.category_id_by_label),
            state="readonly",
            width=22,
            justify="right",
        )
        self.category_combo.pack(side="right", padx=3)
        self.category_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.on_category_changed(),
        )

        self.person_selector_label = tk.Label(
            selector_row,
            text=self.config["categories"][first_category_id].get("person_label", "الاسم/الجهة"),
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        )
        self.person_selector_label.pack(side="right", padx=(18, 3))
        self.name_var = tk.StringVar()
        self.name_combo = ttk.Combobox(
            selector_row,
            textvariable=self.name_var,
            state="readonly",
            width=26,
            justify="right",
        )
        self.name_combo.pack(side="right", padx=3)
        people_menu_button = tk.Menubutton(
            selector_row, text="الأسماء ▾", bg=Palette.GREEN, fg=Palette.WHITE,
            activebackground=Palette.PRIMARY_DARK, activeforeground=Palette.WHITE,
            relief="flat", bd=0, font=("Arial", 9, "bold"),
            padx=11, pady=6, cursor="hand2"
        )
        people_menu = tk.Menu(people_menu_button, tearoff=False)
        people_menu.add_command(label="إضافة اسم سريع", command=self.quick_add_person)
        people_menu.add_command(label="إدارة / تعديل / حذف / ترتيب", command=self.open_people_manager)
        people_menu.add_separator()
        people_menu.add_command(label="إدارة القوائم والأعمدة", command=self.open_structure_manager)
        people_menu_button.configure(menu=people_menu)
        people_menu_button.pack(side="right", padx=3)
        customize_button = tk.Menubutton(
            selector_row, text="تخصيص ▾", bg=Palette.NAVY, fg=Palette.WHITE,
            activebackground=Palette.PRIMARY, activeforeground=Palette.WHITE,
            relief="flat", bd=0, cursor="hand2", font=("Arial", 9, "bold"),
            padx=12, pady=6
        )
        customize_menu = tk.Menu(customize_button, tearoff=False)
        customize_menu.add_command(label="إدارة القوائم والأعمدة", command=self.open_structure_manager)
        customize_menu.add_command(label="تخصيص التقرير والأقسام", command=self.open_configuration_manager)
        customize_button.configure(menu=customize_menu)
        customize_button.pack(side="left", padx=3)

        self.dynamic_fields_frame = tk.Frame(editor, bg=Palette.WHITE)
        self.dynamic_fields_frame.pack(fill="x", pady=(10, 4))
        self.total_var = tk.StringVar(value="0")
        self._refresh_name_values()
        self._rebuild_dynamic_fields()

        editor_actions = tk.Frame(editor, bg=Palette.WHITE)
        editor_actions.pack(fill="x", pady=(5, 0))
        tk.Label(
            editor_actions,
            textvariable=self.total_var,
            bg=Palette.PALE_BLUE,
            fg=Palette.ORANGE,
            font=("Arial", 14, "bold"),
            width=10,
            pady=5,
        ).pack(side="right", padx=4)
        tk.Label(
            editor_actions,
            text="الإجمالي التلقائي",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=4)
        self._editor_button(editor_actions, "إضافة للسجل", self.add_or_update_entry, Palette.GREEN)
        self._editor_button(editor_actions, "تحديث المحدد", self.update_selected_entry, Palette.BLUE)
        self._editor_button(editor_actions, "إلغاء التحديد", self.clear_editor, Palette.MUTED)

    def open_structure_studio(self) -> None:
        if self.require_permission("manage"):
            StructureStudio(self)

    def open_configuration_manager(self) -> None:
        if self.require_permission("manage"):
            ConfigurationManager(self)

    def open_people_manager(self) -> None:
        if self.require_permission("edit"):
            category_id = self.current_category_id if hasattr(self, "category_var") else None
            PeopleManager(self, category_id)

    def open_structure_manager(self) -> None:
        if self.require_permission("manage"):
            category_id = self.current_category_id if hasattr(self, "category_var") else None
            StructureManager(self, category_id)

    def _build_entries_table(self, parent: tk.Misc) -> None:
        table_box = tk.LabelFrame(
            parent,
            text="سجل التقرير الحالي",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=8,
            pady=8,
        )
        table_box.pack(fill="both", expand=True)
        table_container = tk.Frame(table_box, bg=Palette.WHITE)
        table_container.pack(fill="both", expand=True)
        self.entries_tree = ttk.Treeview(
            table_container,
            columns=("number", "category", "name", "details", "total"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width, anchor in (
            ("number", "#", 55, "center"),
            ("category", "الفئة", 145, "e"),
            ("name", "الاسم", 180, "e"),
            ("details", "التفاصيل", 590, "e"),
            ("total", "الإجمالي", 95, "center"),
        ):
            self.entries_tree.heading(column, text=title)
            saved_widths = self.config.get("app_state", {}).get("entry_table_widths", {})
            saved_width = saved_widths.get(column, width) if isinstance(saved_widths, dict) else width
            self.entries_tree.column(column, width=int(saved_width), anchor=anchor)
        vertical = ttk.Scrollbar(
            table_container,
            orient="vertical",
            command=self.entries_tree.yview,
        )
        horizontal = ttk.Scrollbar(
            table_container,
            orient="horizontal",
            command=self.entries_tree.xview,
        )
        self.entries_tree.configure(
            yscrollcommand=vertical.set,
            xscrollcommand=horizontal.set,
        )
        self.entries_tree.grid(row=0, column=0, sticky="nsew")
        self.entries_tree.tag_configure("odd", background=Palette.WHITE)
        self.entries_tree.tag_configure("even", background=Palette.ALT_ROW)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        self.entries_tree.bind("<<TreeviewSelect>>", self.on_entry_selected)
        self.entries_tree.bind("<ButtonRelease-1>", self._save_entry_table_widths, add="+")

        row_actions = tk.Frame(table_box, bg=Palette.WHITE)
        row_actions.pack(fill="x", pady=(7, 0))
        tk.Button(
            row_actions,
            text="حذف السجل المحدد",
            command=self.delete_selected_entry,
            bg=Palette.RED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=10,
            pady=5,
        ).pack(side="right")

    def _save_entry_table_widths(self, _event: Any = None) -> None:
        if not hasattr(self, "entries_tree"):
            return
        widths = {
            column: int(self.entries_tree.column(column, "width"))
            for column in self.entries_tree["columns"]
        }
        if self.config.get("app_state", {}).get("entry_table_widths") != widths:
            self.config.setdefault("app_state", {})["entry_table_widths"] = widths
            self.save_config()

    def _action_button(
        self,
        parent: tk.Misc,
        text: str,
        command: Any,
        color: str,
    ) -> None:
        tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg=Palette.WHITE,
            activebackground=Palette.NAVY,
            activeforeground=Palette.GOLD,
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 10, "bold"),
            padx=13,
            pady=8,
        ).pack(side="right", padx=3, pady=5)

    def _editor_button(
        self,
        parent: tk.Misc,
        text: str,
        command: Any,
        color: str,
    ) -> None:
        tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=11,
            pady=6,
        ).pack(side="left", padx=3)

    @property
    def current_category_id(self) -> str:
        fallback = category_ids(self.config)[0]
        return self.category_id_by_label.get(self.category_var.get(), fallback)

    def on_date_selected(self, selected: date) -> None:
        self.selected_date = selected
        self._update_date_summary()
        self._schedule_draft_save()

    def _update_date_summary(self) -> None:
        self.date_summary_label.config(
            text=f"{self.selected_date.isoformat()}  |  {arabic_day_name(self.selected_date)}"
        )

    def _refresh_name_values(self) -> None:
        category = self.config["categories"][self.current_category_id]
        people = category["people"]
        if hasattr(self, "person_selector_label"):
            self.person_selector_label.configure(
                text=category.get("person_label", "الاسم/الجهة")
            )
        self.name_combo.configure(values=people)
        if self.name_var.get() not in people:
            self.name_var.set("")

    def on_category_changed(self) -> None:
        self.name_var.set("")
        self._refresh_name_values()
        self._rebuild_dynamic_fields()

    def _rebuild_dynamic_fields(self, values: dict[str, Any] | None = None) -> None:
        for widget in self.dynamic_fields_frame.winfo_children():
            widget.destroy()
        self.field_vars = {}
        values = values or {}
        fields = self.config["categories"][self.current_category_id]["fields"]
        if not fields:
            tk.Label(
                self.dynamic_fields_frame,
                text="لا توجد حقول لهذه الفئة. أضف حقولًا من شاشة الإدارة.",
                bg=Palette.PALE,
                fg=Palette.MUTED,
                font=("Arial", 10),
                pady=9,
            ).pack(fill="x")
            self.total_var.set("0")
            return

        columns = min(3, max(1, len(fields)))
        for column in range(columns):
            self.dynamic_fields_frame.grid_columnconfigure(column, weight=1, uniform="field")
        for index, field in enumerate(fields):
            field_box = tk.LabelFrame(
                self.dynamic_fields_frame,
                text=field["label"],
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                font=("Arial", 9, "bold"),
                padx=7,
                pady=5,
            )
            # True RTL layout: the first configured field appears at the
            # right edge, matching Arabic reading and data-entry order.
            visual_column = columns - 1 - (index % columns)
            field_box.grid(
                row=index // columns,
                column=visual_column,
                sticky="ew",
                padx=4,
                pady=3,
            )
            initial = values.get(field["id"], 0 if field["kind"] == "number" else "")
            variable = tk.StringVar(value=str(initial))
            self.field_vars[field["id"]] = variable
            kind = field.get("kind", "text")
            if kind in {"number", "percentage", "currency"}:
                # Direct keyboard entry: no +/- arrows. Selecting the current
                # value on focus lets the operator type the number immediately.
                validator = (
                    self.root.register(
                        lambda proposed, field_kind=kind: self._validate_numeric_text(
                            proposed, field_kind
                        )
                    ),
                    "%P",
                )
                entry = ttk.Entry(
                    field_box,
                    textvariable=variable,
                    justify="center",
                    width=16,
                    validate="key",
                    validatecommand=validator,
                )
                entry.pack(fill="x")
                entry.bind(
                    "<FocusIn>",
                    lambda _event, widget=entry: widget.after_idle(
                        lambda: widget.selection_range(0, tk.END)
                    ),
                )
                if kind == "number":
                    variable.trace_add("write", lambda *_args: self._update_editor_total())
            elif kind == "choice":
                ttk.Combobox(
                    field_box, textvariable=variable, values=field.get("options", []),
                    state="readonly", justify="right",
                ).pack(fill="x")
            elif kind == "boolean":
                ttk.Combobox(
                    field_box, textvariable=variable, values=("نعم", "لا"),
                    state="readonly", justify="right",
                ).pack(fill="x")
            else:
                hint = {"date": "YYYY-MM-DD", "time": "HH:MM", "phone": "05xxxxxxxx"}.get(kind, "")
                entry = ttk.Entry(field_box, textvariable=variable, justify="right")
                entry.pack(fill="x")
                if hint:
                    entry.insert(0, variable.get())
        self._update_editor_total()

    @staticmethod
    def _validate_numeric_text(proposed: str, kind: str = "number") -> bool:
        """Allow fast direct numeric entry while keeping stored values clean."""
        if proposed == "":
            return True
        if kind == "number":
            return proposed.isdigit()
        return re.fullmatch(r"\d*(?:\.\d{0,2})?", proposed) is not None

    def _update_editor_total(self) -> None:
        total = 0
        for field in self.config["categories"][self.current_category_id]["fields"]:
            if field["kind"] == "number":
                variable = self.field_vars.get(field["id"])
                total += coerce_nonnegative_int(variable.get() if variable else 0)
        self.total_var.set(str(total))

    def quick_add_person(self) -> None:
        if not self.require_permission("edit"):
            return
        category = self.config["categories"][self.current_category_id]
        value = simpledialog.askstring(
            "إضافة اسم",
            f"أدخل اسمًا جديدًا ضمن {category['label']}:",
            parent=self.root,
        )
        if value is None:
            return
        name = value.strip()
        if not name:
            messagebox.showwarning("تنبيه", "لا يمكن إضافة اسم فارغ.")
            return
        if name.casefold() in {person.casefold() for person in category["people"]}:
            messagebox.showwarning("تنبيه", "الاسم موجود مسبقًا.")
            return
        category["people"].append(name)
        self.save_config()
        self._refresh_name_values()
        self.name_var.set(name)

    def _editor_values(self) -> dict[str, Any] | None:
        values: dict[str, Any] = {}
        for field in self.config["categories"][self.current_category_id]["fields"]:
            variable = self.field_vars.get(field["id"])
            raw = variable.get().strip() if variable else ""
            if field.get("required") and not raw:
                messagebox.showwarning("بيانات ناقصة", f"الحقل «{field['label']}» إلزامي.")
                return None
            kind = field.get("kind", "text")
            if kind == "number":
                if raw and not re.fullmatch(r"\d+", raw):
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يكون رقمًا صحيحًا موجبًا.")
                    return None
                values[field["id"]] = coerce_nonnegative_int(raw)
            elif kind in {"percentage", "currency"}:
                try:
                    numeric = float(raw.replace(",", "")) if raw else 0.0
                except ValueError:
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يكون رقمًا.")
                    return None
                if numeric < 0 or (kind == "percentage" and numeric > 100):
                    limit = " بين 0 و100" if kind == "percentage" else " موجبًا"
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يكون{limit}.")
                    return None
                values[field["id"]] = round(numeric, 2)
            elif kind == "date" and raw:
                try:
                    date.fromisoformat(raw)
                except ValueError:
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يكون بتاريخ YYYY-MM-DD.")
                    return None
                values[field["id"]] = raw
            elif kind == "time" and raw:
                if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", raw):
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يكون وقتًا HH:MM.")
                    return None
                values[field["id"]] = raw
            elif kind == "phone":
                cleaned = re.sub(r"[\s()-]+", "", raw)
                if cleaned and not re.fullmatch(r"\+?\d{7,15}", cleaned):
                    messagebox.showwarning("قيمة غير صحيحة", f"الحقل «{field['label']}» يجب أن يحتوي رقم هاتف صالحًا.")
                    return None
                values[field["id"]] = cleaned
            elif kind == "boolean":
                values[field["id"]] = raw if raw in {"نعم", "لا"} else ""
            else:
                values[field["id"]] = raw
        return values

    def _push_undo_state(self) -> None:
        self._undo_stack.append(copy.deepcopy(self.entries))
        if len(self._undo_stack) > self._history_limit:
            del self._undo_stack[0]
        self._redo_stack.clear()

    def undo_entries(self) -> None:
        if not self._undo_stack:
            return
        self._redo_stack.append(copy.deepcopy(self.entries))
        self.entries = self._undo_stack.pop()
        self.clear_editor()
        self.refresh_entries_tree()
        self._schedule_draft_save()

    def redo_entries(self) -> None:
        if not self._redo_stack:
            return
        self._undo_stack.append(copy.deepcopy(self.entries))
        self.entries = self._redo_stack.pop()
        self.clear_editor()
        self.refresh_entries_tree()
        self._schedule_draft_save()

    def add_or_update_entry(self) -> None:
        if not self.require_permission("edit"):
            return
        if self.selected_entry_id:
            self.update_selected_entry()
            return
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اختر الاسم أولًا.")
            return
        values = self._editor_values()
        if values is None:
            return
        self._push_undo_state()
        self.entries.append(new_entry(self.current_category_id, name, values))
        self.refresh_entries_tree()
        self.clear_editor()
        self._schedule_draft_save()

    def update_selected_entry(self) -> None:
        if not self.require_permission("edit"):
            return
        if not self.selected_entry_id:
            messagebox.showwarning("تنبيه", "اختر سجلًا من الجدول لتحديثه.")
            return
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اختر الاسم أولًا.")
            return
        values = self._editor_values()
        if values is None:
            return
        entry = next(
            (item for item in self.entries if item["id"] == self.selected_entry_id),
            None,
        )
        if entry is None:
            self.clear_editor()
            return
        self._push_undo_state()
        # Keep values of fields hidden/removed from the active schema. They may
        # be recovered from the structure backup without losing draft records.
        if entry.get("category") == self.current_category_id:
            values = {**entry.get("values", {}), **values}
        entry["category"] = self.current_category_id
        entry["name"] = name
        entry["values"] = values
        self.refresh_entries_tree()
        self.clear_editor()
        self._schedule_draft_save()

    def clear_editor(self) -> None:
        self.selected_entry_id = None
        for selected in self.entries_tree.selection():
            self.entries_tree.selection_remove(selected)
        self.name_var.set("")
        self._refresh_name_values()
        self._rebuild_dynamic_fields()

    def refresh_entries_tree(self) -> None:
        for item in self.entries_tree.get_children():
            self.entries_tree.delete(item)
        for index, entry in enumerate(self.entries, 1):
            category = self.config["categories"].get(entry["category"])
            if not category:
                continue
            details: list[str] = []
            for field in category["fields"]:
                value = entry.get("values", {}).get(
                    field["id"],
                    0 if field["kind"] == "number" else "",
                )
                shown = str(value)
                if field["kind"] == "text" and not shown:
                    shown = "-"
                details.append(f"{field['label']}: {shown}")
            total = calculate_total(self.config, entry)
            self.entries_tree.insert(
                "",
                "end",
                iid=entry["id"],
                values=(
                    index,
                    category["label"],
                    entry["name"],
                    "  |  ".join(details),
                    total,
                ),
                tags=("even" if index % 2 == 0 else "odd",),
            )
        if hasattr(self, "quality_var"):
            quality = report_quality_score(self.current_snapshot())
            self.quality_var.set(
                f"جودة التقرير: {quality['score']}/100 - {quality['level']}"
            )

    def on_entry_selected(self, _event: Any = None) -> None:
        selection = self.entries_tree.selection()
        if not selection:
            return
        entry_id = selection[0]
        entry = next((item for item in self.entries if item["id"] == entry_id), None)
        if entry is None:
            return
        self.selected_entry_id = entry_id
        category_label = self.config["categories"][entry["category"]]["label"]
        self.category_var.set(category_label)
        self._refresh_name_values()
        self.name_var.set(entry["name"])
        self._rebuild_dynamic_fields(entry.get("values", {}))

    def delete_selected_entry(self) -> None:
        if not self.require_permission("delete"):
            return
        selection = self.entries_tree.selection()
        if not selection:
            messagebox.showwarning("تنبيه", "اختر سجلًا لحذفه.")
            return
        entry_id = selection[0]
        if messagebox.askyesno("تأكيد الحذف", "حذف السجل المحدد من التقرير الحالي؟"):
            self._push_undo_state()
            self.entries = [entry for entry in self.entries if entry["id"] != entry_id]
            self.refresh_entries_tree()
            self.clear_editor()
            self._schedule_draft_save()

    def apply_configuration(self, draft: dict[str, Any]) -> bool:
        conflicts = active_category_conflicts(self.config, draft, self.entries)
        if conflicts:
            details = "، ".join(
                f"{self.config['categories'][key]['label']} ({count} سجل)"
                for key, count in conflicts.items()
            )
            messagebox.showwarning(
                "حماية التقرير الحالي",
                "لا يمكن إزالة أقسام لها سجلات في التقرير الجاري: " + details +
                "\nاحفظ التقرير أو انقل السجلات قبل إزالة القسم.",
                parent=self.root,
            )
            return False
        previous_config = copy.deepcopy(self.config)
        live_company = str(self.current_company().get("name", "")).strip()
        live_user = self.actor_name
        live_code = self.code_var.get().strip()
        live_page = self.page_var.get().strip()
        self.config = normalize_config(draft)
        self.config["app_state"]["company_id"] = self.current_company_id
        self.config["app_state"]["template_id"] = self.current_template_id
        self.config["identity"]["company_name"] = live_company
        self.config["identity"]["user_name"] = live_user
        self.config["identity"]["default_code"] = live_code
        self.config["identity"]["default_report_number"] = live_page or "1"
        # Never destructively trim draft records or their values when editing
        # the list of fields. Removed fields simply disappear from the form;
        # their previously entered values remain in the draft record.
        if not self.save_config():
            self.config = previous_config
            return False
        self.category_id_by_label = self._category_label_map()
        self.category_combo.configure(values=list(self.category_id_by_label))
        if self.category_var.get() not in self.category_id_by_label:
            first_id = category_ids(self.config)[0]
            self.category_var.set(self.config["categories"][first_id]["label"])
        self.header_title_label.config(text=self.config["identity"]["report_title"])
        self.root.title(
            f"{self.config['identity']['report_title']} - الإصدار {APP_VERSION}"
        )
        self.header_company_label.config(
            text=live_company or "نظام تقارير مرن وقابل للتخصيص"
        )
        self._refresh_name_values()
        self._rebuild_dynamic_fields()
        self.refresh_entries_tree()
        self._schedule_draft_save()
        return True

    def save_config(self) -> bool:
        self._sync_identity_from_ui()
        try:
            atomic_write_json(self.config_path, self.config)
            return True
        except OSError as error:
            messagebox.showerror("خطأ في الحفظ", f"تعذر حفظ الإعدادات:\n{error}")
            return False

    def _sync_identity_from_ui(self) -> None:
        if not hasattr(self, "company_var"):
            return
        identity = self.config["identity"]
        identity["company_name"] = str(self.current_company().get("name", "")).strip()
        identity["user_name"] = self.actor_name
        identity["default_code"] = self.code_var.get().strip()
        identity["default_report_number"] = self.page_var.get().strip() or "1"
        self.config["app_state"]["company_id"] = self.current_company_id
        self.config["app_state"]["template_id"] = self.current_template_id
        self.header_company_label.config(
            text=identity["company_name"] or "نظام تقارير مرن وقابل للتخصيص"
        )

    def current_snapshot(self) -> dict[str, Any]:
        self._sync_identity_from_ui()
        template = self.current_template() or {}
        status_code = next(
            (
                key
                for key, label in REPORT_STATUSES.items()
                if label == self.status_var.get()
            ),
            "draft",
        )
        return build_report_snapshot(
            self.config,
            self.entries,
            self.selected_date,
            self.code_var.get(),
            self.page_var.get(),
            str(self.current_company().get("name", "")),
            self.actor_name,
            company_profile=self.current_company(),
            company_id=self.current_company_id,
            template_id=self.current_template_id,
            template_name=str(template.get("name", "")),
            status=status_code,
        )

    def _install_draft_tracking(self) -> None:
        for variable in (self.code_var, self.page_var, self.status_var):
            variable.trace_add("write", lambda *_args: self._schedule_draft_save())

    def _schedule_draft_save(self) -> None:
        if not hasattr(self, "draft_store"):
            return
        try:
            if self._draft_save_job is not None:
                self.root.after_cancel(self._draft_save_job)
            self._draft_save_job = self.root.after(900, self._save_current_draft)
        except tk.TclError:
            self._draft_save_job = None

    def _save_current_draft(self) -> None:
        self._draft_save_job = None
        if self.loaded_report_locked or not self.has_permission("edit"):
            return
        try:
            self.draft_store.save(
                self.current_snapshot(),
                report_id=self.loaded_report_id or "",
                is_locked=self.loaded_report_locked,
                raw_entries=self.entries,
            )
        except (OSError, ValueError, tk.TclError):
            return

    def _offer_draft_recovery(self) -> None:
        payload = self.draft_store.load()
        if payload is None:
            return
        decision = messagebox.askyesnocancel(
            "استعادة مسودة تلقائية",
            f"توجد مسودة محفوظة تلقائيًا بتاريخ {payload.get('saved_at', '')}.\n"
            "نعم: استعادتها، لا: حذفها، إلغاء: الاحتفاظ بها دون فتحها.",
            parent=self.root,
        )
        if decision is None:
            return
        if decision is False:
            self.draft_store.clear()
            return
        snapshot = payload["snapshot"]
        report_id = str(payload.get("report_id", ""))
        self.load_snapshot_for_editing(
            report_id,
            snapshot,
            {
                "company_id": snapshot.get("company_id", ""),
                "template_id": snapshot.get("template_id", ""),
                "status": snapshot.get("status", "draft"),
                "is_locked": bool(payload.get("is_locked", False)),
            },
        )
        # New v15.2 drafts include the complete original entry values. Older
        # v15.1 recovery payloads still work via entries_from_snapshot().
        raw_entries = payload.get("raw_entries")
        if isinstance(raw_entries, list) and all(
            isinstance(item, dict)
            and isinstance(item.get("values"), dict)
            and item.get("category") in self.config["categories"]
            and str(item.get("name", "")).strip()
            for item in raw_entries
        ):
            self.entries = copy.deepcopy(raw_entries)
            self.refresh_entries_tree()
            self._schedule_draft_save()
        if not report_id:
            self.loaded_report_id = None
        self.database_status_var.set(
            f"تمت استعادة المسودة التلقائية | {len(self.entries)} سجل"
        )

    def _choose_import_profile(
        self, profiles: list[dict[str, Any]], parent: tk.Misc
    ) -> tuple[bool, dict[str, Any] | None]:
        if not profiles:
            return True, None
        dialog = tk.Toplevel(parent)
        dialog.title("اختيار ملف تعريف الاستيراد")
        dialog.geometry("520x360")
        dialog.minsize(460, 320)
        dialog.configure(bg=Palette.PALE)
        dialog.transient(parent)
        dialog.grab_set()
        tk.Label(
            dialog,
            text="ملف تعريف الاستيراد",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 16, "bold"),
            pady=12,
        ).pack(fill="x")
        choices = ["تلقائي", *[str(item.get("name", "")) for item in profiles]]
        choice_var = tk.StringVar(value=choices[0])
        listbox = tk.Listbox(
            dialog,
            listvariable=tk.StringVar(value=choices),
            font=("Arial", 11),
            justify="right",
            activestyle="none",
            selectmode="browse",
        )
        listbox.pack(fill="both", expand=True, padx=14, pady=12)
        listbox.selection_set(0)
        result: dict[str, Any] = {"accepted": False, "profile": None}

        def accept() -> None:
            selection = listbox.curselection()
            index = selection[0] if selection else 0
            choice_var.set(choices[index])
            result["accepted"] = True
            result["profile"] = None if index == 0 else profiles[index - 1]
            dialog.destroy()

        buttons = tk.Frame(dialog, bg=Palette.PALE)
        buttons.pack(fill="x", padx=14, pady=(0, 12))
        tk.Button(
            buttons,
            text="متابعة إلى المعاينة",
            command=accept,
            bg=Palette.SUCCESS,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=7,
        ).pack(side="right", padx=3)
        tk.Button(
            buttons,
            text="إلغاء",
            command=dialog.destroy,
            bg=Palette.MUTED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=7,
        ).pack(side="right", padx=3)
        listbox.bind("<Double-1>", lambda _event: accept())
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.root.wait_window(dialog)
        return bool(result["accepted"]), result["profile"]

    def show_import_profiles_manager(self) -> None:
        if not self.require_permission("add"):
            return
        window = tk.Toplevel(self.root)
        window.title("ملفات تعريف الاستيراد")
        window.geometry("980x700")
        window.minsize(820, 620)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="تخصيص وربط أعمدة الاستيراد",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=12,
        ).pack(fill="x")
        editor = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        editor.pack(fill="x", padx=12, pady=10)
        profile_name_var = tk.StringVar()
        file_type_var = tk.StringVar(value=".xlsx")
        source_var = tk.StringVar()
        target_var = tk.StringVar(value="الفئة")
        default_category_var = tk.StringVar()
        default_name_var = tk.StringVar()
        shared_var = tk.BooleanVar(value=False)
        source_headers: list[str] = []
        mapping: dict[str, str] = {}

        def labeled_entry(label: str, variable: tk.StringVar, width: int) -> None:
            tk.Label(
                editor, text=label, bg=Palette.WHITE, fg=Palette.NAVY,
                font=("Arial", 9, "bold"),
            ).pack(side="right", padx=(8, 2), pady=8)
            ttk.Entry(editor, textvariable=variable, width=width, justify="right").pack(
                side="right", padx=2, pady=8
            )

        labeled_entry("الاسم", profile_name_var, 22)
        tk.Label(
            editor, text="نوع الملف", bg=Palette.WHITE, fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        ).pack(side="right", padx=(8, 2), pady=8)
        type_combo = ttk.Combobox(
            editor,
            textvariable=file_type_var,
            values=(".xlsx", ".csv", ".tsv", ".txt", ".pptx", ".json"),
            state="readonly",
            width=9,
            justify="center",
        )
        type_combo.pack(side="right", padx=2, pady=8)
        if str(self.current_user.get("role")) == "admin":
            ttk.Checkbutton(editor, text="مشترك", variable=shared_var).pack(
                side="right", padx=10, pady=8
            )

        mapping_box = tk.LabelFrame(
            window,
            text="ربط أعمدة المصدر بحقول التقرير",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
            padx=8,
            pady=8,
        )
        mapping_box.pack(fill="both", expand=True, padx=12, pady=(0, 7))
        mapping_controls = tk.Frame(mapping_box, bg=Palette.WHITE)
        mapping_controls.pack(fill="x")
        source_combo = ttk.Combobox(
            mapping_controls, textvariable=source_var, values=(), width=28, justify="right"
        )
        source_combo.pack(side="right", padx=4, pady=5)
        target_values = ["الفئة", "الاسم/الجهة", "تفاصيل العمل"]
        for category_id in category_ids(self.config):
            for field in self.config["categories"][category_id].get("fields", []):
                label = str(field.get("label", "")).strip()
                if label and label not in target_values:
                    target_values.append(label)
        target_combo = ttk.Combobox(
            mapping_controls,
            textvariable=target_var,
            values=target_values,
            state="readonly",
            width=28,
            justify="right",
        )
        target_combo.pack(side="right", padx=4, pady=5)
        mapping_tree = ttk.Treeview(
            mapping_box, columns=("source", "target"), show="headings", height=10
        )
        mapping_tree.heading("source", text="عمود المصدر")
        mapping_tree.heading("target", text="حقل التقرير")
        mapping_tree.column("source", width=330, anchor="e")
        mapping_tree.column("target", width=330, anchor="e")
        mapping_tree.pack(fill="both", expand=True, pady=7)

        def refresh_mapping_tree() -> None:
            for item in mapping_tree.get_children():
                mapping_tree.delete(item)
            for index, (source, target) in enumerate(mapping.items()):
                mapping_tree.insert("", "end", iid=str(index), values=(source, target))

        def add_mapping() -> None:
            source = source_var.get().strip()
            target = target_var.get().strip()
            if source and target:
                mapping[source] = target
                refresh_mapping_tree()

        def remove_mapping() -> None:
            selection = mapping_tree.selection()
            if selection:
                source = str(mapping_tree.item(selection[0], "values")[0])
                mapping.pop(source, None)
                refresh_mapping_tree()

        def discover_headers() -> None:
            filename = filedialog.askopenfilename(
                parent=window,
                title="اختيار ملف مصدر",
                filetypes=[("ملفات مدعومة", "*.xlsx;*.csv;*.tsv;*.txt;*.pptx;*.json")],
            )
            if not filename:
                return
            try:
                headers, _rows = _read_import_table(Path(filename))
            except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
                messagebox.showerror("تعذر قراءة الملف", str(error), parent=window)
                return
            source_headers[:] = headers
            file_type_var.set(Path(filename).suffix.lower())
            source_combo.configure(values=source_headers)
            if source_headers:
                source_var.set(source_headers[0])
            if not profile_name_var.get().strip():
                profile_name_var.set(Path(filename).stem[:80])
            refresh_profiles()

        tk.Button(
            mapping_controls, text="إضافة الربط", command=add_mapping,
            bg=Palette.SUCCESS, fg=Palette.WHITE, relief="flat",
            font=("Arial", 9, "bold"), padx=10, pady=6,
        ).pack(side="right", padx=3)
        tk.Button(
            mapping_controls, text="إزالة المحدد", command=remove_mapping,
            bg=Palette.DANGER, fg=Palette.WHITE, relief="flat",
            font=("Arial", 9, "bold"), padx=10, pady=6,
        ).pack(side="right", padx=3)
        tk.Button(
            mapping_controls, text="قراءة أعمدة ملف", command=discover_headers,
            bg=Palette.INFO, fg=Palette.WHITE, relief="flat",
            font=("Arial", 9, "bold"), padx=10, pady=6,
        ).pack(side="left", padx=3)

        defaults = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        defaults.pack(fill="x", padx=12, pady=(0, 7))
        tk.Label(
            defaults, text="الفئة الافتراضية", bg=Palette.WHITE, fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        ).pack(side="right", padx=(8, 2), pady=7)
        category_labels = [
            str(self.config["categories"][item].get("label", item))
            for item in category_ids(self.config)
        ]
        ttk.Combobox(
            defaults,
            textvariable=default_category_var,
            values=["", *category_labels],
            state="readonly",
            width=24,
            justify="right",
        ).pack(side="right", padx=3, pady=7)
        labeled_default_name = tk.Label(
            defaults, text="الاسم الافتراضي", bg=Palette.WHITE, fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        )
        labeled_default_name.pack(side="right", padx=(10, 2), pady=7)
        ttk.Entry(defaults, textvariable=default_name_var, width=24, justify="right").pack(
            side="right", padx=3, pady=7
        )

        profiles_box = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        profiles_box.pack(fill="x", padx=12, pady=(0, 10))
        profile_combo_var = tk.StringVar()
        profile_combo = ttk.Combobox(
            profiles_box, textvariable=profile_combo_var, state="readonly", width=38,
            justify="right"
        )
        profile_combo.pack(side="right", padx=5, pady=8)
        loaded_profiles: list[dict[str, Any]] = []

        def refresh_profiles(_event: Any = None) -> None:
            nonlocal loaded_profiles
            try:
                loaded_profiles = self.database.list_import_profiles(
                    file_type_var.get(), self.current_user
                )
            except ValueError:
                loaded_profiles = []
            profile_combo.configure(
                values=[str(item.get("name", "")) for item in loaded_profiles]
            )
            if profile_combo_var.get() not in profile_combo.cget("values"):
                profile_combo_var.set("")

        def load_profile(_event: Any = None) -> None:
            profile = next(
                (item for item in loaded_profiles if item.get("name") == profile_combo_var.get()),
                None,
            )
            if not profile:
                return
            profile_name_var.set(str(profile.get("name", "")))
            file_type_var.set(str(profile.get("file_type", ".xlsx")))
            mapping.clear()
            mapping.update(profile.get("mapping", {}))
            options = profile.get("options", {})
            default_category_var.set(str(options.get("default_category", "")))
            default_name_var.set(str(options.get("default_name", "")))
            shared_var.set(bool(profile.get("is_shared")))
            source_combo.configure(values=list(mapping))
            refresh_mapping_tree()

        def save_profile() -> None:
            try:
                self.database.save_import_profile(
                    profile_name_var.get(),
                    file_type_var.get(),
                    mapping,
                    self.current_user,
                    options={
                        "default_category": default_category_var.get(),
                        "default_name": default_name_var.get(),
                    },
                    is_shared=bool(shared_var.get()),
                )
            except (ValueError, PermissionError, sqlite3.Error) as error:
                messagebox.showerror("تعذر الحفظ", str(error), parent=window)
                return
            refresh_profiles()
            profile_combo_var.set(profile_name_var.get().strip())
            messagebox.showinfo("تم الحفظ", "تم حفظ ملف تعريف الاستيراد.", parent=window)

        def delete_profile() -> None:
            profile = next(
                (item for item in loaded_profiles if item.get("name") == profile_combo_var.get()),
                None,
            )
            if not profile:
                return
            try:
                self.database.delete_import_profile(str(profile["id"]), self.current_user)
            except (ValueError, PermissionError, sqlite3.Error) as error:
                messagebox.showerror("تعذر الحذف", str(error), parent=window)
                return
            profile_combo_var.set("")
            refresh_profiles()

        tk.Button(
            profiles_box, text="حفظ", command=save_profile, bg=Palette.SUCCESS,
            fg=Palette.WHITE, relief="flat", font=("Arial", 9, "bold"), padx=12, pady=6,
        ).pack(side="right", padx=3, pady=7)
        tk.Button(
            profiles_box, text="حذف", command=delete_profile, bg=Palette.DANGER,
            fg=Palette.WHITE, relief="flat", font=("Arial", 9, "bold"), padx=12, pady=6,
        ).pack(side="right", padx=3, pady=7)
        profile_combo.bind("<<ComboboxSelected>>", load_profile)
        type_combo.bind("<<ComboboxSelected>>", refresh_profiles)
        refresh_profiles()

    def import_entries(self) -> None:
        permission = "edit" if self.loaded_report_id else "add"
        if not self.require_permission(permission):
            return
        if self.loaded_report_locked:
            messagebox.showwarning(
                "تقرير مقفل",
                "لا يمكن الاستيراد إلى تقرير معتمد ومقفل.",
                parent=self.root,
            )
            return
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="استيراد سجلات إلى التقرير الحالي",
            filetypes=[
                ("ملفات مدعومة", "*.xlsx;*.csv;*.tsv;*.txt;*.pptx;*.json"),
                ("Excel", "*.xlsx"),
                ("CSV ونص جدولي", "*.csv;*.tsv;*.txt"),
                ("PowerPoint", "*.pptx"),
                ("JSON", "*.json"),
            ],
        )
        if not filename:
            return
        try:
            profiles = self.database.list_import_profiles(
                Path(filename).suffix.lower(), self.current_user
            )
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("تعذر قراءة ملفات التعريف", str(error), parent=self.root)
            return
        accepted, profile = self._choose_import_profile(profiles, self.root)
        if not accepted:
            return
        try:
            preview = analyze_import_path(filename, self.config, profile)
        except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
            messagebox.showerror("تعذر الاستيراد", str(error), parent=self.root)
            return
        preview_lines = []
        for row in preview.get("preview_rows", [])[:5]:
            shown = " | ".join(
                f"{key}: {value}" for key, value in list(row.items())[:6]
                if str(value or "").strip()
            )
            if shown:
                preview_lines.append(shown)
        profile_label = str(preview.get("profile_name") or "تلقائي")
        summary = (
            f"ملف التعريف: {profile_label}\n"
            f"السجلات الجاهزة: {preview['imported']}\n"
            f"السجلات المتجاوزة: {preview['skipped']}\n\n"
            + "\n".join(preview_lines[:5])
        )
        if preview.get("warnings"):
            summary += "\n\nملاحظات:\n" + "\n".join(preview["warnings"][:5])
        if not messagebox.askyesno(
            "معاينة الاستيراد",
            summary + "\n\nاعتماد الاستيراد؟",
            parent=self.root,
        ):
            return
        try:
            result = import_entries_from_path(filename, self.config, profile)
        except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
            messagebox.showerror("تعذر الاستيراد", str(error), parent=self.root)
            return
        imported = int(result["imported"])
        if not imported:
            details = "\n".join(result["warnings"][:8])
            messagebox.showwarning(
                "لم تُستورد سجلات",
                details or "لا يحتوي الملف سجلات قابلة للاستيراد.",
                parent=self.root,
            )
            return
        self.entries.extend(result["entries"])
        self.refresh_entries_tree()
        self.clear_editor()
        self._schedule_draft_save()
        self.database.audit(
            self.actor_name,
            "import_rows",
            "report",
            self.loaded_report_id or "",
            {
                "file": Path(filename).name,
                "imported": imported,
                "skipped": int(result["skipped"]),
            },
        )
        message = f"تم استيراد {imported} سجل إلى التقرير الحالي."
        if result["skipped"]:
            message += f"\nتم تجاوز {result['skipped']} صف غير صالح."
            if result["warnings"]:
                message += "\n\n" + "\n".join(result["warnings"][:8])
        messagebox.showinfo("اكتمل الاستيراد", message, parent=self.root)

    def export_import_template(self) -> None:
        if not self.require_permission("add"):
            return
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="حفظ قالب الاستيراد",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="قالب_استيراد_تقارير_العمل.xlsx",
        )
        if not filename:
            return
        try:
            output = export_import_template_to_path(self.config, filename)
        except (OSError, RuntimeError, ValueError) as error:
            messagebox.showerror("تعذر إنشاء القالب", str(error), parent=self.root)
            return
        messagebox.showinfo("تم إنشاء القالب", f"تم حفظ القالب:\n{output}", parent=self.root)

    def _require_report_rows(self) -> dict[str, Any] | None:
        snapshot = self.current_snapshot()
        if not snapshot["rows"]:
            messagebox.showwarning("لا توجد بيانات", "أضف سجلًا واحدًا على الأقل.")
            return None
        return snapshot

    def save_to_history(self) -> None:
        permission = "edit" if self.loaded_report_id else "add"
        if not self.require_permission(permission):
            return
        if self.loaded_report_locked:
            messagebox.showwarning(
                "تقرير مقفل",
                "هذا التقرير معتمد ومقفل. ألغِ القفل من شاشة السجل قبل تعديله.",
                parent=self.root,
            )
            return
        snapshot = self._require_report_rows()
        if snapshot is None:
            return
        duplicate = self.database.find_duplicate(
            snapshot,
            self.current_company_id,
            self.loaded_report_id,
        )
        if duplicate and not messagebox.askyesno(
            "سجل مكرر",
            "يوجد تقرير مطابق محفوظ مسبقاً:\n"
            f"التاريخ: {duplicate.get('report_date', '')}\n"
            f"الرمز: {duplicate.get('code', '')}\n\n"
            "هل تريد الحفظ رغم التكرار؟",
            parent=self.root,
        ):
            return
        status_code = str(snapshot.get("status", "draft"))
        try:
            self.loaded_report_id = self.database.save_snapshot(
                snapshot,
                self.loaded_report_id,
                company_id=self.current_company_id,
                template_id=self.current_template_id,
                status=status_code,
                actor=self.actor_name,
            )
        except (OSError, sqlite3.Error, ValueError, PermissionError) as error:
            messagebox.showerror(
                "خطأ في قاعدة البيانات",
                f"تعذر حفظ التقرير:\n{error}",
            )
            return
        self.database_status_var.set(
            f"محفوظ - {snapshot['date']} | قاعدة البيانات: "
            f"{self.database.count()} تقرير"
        )
        self.draft_store.clear()
        messagebox.showinfo(
            "تم الحفظ",
            "تم حفظ التقرير في قاعدة البيانات بنجاح.\n"
            "يمكن الرجوع إليه أو تعديله من شاشة السجل والتقارير الدورية.",
        )

    def export_pdf(self) -> None:
        snapshot = self._require_report_rows()
        if snapshot is None:
            return
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                "مكتبة ناقصة",
                "تصدير PDF يحتاج reportlab:\npip install reportlab",
            )
            return
        filename = filedialog.asksaveasfilename(
            title="حفظ التقرير بصيغة PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"تقرير_العمل_{snapshot['date']}.pdf",
        )
        if not filename:
            return
        try:
            output = export_pdf_to_path(snapshot, filename)
        except (OSError, RuntimeError, ValueError) as error:
            messagebox.showerror("تعذر تصدير PDF", str(error))
            return
        self.database.audit(
            self.actor_name,
            "export_pdf",
            "report",
            self.loaded_report_id or "",
            {"path": str(output)},
        )
        messagebox.showinfo("تم التصدير", f"تم إنشاء التقرير:\n{output}")

    def export_images(self) -> None:
        snapshot = self._require_report_rows()
        if snapshot is None:
            return
        if not PILLOW_AVAILABLE:
            messagebox.showerror(
                "مكتبة ناقصة",
                "تصدير الصور يحتاج Pillow:\npip install Pillow",
            )
            return
        filename = filedialog.asksaveasfilename(
            title="حفظ التقرير كصور عالية الدقة",
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            initialfile=f"تقرير_العمل_{snapshot['date']}.png",
        )
        if not filename:
            return
        try:
            outputs = export_images_to_path(snapshot, filename)
        except (OSError, RuntimeError, ValueError) as error:
            messagebox.showerror("تعذر تصدير الصور", str(error))
            return
        self.database.audit(
            self.actor_name,
            "export_image",
            "report",
            self.loaded_report_id or "",
            {"files": len(outputs)},
        )
        message = "\n".join(str(path) for path in outputs)
        messagebox.showinfo("تم التصدير", f"تم إنشاء {len(outputs)} صورة:\n{message}")

    def export_csv(self) -> None:
        snapshot = self._require_report_rows()
        if snapshot is None:
            return
        filename = filedialog.asksaveasfilename(
            title="حفظ ملف CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=f"تقرير_العمل_{snapshot['date']}.csv",
        )
        if not filename:
            return
        try:
            output = export_csv_to_path(snapshot, filename)
        except OSError as error:
            messagebox.showerror("تعذر التصدير", str(error))
            return
        self.database.audit(
            self.actor_name,
            "export_csv",
            "report",
            self.loaded_report_id or "",
            {"path": str(output)},
        )
        messagebox.showinfo("تم التصدير", f"تم إنشاء الملف:\n{output}")

    def export_excel(self) -> None:
        snapshot = self._require_report_rows()
        if snapshot is None:
            return
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(
                "مكتبة ناقصة",
                "تصدير Excel غير متاح في هذا التثبيت.",
                parent=self.root,
            )
            return
        filename = filedialog.asksaveasfilename(
            title="حفظ تقرير Excel منظم",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"تقرير_العمل_{snapshot['date']}.xlsx",
        )
        if not filename:
            return
        try:
            output = export_snapshot_excel_to_path(snapshot, filename)
        except (OSError, RuntimeError, ValueError) as error:
            messagebox.showerror("تعذر تصدير Excel", str(error), parent=self.root)
            return
        self.database.audit(
            self.actor_name,
            "export_excel",
            "report",
            self.loaded_report_id or "",
            {"path": str(output)},
        )
        messagebox.showinfo("تم التصدير", f"تم إنشاء ملف Excel:\n{output}")

    def load_snapshot_for_editing(
        self,
        report_id: str,
        snapshot: dict[str, Any],
        metadata: dict[str, Any] | None = None,
    ) -> None:
        metadata = metadata or {}
        snapshot_config = normalize_config(snapshot.get("config", self.config))
        snapshot_config["identity"]["report_title"] = str(
            snapshot.get("report_title") or snapshot_config["identity"]["report_title"]
        )
        snapshot_config["identity"]["company_name"] = str(
            snapshot.get("company_name", "")
        )
        snapshot_config["identity"]["user_name"] = str(
            snapshot.get("user_name", "")
        )
        snapshot_config["identity"]["default_code"] = str(
            snapshot.get("code", "")
        )
        snapshot_config["identity"]["default_report_number"] = str(
            snapshot.get("page", "1")
        )
        self.config = snapshot_config
        self.category_id_by_label = self._category_label_map()
        self.category_combo.configure(values=list(self.category_id_by_label))
        first_id = category_ids(self.config)[0]
        self.category_var.set(self.config["categories"][first_id]["label"])
        loaded_company_id = str(
            metadata.get("company_id") or snapshot.get("company_id", "")
        )
        if loaded_company_id in self.company_id_to_label:
            self.current_company_id = loaded_company_id
            self.company_var.set(self.company_id_to_label[loaded_company_id])
        loaded_template_id = str(
            metadata.get("template_id") or snapshot.get("template_id", "")
        )
        if loaded_template_id in self.template_id_to_label:
            self.current_template_id = loaded_template_id
            self.template_var.set(self.template_id_to_label[loaded_template_id])
        self.user_var.set(self.actor_name)
        self.code_var.set(str(snapshot.get("code", "")))
        self.page_var.set(str(snapshot.get("page", "1")))
        loaded_date = parse_iso_date(snapshot.get("date")) or date.today()
        self.selected_date = loaded_date
        self.calendar.set_date(loaded_date)
        self.entries = entries_from_snapshot(snapshot)
        self.loaded_report_id = report_id
        self.loaded_report_locked = bool(metadata.get("is_locked", False))
        loaded_status = str(
            metadata.get("status") or snapshot.get("status", "draft")
        )
        self.status_var.set(REPORT_STATUSES.get(loaded_status, REPORT_STATUSES["draft"]))
        self.header_title_label.config(text=self.config["identity"]["report_title"])
        self.root.title(
            f"{self.config['identity']['report_title']} - الإصدار {APP_VERSION}"
        )
        self._update_date_summary()
        self._refresh_name_values()
        self._rebuild_dynamic_fields()
        self.refresh_entries_tree()
        self.clear_editor()
        self.database_status_var.set(
            f"وضع التعديل: {snapshot.get('date', '')}"
            + (" | التقرير معتمد ومقفل" if self.loaded_report_locked else " - احفظ لتحديث التقرير")
        )
        self.save_config()
        self._schedule_draft_save()

    def show_company_manager(self) -> None:
        if not self.require_permission("manage"):
            return
        window = tk.Toplevel(self.root)
        window.title("إدارة ملفات الشركات")
        window.geometry("1050x650")
        window.minsize(900, 580)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="ملفات الشركات والشعارات والبيانات المستقلة",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=12,
        ).pack(fill="x")
        body = tk.Frame(window, bg=Palette.PALE)
        body.pack(fill="both", expand=True, padx=12, pady=10)
        list_box = tk.Frame(body, bg=Palette.WHITE, bd=1, relief="solid")
        list_box.pack(side="right", fill="both", expand=True, padx=(6, 0))
        tree = ttk.Treeview(
            list_box,
            columns=("name", "phone", "tax", "active"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("name", "اسم الشركة", 240),
            ("phone", "الهاتف", 135),
            ("tax", "الرقم الضريبي", 150),
            ("active", "الحالة", 85),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, anchor="e" if column == "name" else "center")
        tree.pack(fill="both", expand=True, padx=6, pady=6)

        form = tk.LabelFrame(
            body,
            text="بيانات الشركة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            width=350,
            padx=12,
            pady=10,
        )
        form.pack(side="left", fill="y", padx=(0, 6))
        form.pack_propagate(False)
        name_var = tk.StringVar()
        address_var = tk.StringVar()
        phone_var = tk.StringVar()
        tax_var = tk.StringVar()
        logo_var = tk.StringVar()
        active_var = tk.BooleanVar(value=True)
        selected_id = tk.StringVar()

        for row_index, (label, variable) in enumerate(
            (
                ("اسم الشركة", name_var),
                ("العنوان", address_var),
                ("الهاتف", phone_var),
                ("الرقم الضريبي", tax_var),
                ("مسار الشعار", logo_var),
            )
        ):
            tk.Label(
                form,
                text=label,
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                font=("Arial", 10, "bold"),
            ).grid(row=row_index * 2, column=0, sticky="e", pady=(7, 2))
            ttk.Entry(
                form,
                textvariable=variable,
                width=38,
                justify="right",
            ).grid(row=row_index * 2 + 1, column=0, sticky="ew", pady=(0, 3))
        tk.Button(
            form,
            text="اختيار شعار PNG / JPG",
            command=lambda: logo_var.set(
                filedialog.askopenfilename(
                    parent=window,
                    title="اختيار شعار الشركة",
                    filetypes=[
                        ("الصور", "*.png;*.jpg;*.jpeg;*.webp;*.bmp"),
                        ("كل الملفات", "*.*"),
                    ],
                )
                or logo_var.get()
            ),
            bg=Palette.INFO,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            pady=6,
        ).grid(row=10, column=0, sticky="ew", pady=7)
        ttk.Checkbutton(
            form,
            text="ملف شركة فعّال",
            variable=active_var,
        ).grid(row=11, column=0, sticky="e", pady=7)
        actions = tk.Frame(form, bg=Palette.WHITE)
        actions.grid(row=12, column=0, sticky="ew", pady=10)

        def clear_form() -> None:
            selected_id.set("")
            name_var.set("")
            address_var.set("")
            phone_var.set("")
            tax_var.set("")
            logo_var.set("")
            active_var.set(True)
            tree.selection_remove(tree.selection())

        def refresh() -> None:
            for item in tree.get_children():
                tree.delete(item)
            for index, company in enumerate(self.database.list_companies(), 1):
                tree.insert(
                    "",
                    "end",
                    iid=str(company["id"]),
                    values=(
                        company["name"],
                        company["phone"],
                        company["tax_number"],
                        "فعّال" if company["active"] else "موقوف",
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            tree.tag_configure("odd", background=Palette.WHITE)
            tree.tag_configure("even", background=Palette.ALT_ROW)

        def load_selected(_event: Any = None) -> None:
            selection = tree.selection()
            if not selection:
                return
            company = self.database.get_company(selection[0])
            if not company:
                return
            selected_id.set(str(company["id"]))
            name_var.set(str(company["name"]))
            address_var.set(str(company["address"]))
            phone_var.set(str(company["phone"]))
            tax_var.set(str(company["tax_number"]))
            logo_var.set(str(company["logo_path"]))
            active_var.set(bool(company["active"]))

        def save_company() -> None:
            clean_name = name_var.get().strip()
            if not clean_name:
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم الشركة.", parent=window)
                return
            company_id = selected_id.get() or uuid.uuid4().hex
            logo_path = logo_var.get().strip()
            stored_logo = logo_path
            if logo_path and Path(logo_path).is_file():
                source_logo = Path(logo_path)
                if self.backup_manager.assets_dir not in source_logo.parents:
                    suffix = source_logo.suffix.lower() or ".png"
                    destination = self.backup_manager.assets_dir / f"{company_id}{suffix}"
                    shutil.copy2(source_logo, destination)
                    stored_logo = str(destination)
            try:
                saved_id = self.database.save_company(
                    {
                        "name": clean_name,
                        "address": address_var.get().strip(),
                        "phone": phone_var.get().strip(),
                        "tax_number": tax_var.get().strip(),
                        "logo_path": stored_logo,
                        "active": active_var.get(),
                    },
                    self.actor_name,
                    company_id,
                )
            except (OSError, sqlite3.Error, ValueError) as error:
                messagebox.showerror("تعذر الحفظ", str(error), parent=window)
                return
            self.current_company_id = saved_id
            self._reload_business_choices()
            self.company_var.set(self.company_id_to_label.get(saved_id, clean_name))
            self.on_company_selected()
            selected_id.set(saved_id)
            logo_var.set(stored_logo)
            refresh()
            messagebox.showinfo("تم الحفظ", "تم حفظ ملف الشركة.", parent=window)

        def deactivate_selected() -> None:
            if not selected_id.get():
                messagebox.showwarning("تنبيه", "اختر شركة أولاً.", parent=window)
                return
            if not messagebox.askyesno(
                "إيقاف ملف الشركة",
                "سيتم إخفاء الشركة من الاختيار مع الاحتفاظ بكل تقاريرها. متابعة؟",
                parent=window,
            ):
                return
            self.database.deactivate_company(selected_id.get(), self.actor_name)
            if self.current_company_id == selected_id.get():
                active_companies = self.database.list_companies(active_only=True)
                if active_companies:
                    self.current_company_id = str(active_companies[0]["id"])
            self._reload_business_choices()
            clear_form()
            refresh()

        tk.Button(
            actions,
            text="جديد",
            command=clear_form,
            bg=Palette.MUTED,
            fg=Palette.WHITE,
            relief="flat",
            padx=11,
            pady=6,
        ).pack(side="right", padx=2)
        tk.Button(
            actions,
            text="حفظ",
            command=save_company,
            bg=Palette.SUCCESS,
            fg=Palette.WHITE,
            relief="flat",
            padx=11,
            pady=6,
        ).pack(side="right", padx=2)
        tk.Button(
            actions,
            text="إيقاف",
            command=deactivate_selected,
            bg=Palette.DANGER,
            fg=Palette.WHITE,
            relief="flat",
            padx=11,
            pady=6,
        ).pack(side="right", padx=2)
        tree.bind("<<TreeviewSelect>>", load_selected)
        refresh()

    def show_template_manager(self) -> None:
        if not self.require_permission("manage"):
            return
        window = tk.Toplevel(self.root)
        window.title("إدارة قوالب التقارير")
        window.geometry("980x620")
        window.minsize(850, 540)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="قوالب التقارير الجاهزة والقوالب المخصصة",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=12,
        ).pack(fill="x")
        tk.Label(
            window,
            text="اختر قالباً جاهزاً أو احفظ مسمياتك وحقولك الحالية كقالب جديد.",
            bg=Palette.PALE,
            fg=Palette.MUTED,
            font=("Arial", 10),
            pady=8,
        ).pack(fill="x")
        body = tk.Frame(window, bg=Palette.PALE)
        body.pack(fill="both", expand=True, padx=12, pady=4)
        tree = ttk.Treeview(
            body,
            columns=("name", "theme", "default"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("name", "اسم القالب", 330),
            ("theme", "التصميم المرئي", 330),
            ("default", "افتراضي", 85),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, anchor="e" if column != "default" else "center")
        tree.pack(fill="both", expand=True)
        form = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        form.pack(fill="x", padx=12, pady=8)
        name_var = tk.StringVar()
        theme_var = tk.StringVar(value=REPORT_THEMES["executive_teal"]["label"])
        default_var = tk.BooleanVar(value=False)
        selected_id = tk.StringVar()
        theme_label_to_id = {
            str(theme["label"]): theme_id for theme_id, theme in REPORT_THEMES.items()
        }
        tk.Label(
            form,
            text="اسم القالب",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(8, 2), pady=10)
        ttk.Entry(form, textvariable=name_var, width=30, justify="right").pack(
            side="right", padx=3, pady=10
        )
        tk.Label(
            form,
            text="التصميم",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(12, 2), pady=10)
        ttk.Combobox(
            form,
            textvariable=theme_var,
            values=list(theme_label_to_id),
            state="readonly",
            width=34,
            justify="right",
        ).pack(side="right", padx=3, pady=10)
        ttk.Checkbutton(form, text="قالب افتراضي", variable=default_var).pack(
            side="right", padx=12, pady=10
        )

        def refresh() -> None:
            self._reload_business_choices()
            for item in tree.get_children():
                tree.delete(item)
            for index, template in enumerate(self.template_records, 1):
                theme_id = str(template["config"].get("report_theme", "executive_teal"))
                tree.insert(
                    "",
                    "end",
                    iid=str(template["id"]),
                    values=(
                        template["name"],
                        REPORT_THEMES.get(theme_id, REPORT_THEMES["executive_teal"])["label"],
                        "نعم" if template["is_default"] else "",
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            tree.tag_configure("odd", background=Palette.WHITE)
            tree.tag_configure("even", background=Palette.ALT_ROW)

        def load_selected(_event: Any = None) -> None:
            selection = tree.selection()
            if not selection:
                return
            template = next(
                (item for item in self.template_records if str(item["id"]) == selection[0]),
                None,
            )
            if not template:
                return
            selected_id.set(str(template["id"]))
            name_var.set(str(template["name"]))
            theme_id = str(template["config"].get("report_theme", "executive_teal"))
            theme_var.set(REPORT_THEMES.get(theme_id, REPORT_THEMES["executive_teal"])["label"])
            default_var.set(bool(template["is_default"]))

        def clear_form() -> None:
            selected_id.set("")
            name_var.set("")
            theme_var.set(REPORT_THEMES["executive_teal"]["label"])
            default_var.set(False)
            tree.selection_remove(tree.selection())

        def save_template() -> None:
            if not name_var.get().strip():
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم القالب.", parent=window)
                return
            template_config = normalize_config(copy.deepcopy(self.config))
            template_config["report_theme"] = theme_label_to_id.get(
                theme_var.get(), "executive_teal"
            )
            try:
                saved_id = self.database.save_template(
                    name_var.get(),
                    template_config,
                    self.actor_name,
                    selected_id.get(),
                    default_var.get(),
                )
            except (sqlite3.Error, ValueError) as error:
                messagebox.showerror("تعذر الحفظ", str(error), parent=window)
                return
            selected_id.set(saved_id)
            refresh()
            messagebox.showinfo("تم الحفظ", "تم حفظ القالب.", parent=window)

        def apply_selected() -> None:
            if not selected_id.get():
                messagebox.showwarning("تنبيه", "اختر قالباً أولاً.", parent=window)
                return
            self.current_template_id = selected_id.get()
            self._reload_business_choices()
            self.template_var.set(
                self.template_id_to_label.get(self.current_template_id, "")
            )
            current = next(
                (
                    item
                    for item in self.template_records
                    if str(item["id"]) == self.current_template_id
                ),
                None,
            )
            if current:
                self.apply_configuration(copy.deepcopy(current["config"]))
                self.database.audit(
                    self.actor_name,
                    "apply",
                    "template",
                    self.current_template_id,
                )
            window.destroy()

        def delete_selected() -> None:
            if not selected_id.get():
                return
            if not messagebox.askyesno(
                "حذف القالب",
                "حذف القالب المحدد؟ لن تتأثر التقارير المحفوظة سابقاً.",
                parent=window,
            ):
                return
            try:
                self.database.delete_template(selected_id.get(), self.actor_name)
            except (sqlite3.Error, ValueError) as error:
                messagebox.showerror("تعذر الحذف", str(error), parent=window)
                return
            clear_form()
            refresh()

        actions = tk.Frame(window, bg=Palette.PALE)
        actions.pack(fill="x", padx=12, pady=(0, 10))
        for text_value, command, color in (
            ("قالب جديد", clear_form, Palette.MUTED),
            ("حفظ الحالي كقالب", save_template, Palette.SUCCESS),
            ("تطبيق القالب المحدد", apply_selected, Palette.PRIMARY),
            ("حذف القالب", delete_selected, Palette.DANGER),
        ):
            tk.Button(
                actions,
                text=text_value,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 9, "bold"),
                padx=12,
                pady=7,
            ).pack(side="right", padx=3)
        tree.bind("<<TreeviewSelect>>", load_selected)
        refresh()

    def show_backup_manager(self) -> None:
        if not self.require_permission("manage"):
            return
        window = tk.Toplevel(self.root)
        window.title("النسخ الاحتياطي والاستعادة")
        window.geometry("880x560")
        window.minsize(760, 500)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="حماية البيانات والنسخ الاحتياطي",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=12,
        ).pack(fill="x")
        info_var = tk.StringVar(
            value="يُنشئ البرنامج نسخة تلقائية يومية ويحتفظ بآخر 30 نسخة."
        )
        tk.Label(
            window,
            textvariable=info_var,
            bg=Palette.PALE,
            fg=Palette.MUTED,
            font=("Arial", 10),
            pady=8,
        ).pack(fill="x")
        tree_box = tk.LabelFrame(
            window,
            text="النسخ التلقائية الموجودة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 11, "bold"),
            padx=7,
            pady=7,
        )
        tree_box.pack(fill="both", expand=True, padx=12, pady=6)
        tree = ttk.Treeview(
            tree_box,
            columns=("file", "date", "size"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("file", "اسم النسخة", 390),
            ("date", "تاريخ الإنشاء", 190),
            ("size", "الحجم", 110),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, anchor="center")
        tree.pack(fill="both", expand=True)

        def refresh() -> None:
            for item in tree.get_children():
                tree.delete(item)
            backups = sorted(
                self.backup_manager.auto_dir.glob("reports_auto_*.db"),
                key=lambda item: item.stat().st_mtime,
                reverse=True,
            )
            for index, backup in enumerate(backups, 1):
                modified = datetime.fromtimestamp(backup.stat().st_mtime).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                size_mb = backup.stat().st_size / (1024 * 1024)
                tree.insert(
                    "",
                    "end",
                    iid=str(backup),
                    values=(backup.name, modified, f"{size_mb:.2f} MB"),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            tree.tag_configure("odd", background=Palette.WHITE)
            tree.tag_configure("even", background=Palette.ALT_ROW)

        def create_auto_now() -> None:
            try:
                output = self.backup_manager.automatic_backup(force=True)
            except (OSError, sqlite3.Error) as error:
                messagebox.showerror("تعذر النسخ", str(error), parent=window)
                return
            self.database.audit(
                self.actor_name,
                "automatic_backup",
                "database",
                details={"path": str(output or "")},
            )
            info_var.set(f"تم إنشاء نسخة تلقائية: {output}")
            refresh()

        def check_health() -> None:
            try:
                result = self.database.health_check()
            except sqlite3.Error as error:
                messagebox.showerror("تعذر الفحص", str(error), parent=window)
                return
            size_mb = int(result["database_size_bytes"]) / (1024 * 1024)
            if result["ok"]:
                info_var.set(
                    f"قاعدة البيانات سليمة | {result['reports']} تقرير | "
                    f"{result['rows']} سجل | {size_mb:.2f} MB"
                )
            else:
                info_var.set(
                    f"قاعدة البيانات تحتاج مراجعة | "
                    f"مشكلات العلاقات: {len(result['foreign_key_issues'])}"
                )

        def maintain_database() -> None:
            if not messagebox.askyesno(
                "تأكيد الصيانة الآمنة",
                "سيُنشئ البرنامج نسخة حماية أولًا، ثم يفحص قاعدة البيانات "
                "ويضغطها ويحسن أداءها. هل تريد المتابعة؟",
                parent=window,
            ):
                return
            try:
                result = self.backup_manager.maintain(self.actor_name)
            except (OSError, sqlite3.Error, ValueError) as error:
                messagebox.showerror("تعذر الصيانة", str(error), parent=window)
                return
            reclaimed_mb = int(result["reclaimed_bytes"]) / (1024 * 1024)
            messagebox.showinfo(
                "اكتملت الصيانة",
                "تم إنشاء نسخة حماية وفحص قاعدة البيانات وتحسينها بنجاح.\n"
                f"المساحة المستعادة: {reclaimed_mb:.2f} MB\n"
                f"نسخة الحماية: {result['backup']}",
                parent=window,
            )
            refresh()
            check_health()

        def create_manual() -> None:
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ نسخة احتياطية كاملة",
                defaultextension=".zip",
                filetypes=[("نسخة احتياطية ZIP", "*.zip")],
                initialfile=f"نسخة_تقارير_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            )
            if not filename:
                return
            try:
                output = self.backup_manager.create_archive(Path(filename))
            except (OSError, sqlite3.Error, zipfile.BadZipFile) as error:
                messagebox.showerror("تعذر النسخ", str(error), parent=window)
                return
            self.database.audit(
                self.actor_name,
                "manual_backup",
                "database",
                details={"path": str(output)},
            )
            messagebox.showinfo(
                "تم النسخ",
                "تم إنشاء نسخة كاملة تشمل قاعدة البيانات والإعدادات وشعارات الشركات.",
                parent=window,
            )

        def restore_selected() -> None:
            filename = filedialog.askopenfilename(
                parent=window,
                title="اختيار نسخة احتياطية للاستعادة",
                filetypes=[
                    ("نسخة احتياطية", "*.zip;*.db;*.sqlite;*.sqlite3"),
                    ("كل الملفات", "*.*"),
                ],
            )
            if not filename:
                return
            if not messagebox.askyesno(
                "تأكيد الاستعادة",
                "سيتم أخذ نسخة حماية تلقائية أولاً ثم استبدال البيانات الحالية "
                "بمحتوى النسخة المحددة. هل تريد المتابعة؟",
                parent=window,
            ):
                return
            try:
                result = self.backup_manager.restore_archive(Path(filename))
                self.database.audit(
                    self.actor_name,
                    "restore",
                    "database",
                    details={"source": filename, **result},
                )
                self.config = normalize_config(read_json(self.config_path, self.config))
                default_company, default_template = self.database.ensure_defaults(self.config)
                self.current_company_id = str(
                    self.config.get("app_state", {}).get("company_id")
                    or default_company
                )
                self.current_template_id = str(
                    self.config.get("app_state", {}).get("template_id")
                    or default_template
                )
                self._reload_business_choices()
            except (OSError, sqlite3.Error, ValueError, zipfile.BadZipFile) as error:
                messagebox.showerror("تعذر الاستعادة", str(error), parent=window)
                return
            self.database_status_var.set(
                f"تمت الاستعادة | قاعدة البيانات: {self.database.count()} تقرير"
            )
            messagebox.showinfo(
                "تمت الاستعادة",
                "تمت استعادة البيانات بنجاح.\n"
                "أغلق البرنامج وافتحه مرة واحدة لتطبيق جميع إعدادات الواجهة المستعادة.",
                parent=window,
            )
            refresh()

        actions = tk.Frame(window, bg=Palette.PALE)
        actions.pack(fill="x", padx=12, pady=(4, 12))
        action_rows = (
            tk.Frame(actions, bg=Palette.PALE),
            tk.Frame(actions, bg=Palette.PALE),
        )
        for row in action_rows:
            row.pack(fill="x", pady=2)
        for index, (text_value, command, color) in enumerate((
            ("نسخة تلقائية الآن", create_auto_now, Palette.INFO),
            ("حفظ نسخة كاملة ZIP", create_manual, Palette.SUCCESS),
            ("استعادة نسخة", restore_selected, Palette.DANGER),
            ("فحص البيانات", check_health, Palette.PRIMARY),
            ("نسخة حماية وصيانة", maintain_database, Palette.ACCENT),
            ("تحديث القائمة", refresh, Palette.MUTED),
        )):
            tk.Button(
                action_rows[0 if index < 3 else 1],
                text=text_value,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 10, "bold"),
                padx=13,
                pady=7,
            ).pack(side="right", padx=3)
        refresh()
        check_health()

    def show_administration(self) -> None:
        if not self.require_permission("manage"):
            return
        window = tk.Toplevel(self.root)
        window.title("المستخدمون والصلاحيات وسجل التدقيق")
        window.geometry("1180x720")
        window.minsize(1000, 620)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="إدارة المستخدمين والصلاحيات وسجل التدقيق",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=12,
        ).pack(fill="x")
        notebook = ttk.Notebook(window)
        notebook.pack(fill="both", expand=True, padx=12, pady=10)
        users_tab = tk.Frame(notebook, bg=Palette.PALE)
        audit_tab = tk.Frame(notebook, bg=Palette.PALE)
        notebook.add(users_tab, text="المستخدمون والصلاحيات")
        notebook.add(audit_tab, text="سجل التدقيق")

        users_tree = ttk.Treeview(
            users_tab,
            columns=("username", "display", "role", "active", "last"),
            show="headings",
            selectmode="browse",
            height=12,
        )
        for column, title, width in (
            ("username", "اسم الدخول", 170),
            ("display", "الاسم الظاهر", 210),
            ("role", "الصلاحية", 150),
            ("active", "الحالة", 90),
            ("last", "آخر دخول", 190),
        ):
            users_tree.heading(column, text=title)
            users_tree.column(column, width=width, anchor="center")
        users_tree.pack(fill="both", expand=True, padx=6, pady=6)
        user_form = tk.Frame(users_tab, bg=Palette.WHITE, bd=1, relief="solid")
        user_form.pack(fill="x", padx=6, pady=(0, 6))
        username_var = tk.StringVar()
        display_var = tk.StringVar()
        password_var = tk.StringVar()
        role_var = tk.StringVar(value=ROLE_LABELS["viewer"])
        active_var = tk.BooleanVar(value=True)
        selected_user_id = tk.StringVar()
        role_label_to_id = {label: key for key, label in ROLE_LABELS.items()}
        for label, variable, width, show in (
            ("اسم الدخول", username_var, 18, ""),
            ("الاسم الظاهر", display_var, 20, ""),
            ("كلمة المرور الجديدة", password_var, 20, "●"),
        ):
            tk.Label(
                user_form,
                text=label,
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                font=("Arial", 9, "bold"),
            ).pack(side="right", padx=(8, 2), pady=9)
            ttk.Entry(
                user_form,
                textvariable=variable,
                width=width,
                justify="right",
                show=show,
            ).pack(side="right", padx=2, pady=9)
        tk.Label(
            user_form,
            text="الصلاحية",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        ).pack(side="right", padx=(8, 2), pady=9)
        ttk.Combobox(
            user_form,
            textvariable=role_var,
            values=list(role_label_to_id),
            state="readonly",
            width=14,
            justify="right",
        ).pack(side="right", padx=2, pady=9)
        ttk.Checkbutton(user_form, text="فعّال", variable=active_var).pack(
            side="right", padx=8, pady=9
        )

        def refresh_users() -> None:
            for item in users_tree.get_children():
                users_tree.delete(item)
            for index, user in enumerate(self.database.list_users(), 1):
                users_tree.insert(
                    "",
                    "end",
                    iid=str(user["id"]),
                    values=(
                        user["username"],
                        user["display_name"],
                        ROLE_LABELS.get(str(user["role"]), user["role"]),
                        "فعّال" if user["active"] else "موقوف",
                        str(user["last_login"]).replace("T", " "),
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            users_tree.tag_configure("odd", background=Palette.WHITE)
            users_tree.tag_configure("even", background=Palette.ALT_ROW)

        def clear_user_form() -> None:
            selected_user_id.set("")
            username_var.set("")
            display_var.set("")
            password_var.set("")
            role_var.set(ROLE_LABELS["viewer"])
            active_var.set(True)
            users_tree.selection_remove(users_tree.selection())

        def load_user(_event: Any = None) -> None:
            selection = users_tree.selection()
            if not selection:
                return
            user = next(
                (
                    item
                    for item in self.database.list_users()
                    if str(item["id"]) == selection[0]
                ),
                None,
            )
            if not user:
                return
            selected_user_id.set(str(user["id"]))
            username_var.set(str(user["username"]))
            display_var.set(str(user["display_name"]))
            password_var.set("")
            role_var.set(ROLE_LABELS.get(str(user["role"]), ROLE_LABELS["viewer"]))
            active_var.set(bool(user["active"]))

        def save_user() -> None:
            try:
                saved_id = self.database.save_user(
                    username_var.get(),
                    display_var.get(),
                    role_label_to_id.get(role_var.get(), "viewer"),
                    active_var.get(),
                    self.actor_name,
                    user_id=selected_user_id.get(),
                    password=password_var.get() if password_var.get() else None,
                )
            except (sqlite3.Error, ValueError) as error:
                messagebox.showerror("تعذر حفظ المستخدم", str(error), parent=window)
                return
            selected_user_id.set(saved_id)
            password_var.set("")
            refresh_users()
            messagebox.showinfo("تم الحفظ", "تم حفظ حساب المستخدم.", parent=window)

        user_actions = tk.Frame(users_tab, bg=Palette.PALE)
        user_actions.pack(fill="x", padx=6, pady=(0, 8))
        for text_value, command, color in (
            ("مستخدم جديد", clear_user_form, Palette.MUTED),
            ("حفظ المستخدم", save_user, Palette.SUCCESS),
        ):
            tk.Button(
                user_actions,
                text=text_value,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 10, "bold"),
                padx=14,
                pady=7,
            ).pack(side="right", padx=3)
        users_tree.bind("<<TreeviewSelect>>", load_user)
        refresh_users()

        audit_bar = tk.Frame(audit_tab, bg=Palette.WHITE, bd=1, relief="solid")
        audit_bar.pack(fill="x", padx=6, pady=6)
        audit_search_var = tk.StringVar()
        tk.Label(
            audit_bar,
            text="بحث في التدقيق",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(8, 2), pady=8)
        ttk.Entry(
            audit_bar,
            textvariable=audit_search_var,
            width=35,
            justify="right",
        ).pack(side="right", padx=3, pady=8)
        audit_tree = ttk.Treeview(
            audit_tab,
            columns=("time", "user", "action", "entity", "id", "details"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("time", "الوقت", 175),
            ("user", "المستخدم", 140),
            ("action", "الإجراء", 110),
            ("entity", "نوع العنصر", 110),
            ("id", "المعرّف", 135),
            ("details", "التفاصيل", 360),
        ):
            audit_tree.heading(column, text=title)
            audit_tree.column(column, width=width, anchor="e" if column == "details" else "center")
        audit_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        def refresh_audit() -> None:
            for item in audit_tree.get_children():
                audit_tree.delete(item)
            for index, row in enumerate(
                self.database.list_audit(audit_search_var.get()),
                1,
            ):
                audit_tree.insert(
                    "",
                    "end",
                    values=(
                        str(row["created_at"]).replace("T", " "),
                        row["username"],
                        row["action"],
                        row["entity_type"],
                        str(row["entity_id"])[:18],
                        row["details_json"],
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            audit_tree.tag_configure("odd", background=Palette.WHITE)
            audit_tree.tag_configure("even", background=Palette.ALT_ROW)

        def export_audit() -> None:
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="تصدير سجل التدقيق",
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")],
                initialfile=f"سجل_التدقيق_{date.today().isoformat()}.csv",
            )
            if not filename:
                return
            rows = self.database.list_audit(audit_search_var.get(), 10_000)
            with Path(filename).open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(
                    ["الوقت", "المستخدم", "الإجراء", "نوع العنصر", "المعرف", "التفاصيل"]
                )
                for row in rows:
                    writer.writerow(
                        [
                            row["created_at"],
                            row["username"],
                            row["action"],
                            row["entity_type"],
                            row["entity_id"],
                            row["details_json"],
                        ]
                    )
            messagebox.showinfo("تم التصدير", f"تم إنشاء:\n{filename}", parent=window)

        tk.Button(
            audit_bar,
            text="بحث / تحديث",
            command=refresh_audit,
            bg=Palette.INFO,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=11,
            pady=6,
        ).pack(side="right", padx=3, pady=6)
        tk.Button(
            audit_bar,
            text="تصدير CSV",
            command=export_audit,
            bg=Palette.SUCCESS,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 9, "bold"),
            padx=11,
            pady=6,
        ).pack(side="left", padx=6, pady=6)
        refresh_audit()

    def show_dashboard(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("لوحة المؤشرات الإدارية")
        window.geometry("1380x860")
        window.minsize(1160, 720)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="لوحة المؤشرات اليومية والأسبوعية والشهرية",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 18, "bold"),
            pady=12,
        ).pack(fill="x")
        start_default, end_default = month_range(date.today())
        saved_dashboard = self.config.get("app_state", {}).get("dashboard_filters", {})
        if not isinstance(saved_dashboard, dict):
            saved_dashboard = {}
        start_var = tk.StringVar(value=str(saved_dashboard.get("start_date") or start_default.isoformat()))
        end_var = tk.StringVar(value=str(saved_dashboard.get("end_date") or end_default.isoformat()))
        company_var = tk.StringVar(value=str(saved_dashboard.get("company") or "كل الشركات"))
        scenario_var = tk.StringVar(value=str(saved_dashboard.get("scenario_percent") or "0"))
        company_map = {
            "كل الشركات": "",
            **{
                str(company["name"]): str(company["id"])
                for company in self.database.list_companies()
            },
        }
        if company_var.get() not in company_map:
            company_var.set("كل الشركات")
        filters = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        filters.pack(fill="x", padx=12, pady=(10, 5))
        for label, variable in (("من", start_var), ("إلى", end_var)):
            tk.Label(
                filters,
                text=label,
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                font=("Arial", 10, "bold"),
            ).pack(side="right", padx=(9, 2), pady=8)
            ttk.Entry(
                filters,
                textvariable=variable,
                width=13,
                justify="center",
            ).pack(side="right", padx=2, pady=8)
        tk.Label(
            filters,
            text="الشركة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(12, 2), pady=8)
        ttk.Combobox(
            filters,
            textvariable=company_var,
            values=list(company_map),
            state="readonly",
            width=24,
            justify="right",
        ).pack(side="right", padx=3, pady=8)
        tk.Label(
            filters,
            text="سيناريو التغير %",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(12, 2), pady=8)
        ttk.Spinbox(
            filters,
            from_=-80,
            to=300,
            increment=5,
            textvariable=scenario_var,
            width=7,
            justify="center",
        ).pack(side="right", padx=3, pady=8)

        kpi_frame = tk.Frame(window, bg=Palette.PALE)
        kpi_frame.pack(fill="x", padx=12, pady=5)
        kpi_variables = [tk.StringVar(value="0") for _ in range(5)]
        for index, (label, variable) in enumerate(
            zip(
                (
                    "عدد التقارير",
                    "عدد السجلات",
                    "إجمالي الأعمال",
                    "المتوسط لكل تقرير",
                    "متوسط الجودة",
                ),
                kpi_variables,
            )
        ):
            card = tk.Frame(
                kpi_frame,
                bg=Palette.WHITE,
                bd=1,
                relief="solid",
                padx=10,
                pady=7,
            )
            card.grid(row=0, column=index, sticky="ew", padx=4)
            kpi_frame.grid_columnconfigure(index, weight=1, uniform="kpi")
            tk.Label(
                card,
                textvariable=variable,
                bg=Palette.WHITE,
                fg=Palette.PRIMARY,
                font=("Arial", 20, "bold"),
            ).pack()
            tk.Label(
                card,
                text=label,
                bg=Palette.WHITE,
                fg=Palette.MUTED,
                font=("Arial", 9, "bold"),
            ).pack()

        charts_frame = tk.Frame(window, bg=Palette.PALE)
        charts_frame.pack(fill="both", expand=True, padx=12, pady=5)
        daily_canvas = tk.Canvas(
            charts_frame,
            bg=Palette.WHITE,
            highlightbackground=Palette.BORDER,
            highlightthickness=1,
            height=300,
        )
        category_canvas = tk.Canvas(
            charts_frame,
            bg=Palette.WHITE,
            highlightbackground=Palette.BORDER,
            highlightthickness=1,
            height=300,
        )
        people_canvas = tk.Canvas(
            charts_frame,
            bg=Palette.WHITE,
            highlightbackground=Palette.BORDER,
            highlightthickness=1,
            height=300,
        )
        for index, canvas_widget in enumerate(
            (daily_canvas, category_canvas, people_canvas)
        ):
            canvas_widget.grid(row=0, column=index, sticky="nsew", padx=4)
            charts_frame.grid_columnconfigure(index, weight=1, uniform="chart")
        charts_frame.grid_rowconfigure(0, weight=1)

        comparison_box = tk.LabelFrame(
            window,
            text="المقارنات التفصيلية",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
            padx=6,
            pady=6,
        )
        comparison_box.pack(fill="x", padx=12, pady=5)
        comparison_tree = ttk.Treeview(
            comparison_box,
            columns=("type", "label", "records", "total"),
            show="headings",
            height=6,
        )
        for column, title, width in (
            ("type", "نوع المقارنة", 150),
            ("label", "الفئة / الاسم / الفترة", 360),
            ("records", "عدد السجلات", 130),
            ("total", "إجمالي الأعمال", 150),
        ):
            comparison_tree.heading(column, text=title)
            comparison_tree.column(
                column,
                width=width,
                anchor="e" if column == "label" else "center",
            )
        comparison_tree.pack(fill="x")
        freshness_var = tk.StringVar()
        tk.Label(
            window,
            textvariable=freshness_var,
            bg=Palette.PALE,
            fg=Palette.MUTED,
            font=("Arial", 8),
        ).pack(anchor="e", padx=16)
        intelligence_var = tk.StringVar(value="")
        tk.Label(
            window,
            textvariable=intelligence_var,
            bg=Palette.WHITE,
            fg=Palette.PRIMARY_DARK,
            font=("Arial", 9, "bold"),
            justify="right",
            anchor="e",
            wraplength=1280,
            padx=12,
            pady=7,
        ).pack(fill="x", padx=12, pady=(2, 4))
        current_analytics: dict[str, Any] = {}

        def draw_empty(canvas_widget: tk.Canvas, title: str) -> tuple[int, int]:
            canvas_widget.delete("all")
            width = max(340, canvas_widget.winfo_width())
            height = max(250, canvas_widget.winfo_height())
            canvas_widget.create_text(
                width - 18,
                22,
                text=title,
                anchor="ne",
                fill=Palette.NAVY,
                font=("Arial", 11, "bold"),
            )
            return width, height

        def draw_bar_chart(
            canvas_widget: tk.Canvas,
            rows: list[dict[str, Any]],
            title: str,
        ) -> None:
            width, height = draw_empty(canvas_widget, title)
            shown = rows[:7]
            if not shown:
                canvas_widget.create_text(
                    width / 2,
                    height / 2,
                    text="لا توجد بيانات في الفترة المحددة",
                    fill=Palette.MUTED,
                    font=("Arial", 10),
                )
                return
            max_value = max(
                [coerce_nonnegative_int(row.get("total", 0)) for row in shown] or [1]
            )
            plot_left = 35
            plot_right = width - 135
            plot_top = 48
            row_height = (height - plot_top - 18) / len(shown)
            for index, row in enumerate(shown):
                y = plot_top + index * row_height
                value = coerce_nonnegative_int(row.get("total", 0))
                bar_width = max(
                    2,
                    (plot_right - plot_left)
                    * (value / max_value if max_value else 0),
                )
                canvas_widget.create_rectangle(
                    plot_left,
                    y + 6,
                    plot_left + bar_width,
                    y + row_height - 6,
                    fill=Palette.ACCENT if index % 2 == 0 else Palette.INFO,
                    outline="",
                )
                canvas_widget.create_text(
                    plot_left + bar_width + 5,
                    y + row_height / 2,
                    text=str(value),
                    anchor="w",
                    fill=Palette.INK,
                    font=("Arial", 8, "bold"),
                )
                canvas_widget.create_text(
                    width - 10,
                    y + row_height / 2,
                    text=rtl_visual(str(row.get("label", ""))[:23]),
                    anchor="e",
                    fill=Palette.INK,
                    font=("Arial", 8),
                )

        def draw_line_chart(rows: list[dict[str, Any]]) -> None:
            width, height = draw_empty(daily_canvas, "اتجاه الإنتاجية خلال الفترة")
            if not rows:
                daily_canvas.create_text(
                    width / 2,
                    height / 2,
                    text="لا توجد بيانات في الفترة المحددة",
                    fill=Palette.MUTED,
                    font=("Arial", 10),
                )
                return
            shown = rows[-18:]
            values = [coerce_nonnegative_int(row.get("total", 0)) for row in shown]
            max_value = max(values or [1])
            left, right, top, bottom = 38, width - 20, 52, height - 36
            daily_canvas.create_line(left, bottom, right, bottom, fill=Palette.BORDER)
            daily_canvas.create_line(left, top, left, bottom, fill=Palette.BORDER)
            points: list[float] = []
            for index, value in enumerate(values):
                x = left + (right - left) * index / max(1, len(values) - 1)
                y = bottom - (bottom - top) * (value / max_value if max_value else 0)
                points.extend([x, y])
                daily_canvas.create_oval(
                    x - 3,
                    y - 3,
                    x + 3,
                    y + 3,
                    fill=Palette.ACCENT,
                    outline=Palette.PRIMARY,
                )
                if index in {0, len(values) - 1} or len(values) <= 7:
                    daily_canvas.create_text(
                        x,
                        bottom + 14,
                        text=str(shown[index].get("label", ""))[5:],
                        fill=Palette.MUTED,
                        font=("Arial", 7),
                    )
            if len(points) >= 4:
                daily_canvas.create_line(
                    *points,
                    fill=Palette.PRIMARY,
                    width=3,
                    smooth=True,
                )
            daily_canvas.create_text(
                left,
                top - 10,
                text=str(max_value),
                anchor="w",
                fill=Palette.MUTED,
                font=("Arial", 7),
            )

        def validated_dates() -> tuple[str, str] | None:
            start = parse_iso_date(start_var.get())
            end = parse_iso_date(end_var.get())
            if start is None or end is None or start > end:
                messagebox.showwarning(
                    "فترة غير صحيحة",
                    "تحقق من تاريخ البداية والنهاية بصيغة YYYY-MM-DD.",
                    parent=window,
                )
                return None
            return start.isoformat(), end.isoformat()

        def refresh_dashboard() -> None:
            selected_range = validated_dates()
            if selected_range is None:
                return
            analytics = self.database.analytics_summary(
                selected_range[0],
                selected_range[1],
                company_map.get(company_var.get(), ""),
            )
            try:
                scenario_percent = max(-80.0, min(300.0, float(scenario_var.get() or 0)))
            except ValueError:
                scenario_percent = 0.0
                scenario_var.set("0")
            analytics["scenario"] = scenario_forecast(
                [item.get("total", 0) for item in analytics.get("daily", [])],
                scenario_percent,
            )
            self.config["app_state"]["dashboard_filters"] = {
                "start_date": selected_range[0],
                "end_date": selected_range[1],
                "company": company_var.get(),
                "scenario_percent": scenario_var.get(),
            }
            try:
                atomic_write_json(self.config_path, self.config)
            except OSError:
                pass
            current_analytics.clear()
            current_analytics.update(analytics)
            for variable, value in zip(
                kpi_variables,
                (
                    analytics["reports"],
                    analytics["records"],
                    analytics["work_total"],
                    analytics["average_per_report"],
                    round(float(analytics.get("quality", {}).get("average") or 0), 1),
                ),
            ):
                variable.set(f"{value:,}" if isinstance(value, int) else str(value))
            window.update_idletasks()
            draw_line_chart(analytics["daily"])
            draw_bar_chart(category_canvas, analytics["categories"], "حجم الأعمال حسب الفئة")
            draw_bar_chart(people_canvas, analytics["people"], "حجم الأعمال حسب الاسم أو الجهة")
            for item in comparison_tree.get_children():
                comparison_tree.delete(item)
            comparison_rows: list[tuple[str, dict[str, Any]]] = [
                *[("فئة", row) for row in analytics["categories"]],
                *[("اسم/جهة", row) for row in analytics["people"][:10]],
                *[("يوم", row) for row in analytics["daily"][-10:]],
            ]
            for index, (kind, row) in enumerate(comparison_rows, 1):
                comparison_tree.insert(
                    "",
                    "end",
                    values=(
                        kind,
                        row.get("label", ""),
                        row.get("records", row.get("reports", 0)),
                        row.get("total", 0),
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            comparison_tree.tag_configure("odd", background=Palette.WHITE)
            comparison_tree.tag_configure("even", background=Palette.ALT_ROW)
            freshness_var.set(
                f"مصدر البيانات: قاعدة التقارير المحلية | مستوى البيانات: سجل عمل "
                f"| آخر تحديث: {str(analytics.get('freshness', '')).replace('T', ' ') or '-'}"
            )
            insight_text = "   |   ".join(
                    f"{item.get('title', '')}: {item.get('message', '')}"
                    for item in analytics.get("insights", [])
                )
            scenario = analytics.get("scenario", {})
            if scenario.get("available"):
                insight_text += (
                    f"   |   محاكاة {float(scenario.get('adjustment_percent') or 0):+.1f}%: "
                    f"{float(scenario.get('prediction') or 0):.1f} وحدة، من "
                    f"{float(scenario.get('lower_bound') or 0):.1f} إلى "
                    f"{float(scenario.get('upper_bound') or 0):.1f}"
                )
            intelligence_var.set("الذكاء الداخلي: " + insight_text)

        def set_range(kind: str) -> None:
            if kind == "day":
                start = end = date.today()
            elif kind == "week":
                start, end = week_range(date.today())
            else:
                start, end = month_range(date.today())
            start_var.set(start.isoformat())
            end_var.set(end.isoformat())
            refresh_dashboard()

        def export_dashboard_excel() -> None:
            if not current_analytics:
                refresh_dashboard()
            if not current_analytics:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ التقرير الإداري Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"تقرير_إداري_{start_var.get()}_{end_var.get()}.xlsx",
            )
            if not filename:
                return
            company_id = company_map.get(company_var.get(), "")
            company = self.database.get_company(company_id) if company_id else None
            try:
                output = export_analytics_excel_to_path(
                    current_analytics,
                    company,
                    filename,
                )
            except (OSError, RuntimeError, ValueError) as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            self.database.audit(
                self.actor_name,
                "export_excel",
                "management_report",
                details={"path": str(output)},
            )
            messagebox.showinfo("تم التصدير", f"تم إنشاء:\n{output}", parent=window)

        def export_dashboard_pdf() -> None:
            if not current_analytics:
                refresh_dashboard()
            if not current_analytics:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ التقرير الإداري الجاهز للطباعة",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"تقرير_إداري_{start_var.get()}_{end_var.get()}.pdf",
            )
            if not filename:
                return
            company_id = company_map.get(company_var.get(), "")
            company = self.database.get_company(company_id) if company_id else None
            try:
                output = export_management_pdf_to_path(
                    current_analytics,
                    company,
                    filename,
                )
            except (OSError, RuntimeError, ValueError) as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            self.database.audit(
                self.actor_name,
                "export_pdf",
                "management_report",
                details={"path": str(output)},
            )
            messagebox.showinfo(
                "تم التصدير",
                f"تم إنشاء التقرير الإداري الجاهز للطباعة:\n{output}",
                parent=window,
            )

        button_specs = (
            ("تطبيق", refresh_dashboard, Palette.INFO),
            ("اليوم", lambda: set_range("day"), Palette.PRIMARY),
            ("هذا الأسبوع", lambda: set_range("week"), Palette.SUCCESS),
            ("هذا الشهر", lambda: set_range("month"), Palette.BLUE),
            ("تقرير إداري PDF", export_dashboard_pdf, Palette.DANGER),
            ("تقرير إداري Excel", export_dashboard_excel, Palette.ACCENT),
        )
        for text_value, command, color in button_specs:
            tk.Button(
                filters,
                text=text_value,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 9, "bold"),
                padx=9,
                pady=6,
            ).pack(side="left", padx=2, pady=6)
        refresh_dashboard()

    def show_trash(self) -> None:
        if not self.require_permission("view"):
            return
        window = tk.Toplevel(self.root)
        window.title("المحذوفات / أرشيف التقارير")
        window.geometry("980x620")
        window.configure(bg=Palette.PALE)
        tk.Label(window, text="المحذوفات / الأرشيف", bg=Palette.NAVY, fg=Palette.WHITE, font=("Arial", 18, "bold"), pady=12).pack(fill="x")
        frame = tk.Frame(window, bg=Palette.WHITE, padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        tree = ttk.Treeview(frame, columns=("date","company","number","archived","by"), show="headings", selectmode="extended")
        for column, title, width in (("date","التاريخ",120),("company","الشركة",250),("number","رقم التقرير",120),("archived","تاريخ الحذف",180),("by","بواسطة",160)):
            tree.heading(column, text=title); tree.column(column, width=width, anchor="center")
        scroll=ttk.Scrollbar(frame, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        def refresh():
            tree.delete(*tree.get_children())
            for item in self.database.list_reports(archived_only=True):
                tree.insert("", "end", iid=str(item["id"]), values=(item.get("report_date",""), item.get("company_name",""), item.get("report_number",""), item.get("archived_at",""), item.get("archived_by","")))
        def selected_ids(): return list(tree.selection())
        def restore():
            ids=selected_ids()
            if not ids: return
            self.database.restore_archived_reports(ids, actor=self.actor_name); refresh()
            self.database_status_var.set(f"قاعدة البيانات: {self.database.count()} تقرير")
        def purge():
            ids=selected_ids()
            if not ids: return
            if not self.require_permission("delete"): return
            if messagebox.askyesno("حذف نهائي", "حذف التقارير المحددة نهائيًا؟ لا يمكن التراجع.", parent=window):
                self.database.purge_archived_reports(ids, actor=self.actor_name); refresh()
        actions=tk.Frame(window,bg=Palette.PALE); actions.pack(fill="x",padx=12,pady=(0,12))
        tk.Button(actions,text="استعادة المحدد",command=restore,bg=Palette.GREEN,fg=Palette.WHITE,relief="flat",font=("Arial",10,"bold"),padx=14,pady=7).pack(side="right",padx=4)
        tk.Button(actions,text="حذف نهائي",command=purge,bg=Palette.RED,fg=Palette.WHITE,relief="flat",font=("Arial",10,"bold"),padx=14,pady=7).pack(side="right",padx=4)
        tk.Button(actions,text="إغلاق",command=window.destroy,bg=Palette.MUTED,fg=Palette.WHITE,relief="flat",font=("Arial",10,"bold"),padx=14,pady=7).pack(side="left",padx=4)
        refresh()

    def show_history(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("قاعدة البيانات والتقارير الدورية")
        window.geometry("1380x820")
        window.minsize(1120, 700)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)

        tk.Label(
            window,
            text="قاعدة البيانات والتقارير الأسبوعية والشهرية",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 18, "bold"),
            pady=12,
        ).pack(fill="x")

        start_default, end_default = month_range(date.today())
        saved_history = self.config.get("app_state", {}).get("history_filters", {})
        if not isinstance(saved_history, dict):
            saved_history = {}
        range_bar = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        range_bar.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(
            range_bar,
            text="من تاريخ",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(8, 3), pady=8)
        start_var = tk.StringVar(value=str(saved_history.get("start_date") or start_default.isoformat()))
        ttk.Entry(
            range_bar,
            textvariable=start_var,
            width=13,
            justify="center",
        ).pack(side="right", padx=3, pady=8)
        tk.Label(
            range_bar,
            text="إلى تاريخ",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(12, 3), pady=8)
        end_var = tk.StringVar(value=str(saved_history.get("end_date") or end_default.isoformat()))
        ttk.Entry(
            range_bar,
            textvariable=end_var,
            width=13,
            justify="center",
        ).pack(side="right", padx=3, pady=8)
        period_label_var = tk.StringVar(value=str(saved_history.get("period_label") or "تقرير شهري"))
        count_var = tk.StringVar()
        quick_bar = tk.Frame(window, bg=Palette.PALE)
        quick_bar.pack(fill="x", padx=12, pady=(0, 4))
        views_bar = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        views_bar.pack(fill="x", padx=12, pady=(0, 4))
        filter_bar = tk.Frame(window, bg=Palette.WHITE, bd=1, relief="solid")
        filter_bar.pack(fill="x", padx=12, pady=(0, 4))
        search_var = tk.StringVar(value=str(saved_history.get("search") or ""))
        company_filter_var = tk.StringVar(value=str(saved_history.get("company") or "كل الشركات"))
        category_filter_var = tk.StringVar(value=str(saved_history.get("category") or "كل الفئات"))
        person_filter_var = tk.StringVar(value=str(saved_history.get("person") or "كل الأسماء"))
        status_filter_var = tk.StringVar(value=str(saved_history.get("status") or "كل الحالات"))
        filter_values = self.database.list_filter_values()
        company_filter_map = {
            "كل الشركات": "",
            **{
                str(company["name"]): str(company["id"])
                for company in self.database.list_companies()
            },
        }
        category_filter_map = {
            "كل الفئات": "",
            **{
                str(item["label"]): str(item["id"])
                for item in filter_values["categories"]
            },
        }
        status_filter_map = {
            "كل الحالات": "",
            **{label: key for key, label in REPORT_STATUSES.items()},
        }
        people_filter_values = ["كل الأسماء", *[item["label"] for item in filter_values["people"]]]
        if company_filter_var.get() not in company_filter_map:
            company_filter_var.set("كل الشركات")
        if category_filter_var.get() not in category_filter_map:
            category_filter_var.set("كل الفئات")
        if person_filter_var.get() not in people_filter_values:
            person_filter_var.set("كل الأسماء")
        if status_filter_var.get() not in status_filter_map:
            status_filter_var.set("كل الحالات")
        tk.Label(
            filter_bar,
            text="بحث سريع",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        ).pack(side="right", padx=(8, 2), pady=7)
        search_entry = ttk.Entry(
            filter_bar,
            textvariable=search_var,
            width=24,
            justify="right",
        )
        search_entry.pack(side="right", padx=2, pady=7)
        for label, variable, values, width in (
            ("الشركة", company_filter_var, list(company_filter_map), 18),
            ("الفئة", category_filter_var, list(category_filter_map), 18),
            (
                "الاسم",
                person_filter_var,
                people_filter_values,
                18,
            ),
            ("الحالة", status_filter_var, list(status_filter_map), 14),
        ):
            tk.Label(
                filter_bar,
                text=label,
                bg=Palette.WHITE,
                fg=Palette.NAVY,
                font=("Arial", 9, "bold"),
            ).pack(side="right", padx=(7, 2), pady=7)
            ttk.Combobox(
                filter_bar,
                textvariable=variable,
                values=values,
                state="readonly",
                width=width,
                justify="right",
            ).pack(side="right", padx=2, pady=7)
        history_actions = tk.Frame(window, bg=Palette.PALE)
        history_actions.pack(side="bottom", fill="x", padx=12, pady=(4, 10))
        workflow_actions = tk.Frame(window, bg=Palette.PALE)
        workflow_actions.pack(side="bottom", fill="x", padx=12, pady=(2, 2))

        def range_button(
            parent: tk.Misc,
            text: str,
            command: Any,
            color: str,
        ) -> None:
            tk.Button(
                parent,
                text=text,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 9, "bold"),
                padx=9,
                pady=5,
            ).pack(side="right", padx=2)

        tree_frame = tk.Frame(window, bg=Palette.WHITE)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=4)
        tree = ttk.Treeview(
            tree_frame,
            columns=("date", "company", "user", "code", "status", "locked", "count", "updated"),
            show="headings",
            height=12,
            selectmode="extended",
        )
        for column, title, width in (
            ("date", "تاريخ العمل", 130),
            ("company", "الشركة", 210),
            ("user", "المستخدم", 150),
            ("code", "الرمز", 110),
            ("status", "الحالة", 90),
            ("locked", "القفل", 70),
            ("count", "السجلات", 85),
            ("updated", "آخر حفظ", 175),
        ):
            tree.heading(column, text=title)
            tree.column(
                column,
                width=width,
                anchor="e" if column in {"company", "user"} else "center",
            )
        tree_scroll = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        details = tk.Text(
            window,
            wrap="word",
            font=("Arial", 11),
            bg=Palette.WHITE,
            fg=Palette.INK,
            height=9,
            padx=10,
            pady=10,
        )
        details.pack(fill="x", padx=12, pady=(4, 8))

        def validated_range() -> tuple[str, str] | None:
            start = parse_iso_date(start_var.get())
            end = parse_iso_date(end_var.get())
            if start is None or end is None:
                messagebox.showwarning(
                    "تاريخ غير صحيح",
                    "اكتب التاريخ بصيغة YYYY-MM-DD.",
                    parent=window,
                )
                return None
            if start > end:
                messagebox.showwarning(
                    "نطاق غير صحيح",
                    "تاريخ البداية يجب أن يسبق تاريخ النهاية.",
                    parent=window,
                )
                return None
            return start.isoformat(), end.isoformat()

        def refresh_records() -> None:
            selected_range = validated_range()
            if selected_range is None:
                return
            for item in tree.get_children():
                tree.delete(item)
            records = self.database.list_reports(
                *selected_range,
                search=search_var.get(),
                company_id=company_filter_map.get(company_filter_var.get(), ""),
                category_id=category_filter_map.get(category_filter_var.get(), ""),
                person_name=(
                    "" if person_filter_var.get() == "كل الأسماء"
                    else person_filter_var.get()
                ),
                status=status_filter_map.get(status_filter_var.get(), ""),
            )
            self.config["app_state"]["history_filters"] = {
                "start_date": selected_range[0],
                "end_date": selected_range[1],
                "period_label": period_label_var.get(),
                "search": search_var.get(),
                "company": company_filter_var.get(),
                "category": category_filter_var.get(),
                "person": person_filter_var.get(),
                "status": status_filter_var.get(),
            }
            try:
                atomic_write_json(self.config_path, self.config)
            except OSError:
                pass
            for index, record in enumerate(records, 1):
                tree.insert(
                    "",
                    "end",
                    iid=record["id"],
                    values=(
                        record["report_date"],
                        record["company_name"],
                        record["user_name"],
                        record["code"],
                        REPORT_STATUSES.get(record["status"], record["status"]),
                        "مقفل" if record["is_locked"] else "مفتوح",
                        record["row_count"],
                        record["updated_at"].replace("T", " "),
                    ),
                    tags=("even" if index % 2 == 0 else "odd",),
                )
            tree.tag_configure("odd", background=Palette.WHITE)
            tree.tag_configure("even", background=Palette.ALT_ROW)
            count_var.set(f"{period_label_var.get()} - {len(records)} تقرير محفوظ")
            details.configure(state="normal")
            details.delete("1.0", "end")
            details.insert(
                "end",
                "يمكن تحديد تقرير واحد أو عدة تقارير من الجدول. "
                "إذا لم تحدد شيئًا فسيتم إصدار التقرير عن كامل الفترة المعروضة.",
            )
            details.configure(state="disabled")

        def set_week() -> None:
            start, end = week_range(date.today())
            start_var.set(start.isoformat())
            end_var.set(end.isoformat())
            period_label_var.set("تقرير أسبوعي")
            refresh_records()

        def set_month() -> None:
            start, end = month_range(date.today())
            start_var.set(start.isoformat())
            end_var.set(end.isoformat())
            period_label_var.set("تقرير شهري")
            refresh_records()

        last_days_var = tk.StringVar(value="7")

        def set_last_days() -> None:
            days = coerce_nonnegative_int(last_days_var.get())
            if days < 1 or days > 3650:
                messagebox.showwarning(
                    "عدد غير صحيح",
                    "أدخل عدد الأيام من 1 إلى 3650.",
                    parent=window,
                )
                return
            end = date.today()
            start = date.fromordinal(end.toordinal() - days + 1)
            start_var.set(start.isoformat())
            end_var.set(end.isoformat())
            period_label_var.set(f"تقرير آخر {days} أيام")
            refresh_records()

        def set_all() -> None:
            records = self.database.list_reports()
            if records:
                dates = sorted(record["report_date"] for record in records)
                start_var.set(dates[0])
                end_var.set(dates[-1])
            else:
                today = date.today().isoformat()
                start_var.set(today)
                end_var.set(today)
            period_label_var.set("تقرير جميع البيانات")
            refresh_records()

        def apply_custom_range() -> None:
            period_label_var.set("تقرير فترة مخصصة")
            refresh_records()

        range_button(
            range_bar,
            "تطبيق الفترة من/إلى",
            apply_custom_range,
            Palette.INFO,
        )
        tk.Label(
            quick_bar,
            text="آخر عدد من الأيام:",
            bg=Palette.PALE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
        ).pack(side="right", padx=(4, 3), pady=4)
        ttk.Spinbox(
            quick_bar,
            from_=1,
            to=3650,
            textvariable=last_days_var,
            width=7,
            justify="center",
        ).pack(side="right", padx=3, pady=4)
        range_button(quick_bar, "عرض آخر الأيام", set_last_days, Palette.INFO)
        range_button(quick_bar, "هذا الأسبوع", set_week, Palette.GREEN)
        range_button(quick_bar, "هذا الشهر", set_month, Palette.BLUE)
        range_button(quick_bar, "كل البيانات", set_all, Palette.MUTED)
        tk.Label(
            range_bar,
            textvariable=count_var,
            bg=Palette.WHITE,
            fg=Palette.MUTED,
            font=("Arial", 9, "bold"),
        ).pack(side="left", padx=8)
        range_button(filter_bar, "بحث / تطبيق الفلاتر", refresh_records, Palette.INFO)

        def reset_filters() -> None:
            search_var.set("")
            company_filter_var.set("كل الشركات")
            category_filter_var.set("كل الفئات")
            person_filter_var.set("كل الأسماء")
            status_filter_var.set("كل الحالات")
            refresh_records()

        range_button(filter_bar, "مسح الفلاتر", reset_filters, Palette.MUTED)
        saved_view_var = tk.StringVar()
        saved_views: list[dict[str, Any]] = []
        saved_view_combo = ttk.Combobox(
            views_bar,
            textvariable=saved_view_var,
            state="readonly",
            width=34,
            justify="right",
        )
        tk.Label(
            views_bar,
            text="العروض المحفوظة",
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 9, "bold"),
        ).pack(side="right", padx=(8, 3), pady=7)
        saved_view_combo.pack(side="right", padx=3, pady=7)

        def current_view_filters() -> dict[str, str]:
            return {
                "start_date": start_var.get(),
                "end_date": end_var.get(),
                "period_label": period_label_var.get(),
                "search": search_var.get(),
                "company": company_filter_var.get(),
                "category": category_filter_var.get(),
                "person": person_filter_var.get(),
                "status": status_filter_var.get(),
            }

        def refresh_saved_views() -> None:
            nonlocal saved_views
            saved_views = self.database.list_views("history", self.current_user)
            names = [str(item.get("name", "")) for item in saved_views]
            saved_view_combo.configure(values=names)
            if saved_view_var.get() not in names:
                saved_view_var.set("")

        def save_current_view() -> None:
            name = simpledialog.askstring(
                "حفظ العرض", "اسم العرض:", parent=window
            )
            if not name:
                return
            shared = False
            if str(self.current_user.get("role")) == "admin":
                shared = messagebox.askyesno(
                    "عرض مشترك", "إتاحة هذا العرض لجميع المستخدمين؟", parent=window
                )
            make_default = messagebox.askyesno(
                "العرض الافتراضي", "استخدام هذا العرض كافتراضي؟", parent=window
            )
            try:
                self.database.save_view(
                    "history",
                    name,
                    current_view_filters(),
                    self.current_user,
                    is_shared=shared,
                    is_default=make_default,
                )
            except (ValueError, PermissionError, sqlite3.Error) as error:
                messagebox.showerror("تعذر الحفظ", str(error), parent=window)
                return
            refresh_saved_views()
            saved_view_var.set(name.strip())

        def apply_saved_view(_event: Any = None) -> None:
            selected = next(
                (item for item in saved_views if item.get("name") == saved_view_var.get()),
                None,
            )
            if not selected:
                return
            values = selected.get("filters", {})
            if not isinstance(values, dict):
                return
            start_var.set(str(values.get("start_date") or start_var.get()))
            end_var.set(str(values.get("end_date") or end_var.get()))
            period_label_var.set(str(values.get("period_label") or "عرض محفوظ"))
            search_var.set(str(values.get("search") or ""))
            company_filter_var.set(
                str(values.get("company"))
                if str(values.get("company")) in company_filter_map
                else "كل الشركات"
            )
            category_filter_var.set(
                str(values.get("category"))
                if str(values.get("category")) in category_filter_map
                else "كل الفئات"
            )
            person_filter_var.set(
                str(values.get("person"))
                if str(values.get("person")) in people_filter_values
                else "كل الأسماء"
            )
            status_filter_var.set(
                str(values.get("status"))
                if str(values.get("status")) in status_filter_map
                else "كل الحالات"
            )
            refresh_records()

        def delete_saved_view() -> None:
            selected = next(
                (item for item in saved_views if item.get("name") == saved_view_var.get()),
                None,
            )
            if not selected:
                return
            try:
                self.database.delete_view(str(selected["id"]), self.current_user)
            except (ValueError, PermissionError, sqlite3.Error) as error:
                messagebox.showerror("تعذر الحذف", str(error), parent=window)
                return
            saved_view_var.set("")
            refresh_saved_views()

        range_button(views_bar, "حفظ العرض الحالي", save_current_view, Palette.SUCCESS)
        range_button(views_bar, "تطبيق", apply_saved_view, Palette.INFO)
        range_button(views_bar, "حذف", delete_saved_view, Palette.DANGER)
        saved_view_combo.bind("<<ComboboxSelected>>", apply_saved_view)
        refresh_saved_views()
        search_entry.bind("<Return>", lambda _event: refresh_records())

        def show_selected(_event: Any = None) -> None:
            details.configure(state="normal")
            details.delete("1.0", "end")
            selection = list(tree.selection())
            if not selection:
                details.insert(
                    "end",
                    "لم يتم تحديد تقرير؛ ستستخدم أزرار التصدير جميع بيانات الفترة.",
                )
                details.configure(state="disabled")
                return
            if len(selection) > 1:
                details.insert(
                    "end",
                    f"تم تحديد {len(selection)} تقارير. "
                    "يمكن إصدار تقرير واحد يجمع هذه الأيام فقط أو حذفها.",
                )
                details.configure(state="disabled")
                return
            record = self.database.get_report(selection[0])
            if not record:
                details.insert("end", "تعذر قراءة التقرير المحدد.")
                details.configure(state="disabled")
                return
            report = record["snapshot"]
            details.insert(
                "end",
                (
                    f"التاريخ: {report.get('date', '')} - {report.get('day', '')}\n"
                    f"الشركة: {report.get('company_name', '')}\n"
                    f"المستخدم: {report.get('user_name', '')}\n"
                    f"الرمز: {report.get('code', '')}\n"
                    f"الحالة: {REPORT_STATUSES.get(record.get('status', 'draft'), 'مسودة')}\n"
                    f"القفل: {'معتمد ومقفل' if record.get('is_locked') else 'مفتوح'}\n"
                    f"عدد السجلات: {len(report.get('rows', []))}\n\n"
                ),
            )
            for index, row in enumerate(report.get("rows", []), 1):
                fields_text = " | ".join(
                    f"{field.get('label')}: {field.get('value')}"
                    for field in row.get("fields", [])
                )
                details.insert(
                    "end",
                    (
                        f"{index}. {row.get('category_label', '')} - "
                        f"{row.get('name', '')}\n"
                        f"   {fields_text} | الإجمالي: {row.get('total', 0)}\n\n"
                    ),
                )
            details.configure(state="disabled")

        def selected_snapshots() -> tuple[list[dict[str, Any]], str, bool]:
            selected_ids = list(tree.selection())
            explicit_selection = bool(selected_ids)
            target_ids = selected_ids or list(tree.get_children())
            snapshots: list[dict[str, Any]] = []
            for report_id in target_ids:
                record = self.database.get_report(report_id)
                if record:
                    snapshots.append(record["snapshot"])
            label = (
                "تقرير أيام محددة"
                if explicit_selection and len(snapshots) > 1
                else period_label_var.get()
            )
            return snapshots, label, explicit_selection

        def snapshot_for_export() -> dict[str, Any] | None:
            snapshots, label, explicit_selection = selected_snapshots()
            if not snapshots:
                messagebox.showwarning(
                    "لا توجد بيانات",
                    "لا توجد تقارير محفوظة ضمن الاختيار الحالي.",
                    parent=window,
                )
                return None
            if explicit_selection and len(snapshots) == 1:
                return snapshots[0]
            try:
                return combine_report_snapshots(snapshots, self.config, label)
            except ValueError as error:
                messagebox.showerror("تعذر إعداد التقرير", str(error), parent=window)
                return None

        def load_selected_for_editing() -> None:
            if not self.require_permission("edit"):
                return
            selection = list(tree.selection())
            if len(selection) != 1:
                messagebox.showwarning(
                    "اختر تقريرًا واحدًا",
                    "حدد تقريرًا واحدًا فقط لتحميله والتعديل عليه.",
                    parent=window,
                )
                return
            record = self.database.get_report(selection[0])
            if not record:
                messagebox.showerror(
                    "خطأ",
                    "تعذر قراءة التقرير المحدد.",
                    parent=window,
                )
                return
            self.load_snapshot_for_editing(record["id"], record["snapshot"], record)
            window.destroy()
            messagebox.showinfo(
                "تم التحميل",
                "تم تحميل التقرير في الشاشة الرئيسية.\n"
                "عدّل السجلات ثم اضغط «حفظ في قاعدة البيانات» لتحديثه.",
            )

        def delete_selected_history() -> None:
            if not self.require_permission("delete"):
                return
            selected_ids = list(tree.selection())
            if not selected_ids:
                messagebox.showwarning(
                    "تنبيه",
                    "حدد تقريرًا واحدًا أو أكثر للحذف.",
                    parent=window,
                )
                return
            if not messagebox.askyesno(
                "تأكيد الأرشفة",
                f"نقل {len(selected_ids)} تقرير إلى المحذوفات؟",
                parent=window,
            ):
                return
            try:
                deleted = self.database.archive_reports(
                    selected_ids,
                    actor=self.actor_name,
                )
            except (sqlite3.Error, PermissionError) as error:
                messagebox.showerror("خطأ", str(error), parent=window)
                return
            if self.loaded_report_id in selected_ids:
                self.loaded_report_id = None
            self.database_status_var.set(
                f"قاعدة البيانات: {self.database.count()} تقرير"
            )
            refresh_records()
            messagebox.showinfo(
                "تمت الأرشفة",
                f"تم نقل {deleted} تقرير إلى المحذوفات ويمكن استعادته لاحقًا.",
                parent=window,
            )

        def toggle_selected_lock() -> None:
            if not self.require_permission("approve"):
                return
            selection = list(tree.selection())
            if len(selection) != 1:
                messagebox.showwarning(
                    "اختر تقريراً واحداً",
                    "حدد تقريراً واحداً لاعتماده وقفلِه أو لإلغاء القفل.",
                    parent=window,
                )
                return
            record = self.database.get_report(selection[0])
            if not record:
                return
            should_lock = not bool(record.get("is_locked"))
            action_label = "اعتماد وقفل" if should_lock else "إلغاء قفل"
            if not messagebox.askyesno(
                action_label,
                f"{action_label} التقرير المحدد؟",
                parent=window,
            ):
                return
            self.database.set_report_lock(
                selection[0],
                should_lock,
                self.actor_name,
            )
            if self.loaded_report_id == selection[0]:
                self.loaded_report_locked = should_lock
            refresh_records()

        def open_selected_workflow() -> None:
            selection = list(tree.selection())
            if len(selection) != 1:
                messagebox.showwarning(
                    "اختر تقريرًا واحدًا",
                    "حدد تقريرًا واحدًا لعرض إصداراته وموافقاته وبصمة التحقق.",
                    parent=window,
                )
                return
            self.show_report_workflow(selection[0], refresh_records)

        def export_selected_sync() -> None:
            selected_ids = list(tree.selection()) or list(tree.get_children())
            self.export_sync_reports(selected_ids, window)

        def import_sync_and_refresh() -> None:
            self.import_sync_reports(window, refresh_records)

        def export_period_pdf() -> None:
            snapshot = snapshot_for_export()
            if snapshot is None:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ التقرير من قاعدة البيانات بصيغة PDF",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"تقرير_دوري_{date.today().isoformat()}.pdf",
            )
            if not filename:
                return
            try:
                output = export_pdf_to_path(snapshot, filename)
            except (OSError, RuntimeError, ValueError) as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            messagebox.showinfo("تم التصدير", f"تم إنشاء:\n{output}", parent=window)

        def export_period_excel() -> None:
            snapshot = snapshot_for_export()
            if snapshot is None:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ تقرير الفترة بصيغة Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"تقرير_دوري_{date.today().isoformat()}.xlsx",
            )
            if not filename:
                return
            try:
                output = export_snapshot_excel_to_path(snapshot, filename)
            except (OSError, RuntimeError, ValueError) as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            self.database.audit(
                self.actor_name,
                "export_excel",
                "period_report",
                "",
                {"path": str(output), "period": period_label_var.get()},
            )
            messagebox.showinfo("تم التصدير", f"تم إنشاء:\n{output}", parent=window)

        def export_period_images() -> None:
            snapshot = snapshot_for_export()
            if snapshot is None:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ التقرير من قاعدة البيانات كصور",
                defaultextension=".png",
                filetypes=[("PNG", "*.png")],
                initialfile=f"تقرير_دوري_{date.today().isoformat()}.png",
            )
            if not filename:
                return
            try:
                outputs = export_images_to_path(snapshot, filename)
            except (OSError, RuntimeError, ValueError) as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            messagebox.showinfo(
                "تم التصدير",
                f"تم إنشاء {len(outputs)} صورة.",
                parent=window,
            )

        def export_period_csv() -> None:
            snapshot = snapshot_for_export()
            if snapshot is None:
                return
            filename = filedialog.asksaveasfilename(
                parent=window,
                title="حفظ بيانات الفترة بصيغة CSV",
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")],
                initialfile=f"بيانات_دورية_{date.today().isoformat()}.csv",
            )
            if not filename:
                return
            try:
                output = export_csv_to_path(snapshot, filename)
            except OSError as error:
                messagebox.showerror("تعذر التصدير", str(error), parent=window)
                return
            messagebox.showinfo("تم التصدير", f"تم إنشاء:\n{output}", parent=window)

        tree.bind("<<TreeviewSelect>>", show_selected)
        for text_value, command, color in (
            ("الإصدارات والموافقات", open_selected_workflow, Palette.PRIMARY),
            ("تصدير حزمة مزامنة", export_selected_sync, Palette.INFO),
            ("استيراد حزمة مزامنة", import_sync_and_refresh, Palette.SUCCESS),
        ):
            tk.Button(
                workflow_actions,
                text=text_value,
                command=command,
                bg=color,
                fg=Palette.WHITE,
                relief="flat",
                font=("Arial", 9, "bold"),
                padx=12,
                pady=6,
            ).pack(side="right", padx=3)
        tk.Button(
            history_actions,
            text="تحميل للتعديل",
            command=load_selected_for_editing,
            bg=Palette.GREEN,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="right", padx=3)
        tk.Button(
            history_actions,
            text="حذف المحدد",
            command=delete_selected_history,
            bg=Palette.RED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="right", padx=3)
        tk.Button(
            history_actions,
            text="اعتماد / إلغاء القفل",
            command=toggle_selected_lock,
            bg=Palette.ACCENT,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="right", padx=3)
        tk.Button(
            history_actions,
            text="PDF للفترة/المحدد",
            command=export_period_pdf,
            bg=Palette.RED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="left", padx=3)
        tk.Button(
            history_actions,
            text="صور للفترة/المحدد",
            command=export_period_images,
            bg=Palette.ORANGE,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="left", padx=3)
        tk.Button(
            history_actions,
            text="Excel للفترة/المحدد",
            command=export_period_excel,
            bg=Palette.SUCCESS,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="left", padx=3)
        tk.Button(
            history_actions,
            text="CSV للفترة/المحدد",
            command=export_period_csv,
            bg=Palette.BLUE,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="left", padx=3)
        tk.Button(
            history_actions,
            text="إغلاق",
            command=window.destroy,
            bg=Palette.MUTED,
            fg=Palette.WHITE,
            relief="flat",
            font=("Arial", 10, "bold"),
            padx=12,
            pady=6,
        ).pack(side="left", padx=3)
        refresh_records()

    def show_report_workflow(
        self, report_id: str, on_changed: Any = None
    ) -> None:
        window = tk.Toplevel(self.root)
        window.title("إصدارات التقرير والموافقات والتحقق")
        window.geometry("1080x700")
        window.minsize(900, 600)
        window.configure(bg=Palette.PALE)
        window.transient(self.root)
        tk.Label(
            window,
            text="مركز دورة حياة التقرير",
            bg=Palette.NAVY,
            fg=Palette.WHITE,
            font=("Arial", 17, "bold"),
            pady=11,
        ).pack(fill="x")
        verification_var = tk.StringVar(value="جارٍ التحقق...")
        verification_label = tk.Label(
            window,
            textvariable=verification_var,
            bg=Palette.WHITE,
            fg=Palette.NAVY,
            font=("Arial", 10, "bold"),
            anchor="e",
            padx=12,
            pady=9,
        )
        verification_label.pack(fill="x", padx=12, pady=(10, 5))

        content = tk.Frame(window, bg=Palette.PALE)
        content.pack(fill="both", expand=True, padx=12, pady=5)
        content.grid_columnconfigure(0, weight=1, uniform="workflow")
        content.grid_columnconfigure(1, weight=1, uniform="workflow")
        content.grid_rowconfigure(0, weight=1)

        revisions_box = tk.LabelFrame(
            content, text="الإصدارات المحفوظة", bg=Palette.WHITE,
            fg=Palette.NAVY, font=("Arial", 10, "bold"), padx=6, pady=6,
        )
        revisions_box.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        revision_tree = ttk.Treeview(
            revisions_box,
            columns=("revision", "changed_at", "changed_by", "hash"),
            show="headings",
            height=15,
        )
        for column, title, width in (
            ("revision", "الإصدار", 75),
            ("changed_at", "التاريخ", 145),
            ("changed_by", "المستخدم", 120),
            ("hash", "البصمة", 140),
        ):
            revision_tree.heading(column, text=title)
            revision_tree.column(column, width=width, anchor="center")
        revision_tree.pack(fill="both", expand=True)

        approvals_box = tk.LabelFrame(
            content, text="مسار الموافقات", bg=Palette.WHITE,
            fg=Palette.NAVY, font=("Arial", 10, "bold"), padx=6, pady=6,
        )
        approvals_box.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        approval_tree = ttk.Treeview(
            approvals_box,
            columns=("order", "stage", "decision", "user", "date"),
            show="headings",
            height=15,
        )
        for column, title, width in (
            ("order", "م", 45),
            ("stage", "المرحلة", 150),
            ("decision", "القرار", 80),
            ("user", "بواسطة", 100),
            ("date", "التاريخ", 135),
        ):
            approval_tree.heading(column, text=title)
            approval_tree.column(column, width=width, anchor="center")
        approval_tree.pack(fill="both", expand=True)

        decision_labels = {
            "pending": "بانتظار القرار",
            "approved": "معتمد",
            "rejected": "مرفوض",
        }

        def refresh() -> None:
            verification = self.database.verify_report(report_id)
            verification_var.set(
                f"الإصدار الحالي: {verification.get('revision', 0)} | "
                f"رمز التحقق: {verification.get('short_code') or 'غير متاح'} | "
                f"{verification.get('reason', '')}"
            )
            verification_label.configure(
                fg=Palette.GREEN if verification.get("valid") else Palette.RED
            )
            revision_tree.delete(*revision_tree.get_children())
            for revision in self.database.list_report_revisions(report_id):
                number = int(revision["revision_number"])
                revision_tree.insert(
                    "", "end", iid=str(number),
                    values=(
                        number,
                        str(revision.get("changed_at", "")).replace("T", " "),
                        revision.get("changed_by", ""),
                        str(revision.get("content_hash", ""))[:16].upper(),
                    ),
                )
            approval_tree.delete(*approval_tree.get_children())
            for step in self.database.list_report_approvals(report_id):
                order = int(step["stage_order"])
                approval_tree.insert(
                    "", "end", iid=str(order),
                    values=(
                        order,
                        step.get("stage_name", ""),
                        decision_labels.get(str(step.get("decision", "")), ""),
                        step.get("decided_by", ""),
                        str(step.get("decided_at", "")).replace("T", " "),
                    ),
                )
            if callable(on_changed):
                on_changed()

        def configure_workflow() -> None:
            if not self.require_permission("manage"):
                return
            value = simpledialog.askstring(
                "إعداد مسار الموافقات",
                "اكتب أسماء المراحل مفصولة بفاصلة، حسب ترتيبها:",
                initialvalue="مراجعة المشرف، اعتماد مدير النظام",
                parent=window,
            )
            if value is None:
                return
            stages = re.split(r"[,،\n]", value)
            try:
                self.database.configure_approval_workflow(report_id, stages, self.actor_name)
            except (ValueError, sqlite3.Error) as error:
                messagebox.showerror("تعذر إعداد المسار", str(error), parent=window)
                return
            refresh()

        def decide(decision: str) -> None:
            if not self.require_permission("approve"):
                return
            selection = approval_tree.selection()
            if len(selection) != 1:
                messagebox.showwarning("اختر مرحلة", "حدد مرحلة موافقة واحدة.", parent=window)
                return
            note = simpledialog.askstring(
                "ملاحظة القرار", "ملاحظة اختيارية:", parent=window
            )
            if note is None:
                return
            try:
                self.database.decide_report_approval(
                    report_id, int(selection[0]), decision, self.actor_name, note
                )
            except (ValueError, PermissionError, sqlite3.Error) as error:
                messagebox.showerror("تعذر تسجيل القرار", str(error), parent=window)
                return
            record = self.database.get_report(report_id)
            if record and self.loaded_report_id == report_id:
                self.loaded_report_locked = bool(record.get("is_locked"))
            refresh()

        def restore_revision() -> None:
            if not self.require_permission("recover"):
                return
            selection = revision_tree.selection()
            if len(selection) != 1:
                messagebox.showwarning("اختر إصدارًا", "حدد إصدارًا واحدًا لاستعادته.", parent=window)
                return
            if not messagebox.askyesno(
                "استعادة الإصدار",
                "سيُنشئ النظام إصدارًا جديدًا من النسخة المحددة ويفتح التقرير للتعديل. متابعة؟",
                parent=window,
            ):
                return
            try:
                self.database.restore_report_revision(
                    report_id, int(selection[0]), self.actor_name
                )
            except (ValueError, sqlite3.Error) as error:
                messagebox.showerror("تعذر الاستعادة", str(error), parent=window)
                return
            if self.loaded_report_id == report_id:
                self.loaded_report_locked = False
            refresh()

        buttons = tk.Frame(window, bg=Palette.PALE)
        buttons.pack(fill="x", padx=12, pady=(4, 10))
        for text_value, command, color in (
            ("إعداد المراحل", configure_workflow, Palette.PRIMARY),
            ("اعتماد المرحلة", lambda: decide("approved"), Palette.GREEN),
            ("رفض المرحلة", lambda: decide("rejected"), Palette.RED),
            ("استعادة الإصدار", restore_revision, Palette.INFO),
            ("إعادة التحقق", refresh, Palette.ACCENT),
            ("إغلاق", window.destroy, Palette.MUTED),
        ):
            tk.Button(
                buttons, text=text_value, command=command, bg=color, fg=Palette.WHITE,
                relief="flat", font=("Arial", 9, "bold"), padx=11, pady=6,
            ).pack(side="right", padx=3)
        refresh()

    def export_sync_reports(self, report_ids: list[str], parent: tk.Misc) -> None:
        if not self.require_permission("sync"):
            return
        filename = filedialog.asksaveasfilename(
            parent=parent,
            title="حفظ حزمة المزامنة الآمنة",
            defaultextension=".dwrsync",
            filetypes=[("Daily Work Reports Sync", "*.dwrsync"), ("ZIP", "*.zip")],
            initialfile=f"DailyWorkReports_Sync_{datetime.now():%Y%m%d_%H%M%S}.dwrsync",
        )
        if not filename:
            return
        try:
            output = self.database.export_sync_package(filename, report_ids, self.actor_name)
        except (OSError, ValueError, sqlite3.Error, zipfile.BadZipFile) as error:
            messagebox.showerror("تعذر إنشاء الحزمة", str(error), parent=parent)
            return
        messagebox.showinfo(
            "تم إنشاء الحزمة",
            f"تم إنشاء حزمة مزامنة متحققة البصمة:\n{output}",
            parent=parent,
        )

    def import_sync_reports(self, parent: tk.Misc, on_changed: Any = None) -> None:
        if not self.require_permission("sync"):
            return
        filename = filedialog.askopenfilename(
            parent=parent,
            title="استيراد حزمة مزامنة",
            filetypes=[("Daily Work Reports Sync", "*.dwrsync *.zip"), ("كل الملفات", "*.*")],
        )
        if not filename:
            return
        try:
            summary = self.database.import_sync_package(filename, self.actor_name)
        except (OSError, ValueError, sqlite3.Error, zipfile.BadZipFile, json.JSONDecodeError) as error:
            messagebox.showerror("تعذر استيراد الحزمة", str(error), parent=parent)
            return
        if callable(on_changed):
            on_changed()
        self.database_status_var.set(f"قاعدة البيانات: {self.database.count()} تقرير")
        messagebox.showinfo(
            "نتيجة المزامنة",
            "\n".join(
                (
                    f"جديدة: {summary.get('created', 0)}",
                    f"محدّثة: {summary.get('updated', 0)}",
                    f"دون تغيير: {summary.get('unchanged', 0)}",
                    f"تعارضات محفوظة دون استبدال: {summary.get('conflicts', 0)}",
                    "تم التعرف على الحزمة سابقًا." if summary.get("duplicate_package") else "تم التحقق من بصمة الحزمة.",
                )
            ),
            parent=parent,
        )

    def clear_report(self) -> None:
        if self.entries and not messagebox.askyesno(
            "تقرير جديد",
            "بدء تقرير جديد ومسح السجلات الحالية من الشاشة؟\n"
            "لن تُحذف التقارير المحفوظة من قاعدة البيانات.",
        ):
            return
        self.entries.clear()
        self.draft_store.clear()
        self.loaded_report_id = None
        self.loaded_report_locked = False
        self.status_var.set(REPORT_STATUSES["draft"])
        self.database_status_var.set(
            f"تقرير جديد | قاعدة البيانات: {self.database.count()} تقرير"
        )
        self.refresh_entries_tree()
        self.clear_editor()

    def on_close(self) -> None:
        if self._draft_save_job is not None:
            try:
                self.root.after_cancel(self._draft_save_job)
            except tk.TclError:
                pass
        self._save_current_draft()
        self.save_config()
        self.root.destroy()


def run_self_test() -> None:
    config = normalize_config(DEFAULT_CONFIG)
    assert arabic_day_name(date(2026, 7, 27)) == "الإثنين"
    assert config["categories"]["drawers"]["people"] == ["موظف اختبار"]
    assert [
        field["id"] for field in config["categories"]["drawers"]["fields"]
    ] == ["colored", "edits", "black_white"]
    assert [
        field["id"] for field in config["categories"]["technicians"]["fields"]
    ] == ["new_measure", "confirmed_measure", "outlets", "cancelled", "new_requests", "notes"]
    assert [
        field["id"] for field in config["categories"]["showroom_employees"]["fields"]
    ] == ["measurements_count", "drawings_count"]
    assert config["categories"]["showroom_employees"]["person_label"] == "اسم الموظف"
    assert KitchensERP._validate_numeric_text("125", "number")
    assert KitchensERP._validate_numeric_text("", "number")
    assert not KitchensERP._validate_numeric_text("12-", "number")
    assert not KitchensERP._validate_numeric_text("1.5", "number")
    assert config["identity"]["report_title"] == APP_NAME
    custom_category_id = "category_quality"
    config["category_order"].append(custom_category_id)
    config["categories"][custom_category_id] = {
        "label": "الجودة",
        "person_label": "المراجع",
        "total_label": "إجمالي الفحص",
        "people": ["موظف اختبار"],
        "fields": [
            {
                "id": "quality_status",
                "label": "الحالة",
                "kind": "choice",
                "options": ["مقبول", "يحتاج مراجعة"],
            }
        ],
    }
    config = normalize_config(config)
    assert custom_category_id in category_ids(config)
    assert config["categories"][custom_category_id]["fields"][0]["kind"] == "choice"
    sample = new_entry(
        "drawers",
        "موظف اختبار",
        {"black_white": 2, "colored": 3, "edits": 1},
    )
    assert calculate_total(config, sample) == 6
    snapshot = build_report_snapshot(
        config,
        [sample],
        date(2026, 7, 27),
        "MCO",
        "1",
    )
    assert snapshot["day"] == "الإثنين"
    assert snapshot["rows"][0]["total"] == 6
    assert snapshot["report_title"] == APP_NAME
    later_snapshot = build_report_snapshot(
        config,
        [sample],
        date(2026, 7, 28),
        "MCO",
        "2",
    )
    combined = combine_report_snapshots(
        [snapshot, later_snapshot],
        config,
        "تقرير أسبوعي",
    )
    assert len(combined["rows"]) == 2
    assert combined["rows"][0]["fields"][0]["label"] == "تاريخ العمل"
    assert "Daily Work Report" in combined["report_title"]
    assert rtl_visual("تقرير") != "تقرير"[::-1]
    assert (
        "2026-07-01 - 2026-07-10"
        in rtl_visual("الفترة: 2026-07-01 - 2026-07-10")
    )
    print("SELF TEST OK")


def run_database_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    database = ReportDatabase(output_dir / f"database_test_{uuid.uuid4().hex}.db")
    config = normalize_config(DEFAULT_CONFIG)
    first = build_report_snapshot(
        config,
        [new_entry("drawers", "موظف اختبار", {"black_white": 1})],
        date(2026, 7, 27),
        "DB-1",
        "1",
    )
    second = build_report_snapshot(
        config,
        [new_entry("drawers", "موظف اختبار", {"colored": 2})],
        date(2026, 7, 30),
        "DB-2",
        "2",
    )
    first_id = database.save_snapshot(first)
    second_id = database.save_snapshot(second)
    assert database.count() == 2
    assert len(database.list_reports("2026-07-27", "2026-07-30")) == 2
    first["code"] = "DB-1-UPDATED"
    database.save_snapshot(first, first_id)
    assert database.get_report(first_id)["snapshot"]["code"] == "DB-1-UPDATED"
    reports = [
        database.get_report(first_id)["snapshot"],
        database.get_report(second_id)["snapshot"],
    ]
    combined = combine_report_snapshots(reports, config, "تقرير أيام محددة")
    assert len(combined["rows"]) == 2
    assert database.delete_reports([first_id, second_id]) == 2
    assert database.count() == 0
    print("DATABASE SMOKE TEST OK")


def run_export_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    config = normalize_config(DEFAULT_CONFIG)
    config["identity"]["company_name"] = "شركة الاختبار"
    config["identity"]["user_name"] = "مستخدم الاختبار"
    entries = [
        new_entry(
            "drawers",
            "موظف اختبار",
            {"black_white": 2, "colored": 1, "edits": 1},
        )
    ]
    snapshot = build_report_snapshot(
        config,
        entries,
        date.today(),
        "TEST",
        "1",
    )
    export_pdf_to_path(snapshot, output_dir / "smoke_test.pdf")
    export_images_to_path(snapshot, output_dir / "smoke_test.png")
    export_csv_to_path(snapshot, output_dir / "smoke_test.csv")
    print("EXPORT SMOKE TEST OK")


def run_period_export_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    config = normalize_config(DEFAULT_CONFIG)
    config["identity"]["company_name"] = "شركة النموذج"
    config["identity"]["user_name"] = "مدير التقارير"
    snapshots: list[dict[str, Any]] = []
    for index in range(10):
        work_date = date(2026, 7, 1 + index)
        entries = [
            new_entry(
                "drawers",
                "موظف اختبار",
                {
                    "black_white": index + 1,
                    "colored": (index % 3) + 1,
                    "edits": index % 2,
                },
            )
        ]
        snapshots.append(
            build_report_snapshot(
                config,
                entries,
                work_date,
                f"PER-{index + 1:02d}",
                str(index + 1),
            )
        )
    combined = combine_report_snapshots(
        snapshots,
        config,
        "تقرير آخر 10 أيام",
    )
    export_pdf_to_path(combined, output_dir / "period_report.pdf")
    export_images_to_path(combined, output_dir / "period_report.png")
    export_csv_to_path(combined, output_dir / "period_report.csv")
    print("PERIOD EXPORT SMOKE TEST OK")


def run_pdf_design_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    config = normalize_config(DEFAULT_CONFIG)
    config["identity"]["company_name"] = "شركة النموذج للتصميم والتقارير"
    config["identity"]["user_name"] = "مدير إدارة التقارير"
    custom_category_id = "quality_follow_up"
    config["category_order"].append(custom_category_id)
    config["categories"][custom_category_id] = {
        "label": "متابعة الجودة وخدمة العملاء",
        "person_label": "المسؤول",
        "total_label": "إجمالي المتابعات",
        "people": ["مسؤول الجودة", "خدمة العملاء"],
        "fields": [
            {
                "id": "status",
                "label": "حالة الطلب",
                "kind": "choice",
                "options": ["مكتمل", "قيد المتابعة", "مؤجل"],
            },
            {
                "id": "project",
                "label": "اسم المشروع",
                "kind": "text",
                "options": [],
            },
            {
                "id": "details",
                "label": "تفاصيل المتابعة",
                "kind": "text",
                "options": [],
            },
            {
                "id": "visits",
                "label": "الزيارات",
                "kind": "number",
                "options": [],
            },
            {
                "id": "calls",
                "label": "الاتصالات",
                "kind": "number",
                "options": [],
            },
        ],
    }
    config = normalize_config(config)
    entries: list[dict[str, Any]] = []
    for index, name in enumerate(("موظف اختبار", "موظف اختبار", "موظف اختبار"), 1):
        entries.append(
            new_entry(
                "drawers",
                name,
                {
                    "black_white": index + 1,
                    "colored": index,
                    "edits": index - 1,
                },
            )
        )
    for index in range(7):
        entries.append(
            new_entry(
                "technicians",
                config["categories"]["technicians"]["people"][index % 3],
                {
                    "new_measure": index + 1,
                    "confirmed_measure": index % 3,
                    "outlets": index % 2,
                    "cancelled": 0,
                    "notes": (
                        "مراجعة المقاسات مع العميل والتأكد من مواقع الكهرباء "
                        f"قبل اعتماد الطلب رقم {index + 1}"
                    ),
                },
            )
        )
    for index in range(8):
        entries.append(
            new_entry(
                "showroom_employees",
                config["categories"]["showroom_employees"]["people"][index],
                {
                    "drawings_count": index + 1,
                    "measurements_count": index % 4,
                },
            )
        )
    for index in range(5):
        entries.append(
            new_entry(
                custom_category_id,
                "مسؤول الجودة" if index % 2 == 0 else "خدمة العملاء",
                {
                    "status": "مكتمل" if index % 2 == 0 else "قيد المتابعة",
                    "project": f"مشروع العميل رقم {index + 1}",
                    "details": (
                        "تم التواصل وتوثيق الملاحظات ومراجعة جميع البنود "
                        "مع القسم المختص قبل الإغلاق"
                    ),
                    "visits": index % 3,
                    "calls": index + 2,
                },
            )
        )
    snapshot = build_report_snapshot(
        config,
        entries,
        date(2026, 7, 30),
        "QA-2026-07",
        "27",
    )
    export_pdf_to_path(snapshot, output_dir / "pdf_design_test.pdf")
    print("PDF DESIGN SMOKE TEST OK")


def run_v9_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    test_dir = output_dir / f"v9_test_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=False)
    database = ReportDatabase(test_dir / "reports.db")
    config = normalize_config(DEFAULT_CONFIG)
    company_id, template_id = database.ensure_defaults(config)
    assert len(database.list_templates()) >= 3
    company_id = database.save_company(
        {
            "name": "شركة اختبار الإصدار التاسع",
            "address": "الرياض",
            "phone": "0110000000",
            "tax_number": "300000000000003",
            "active": True,
        },
        "admin",
        company_id,
    )
    snapshot = build_report_snapshot(
        config,
        [
            new_entry("drawers", "موظف اختبار", {"black_white": 3, "colored": 2}),
            new_entry(
                "technicians",
                "موظف اختبار",
                {"new_measure": 4, "notes": "مراجعة الموقع"},
            ),
        ],
        date(2026, 7, 30),
        "V9-001",
        "1",
        "شركة اختبار الإصدار التاسع",
        "admin",
        company_profile=database.get_company(company_id),
        company_id=company_id,
        template_id=template_id,
        template_name="القالب الرسمي",
        status="completed",
    )
    report_id = database.save_snapshot(
        snapshot,
        company_id=company_id,
        template_id=template_id,
        status="completed",
        actor="admin",
    )
    assert database.find_duplicate(snapshot, company_id) is not None
    assert len(database.list_reports(search="مراجعة الموقع")) == 1
    assert len(database.list_reports(category_id="drawers")) == 1
    assert len(database.list_reports(person_name="موظف اختبار")) == 1
    assert len(database.list_reports(status="completed")) == 1
    analytics = database.analytics_summary(
        "2026-07-01",
        "2026-07-31",
        company_id,
    )
    assert analytics["reports"] == 1
    assert analytics["records"] == 2
    assert analytics["work_total"] == 9
    database.set_report_lock(report_id, True, "admin")
    try:
        database.save_snapshot(
            snapshot,
            report_id,
            company_id=company_id,
            template_id=template_id,
            status="approved",
            actor="admin",
        )
    except PermissionError:
        pass
    else:
        raise AssertionError("Locked reports must reject updates.")
    database.set_report_lock(report_id, False, "admin")
    editor_id = database.save_user(
        "editor_test",
        "محرر الاختبار",
        "editor",
        True,
        "admin",
        password="Secure123!",
    )
    assert editor_id
    assert database.authenticate("editor_test", "Secure123!") is not None
    config_path = test_dir / "config.json"
    atomic_write_json(config_path, config)
    backup_manager = BackupManager(test_dir, database, config_path)
    assert backup_manager.automatic_backup(force=True)
    archive = backup_manager.create_archive(test_dir / "full_backup.zip")
    assert archive.is_file()
    export_snapshot_excel_to_path(snapshot, test_dir / "daily_report.xlsx")
    export_analytics_excel_to_path(
        analytics,
        database.get_company(company_id),
        test_dir / "management_report.xlsx",
    )
    export_management_pdf_to_path(
        analytics,
        database.get_company(company_id),
        test_dir / "management_report.pdf",
    )
    assert (test_dir / "daily_report.xlsx").stat().st_size > 4000
    assert (test_dir / "management_report.xlsx").stat().st_size > 4000
    assert (test_dir / "management_report.pdf").stat().st_size > 1000
    assert len(database.list_audit()) >= 4
    print(f"V9 SMOKE TEST OK: {test_dir}")


def run_v10_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    test_dir = output_dir / f"v10_test_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=False)
    config = normalize_config(DEFAULT_CONFIG)
    category_id = category_ids(config)[0]
    category = config["categories"][category_id]
    numeric_fields = [
        field for field in category["fields"] if field.get("kind") == "number"
    ]
    assert numeric_fields
    field = numeric_fields[0]
    person = (category.get("people") or ["مستخدم الاختبار"])[0]

    csv_path = test_dir / "import.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["الفئة", "الاسم/الجهة", field["label"]])
        writer.writerow([category["label"], person, 7])
        writer.writerow(["فئة غير موجودة", person, 2])
    csv_result = import_entries_from_path(csv_path, config)
    assert csv_result["imported"] == 1 and csv_result["skipped"] == 1
    assert csv_result["entries"][0]["values"][field["id"]] == 7

    excel_path = export_import_template_to_path(config, test_dir / "import.xlsx")
    workbook = load_workbook(excel_path)
    sheet = workbook["الاستيراد"]
    header_values = [cell.value for cell in sheet[1]]
    field_column = header_values.index(field["label"]) + 1
    sheet.cell(row=2, column=1, value=category["label"])
    sheet.cell(row=2, column=2, value=person)
    sheet.cell(row=2, column=field_column, value=9)
    workbook.save(excel_path)
    workbook.close()
    excel_result = import_entries_from_path(excel_path, config)
    assert excel_result["imported"] >= 1

    json_path = test_dir / "import.json"
    atomic_write_json(
        json_path,
        [
            {
                "الفئة": category["label"],
                "الاسم/الجهة": person,
                field["label"]: 11,
            }
        ],
    )
    json_result = import_entries_from_path(json_path, config)
    assert json_result["imported"] == 1

    if PPTX_AVAILABLE:
        pptx_path = test_dir / "import.pptx"
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        table = slide.shapes.add_table(2, 3, 0, 0, 8_000_000, 1_500_000).table
        for index, value in enumerate(("الفئة", "الاسم/الجهة", field["label"])):
            table.cell(0, index).text = value
        for index, value in enumerate((category["label"], person, "13")):
            table.cell(1, index).text = value
        presentation.save(pptx_path)
        pptx_result = import_entries_from_path(pptx_path, config)
        assert pptx_result["imported"] == 1

    snapshot = build_report_snapshot(
        config,
        csv_result["entries"],
        date(2026, 8, 5),
        "V10-IMPORT",
        "1",
    )
    draft_store = DraftStore(test_dir / "drafts" / "draft_test.json")
    draft_store.save(snapshot, report_id="report-under-edit")
    loaded_draft = draft_store.load()
    assert loaded_draft and loaded_draft["report_id"] == "report-under-edit"
    draft_store.clear()
    assert draft_store.load() is None

    database = ReportDatabase(test_dir / "reports.db")
    config_path = test_dir / "config.json"
    atomic_write_json(config_path, config)
    manager = BackupManager(test_dir, database, config_path)
    health = database.health_check()
    assert health["ok"]
    maintenance = manager.maintain("admin")
    assert maintenance["ok"] and Path(maintenance["backup"]).is_file()
    print(f"V10 SMOKE TEST OK: {test_dir}")


def run_v11_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    test_dir = output_dir / f"v11_test_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=False)
    assert "*" in ROLE_PERMISSIONS["admin"]
    config = normalize_config(
        {
            **DEFAULT_CONFIG,
            "app_state": {
                "company_id": "company",
                "template_id": "template",
                "dashboard_filters": {"start_date": "2026-08-01"},
                "history_filters": {"search": "اختبار"},
            },
        }
    )
    assert config["app_state"]["dashboard_filters"]["start_date"] == "2026-08-01"
    assert config["app_state"]["history_filters"]["search"] == "اختبار"

    database = ReportDatabase(test_dir / "source.db")
    snapshot = {
        "date": "2026-08-01",
        "day": "السبت",
        "company_name": "شركة الاختبار",
        "user_name": "مدير النظام",
        "code": "V11",
        "page": "1",
        "rows": [],
        "status": "draft",
    }
    report_id = database.save_snapshot(snapshot, actor="مدير النظام")
    verification = database.verify_report(report_id)
    assert verification["valid"] and verification["revision"] == 1
    snapshot["rows"] = [
        {
            "category": "quality",
            "category_label": "الجودة",
            "name": "الموظف التجريبي",
            "fields": [{"id": "status", "label": "الحالة", "value": "مكتمل"}],
            "total": 7,
        }
    ]
    database.save_snapshot(snapshot, report_id, actor="مدير النظام")
    assert len(database.list_report_revisions(report_id)) == 2
    assert database.list_reports(status="draft")[0]["id"] == report_id

    database.configure_approval_workflow(
        report_id, ["مراجعة المشرف", "اعتماد مدير النظام"], "مدير النظام"
    )
    try:
        database.decide_report_approval(
            report_id, 2, "approved", "مدير النظام"
        )
    except PermissionError:
        pass
    else:
        raise AssertionError("قُبلت مرحلة موافقة قبل المرحلة السابقة.")
    database.decide_report_approval(report_id, 1, "approved", "مدير النظام")
    database.decide_report_approval(report_id, 2, "approved", "مدير النظام")
    assert database.get_report(report_id)["is_locked"]
    database.restore_report_revision(report_id, 1, "مدير النظام")
    assert not database.get_report(report_id)["is_locked"]

    sync_path = database.export_sync_package(
        test_dir / "reports.dwrsync", [report_id], "مدير النظام"
    )
    destination = ReportDatabase(test_dir / "destination.db")
    sync_result = destination.import_sync_package(sync_path, "مدير النظام")
    assert sync_result["created"] == 1
    assert destination.verify_report(report_id)["valid"]
    duplicate = destination.import_sync_package(sync_path, "مدير النظام")
    assert duplicate["duplicate_package"]

    forecast = neural_forecast([3, 5, 4, 7, 6, 8, 9, 10, 11])
    assert forecast["available"] and forecast["prediction"] is not None
    analytics = destination.analytics_summary("2026-08-01", "2026-08-31")
    assert "forecast" in analytics and "insights" in analytics and "anomalies" in analytics
    print(f"V11 SMOKE TEST OK: {test_dir}")


def run_v12_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    test_dir = output_dir / f"v12_test_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=False)
    assert _config_version_number(APP_VERSION.split(".")[0]) >= 14 and CONFIG_VERSION >= 8
    assert "*" in ROLE_PERMISSIONS["admin"]

    config = normalize_config(DEFAULT_CONFIG)
    database = ReportDatabase(test_dir / "reports.db")
    users = database.list_users()
    admin_row = next(item for item in users if item["role"] == "admin")
    admin = {
        "id": admin_row["id"],
        "username": admin_row["username"],
        "display_name": admin_row["display_name"],
        "role": "admin",
    }
    viewer_id = database.save_user(
        "viewer_v12", "مستخدم العرض", "viewer", True, "مدير النظام"
    )
    viewer = {
        "id": viewer_id,
        "username": "viewer_v12",
        "display_name": "مستخدم العرض",
        "role": "viewer",
    }

    shared_view_id = database.save_view(
        "history",
        "الفترة النشطة",
        {"start_date": "2026-08-01", "search": "الجودة"},
        admin,
        is_shared=True,
        is_default=True,
    )
    assert database.list_views("history", viewer)[0]["id"] == shared_view_id
    private_view_id = database.save_view(
        "history", "خاص", {"status": "draft"}, viewer
    )
    assert any(item["id"] == private_view_id for item in database.list_views("history", viewer))
    try:
        database.save_view(
            "dashboard", "مشاركة ممنوعة", {}, viewer, is_shared=True
        )
    except PermissionError:
        pass
    else:
        raise AssertionError("تمكن مستخدم غير مدير من مشاركة عرض.")

    first_category_id = category_ids(config)[0]
    first_category = config["categories"][first_category_id]
    first_field = first_category["fields"][0]
    csv_path = test_dir / "custom_headers.csv"
    csv_path.write_text(
        "Team,Worker,Done\n"
        f"{first_category['label']},{first_category['people'][0]},9\n",
        encoding="utf-8-sig",
    )
    mapping = {
        "Team": "الفئة",
        "Worker": "الاسم/الجهة",
        "Done": str(first_field["label"]),
    }
    profile_id = database.save_import_profile(
        "ربط المورد الخارجي",
        ".csv",
        mapping,
        admin,
        is_shared=True,
    )
    profile = next(
        item
        for item in database.list_import_profiles(".csv", viewer)
        if item["id"] == profile_id
    )
    preview = analyze_import_path(csv_path, config, profile)
    assert preview["imported"] == 1 and preview["skipped"] == 0
    imported = import_entries_from_path(csv_path, config, profile)
    assert imported["entries"][0]["name"] == first_category["people"][0]
    assert imported["entries"][0]["values"][first_field["id"]] == 9
    try:
        database.save_import_profile(
            "غير مصرح", ".csv", {}, viewer, is_shared=True
        )
    except PermissionError:
        pass
    else:
        raise AssertionError("تمكن مستخدم غير مدير من مشاركة ملف استيراد.")

    snapshot = build_report_snapshot(
        config,
        imported["entries"],
        date(2026, 8, 5),
        "V12",
        "1",
        "شركة الاختبار",
        "مدير النظام",
    )
    quality = report_quality_score(snapshot)
    assert quality["score"] >= 85 and quality["model"] == "report-quality-v2"
    assert report_quality_score({"rows": []})["score"] < quality["score"]

    draft_store = DraftStore(test_dir / "drafts" / "current.json")
    draft_store.save(snapshot)
    newer_snapshot = copy.deepcopy(snapshot)
    newer_snapshot["code"] = "V12-RECOVERED"
    draft_store.save(newer_snapshot)
    draft_store.path.write_text("{corrupted", encoding="utf-8")
    recovered = draft_store.load()
    assert recovered and recovered["snapshot"]["code"] == "V12-RECOVERED"
    assert len(list(draft_store.history_dir.glob("draft_*.json"))) == 2

    forecast = neural_forecast([5, 6, 7, 7, 8, 9, 11, 12, 13, 15])
    assert forecast["available"] and forecast["model"] == "daily-local-neural-v2"
    assert forecast["lower_bound"] <= forecast["prediction"] <= forecast["upper_bound"]
    scenario = scenario_forecast([5, 6, 7, 7, 8, 9, 11, 12, 13, 15], 10)
    assert scenario["available"] and scenario["adjustment_percent"] == 10
    assert scenario["prediction"] > forecast["prediction"]

    report_id = database.save_snapshot(snapshot, actor="مدير النظام")
    analytics = database.analytics_summary("2026-08-01", "2026-08-31")
    assert report_id and analytics["quality"]["reports"] == 1
    assert analytics["quality"]["average"] >= 85
    assert "scenario" in analytics and "forecast" in analytics
    database.delete_view(private_view_id, admin)
    database.delete_import_profile(profile_id, admin)
    print(f"V12 SMOKE TEST OK: {test_dir}")


def run_v13_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    test_dir = output_dir / f"v13_test_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=False)
    assert APP_VERSION in ("15.2", "15.3", "15.3.1") and CONFIG_VERSION == 11

    legacy = copy.deepcopy(DEFAULT_CONFIG)
    legacy["version"] = 6
    drawer_by_id = {
        field["id"]: field for field in legacy["categories"]["drawers"]["fields"]
    }
    legacy["categories"]["drawers"]["fields"] = [
        drawer_by_id["black_white"],
        drawer_by_id["colored"],
        drawer_by_id["edits"],
        {"id": "designer_note", "label": "بيان إضافي", "kind": "text"},
    ]
    legacy["categories"]["technicians"]["fields"] = [
        field
        for field in reversed(legacy["categories"]["technicians"]["fields"])
        if field["id"] != "notes"
    ]
    migrated = normalize_config(legacy)
    assert [
        field["id"] for field in migrated["categories"]["drawers"]["fields"]
    ] == ["colored", "edits", "black_white", "designer_note"]
    assert [
        field["id"] for field in migrated["categories"]["technicians"]["fields"]
    ] == ["new_measure", "confirmed_measure", "outlets", "cancelled", "new_requests", "notes"]

    customized = copy.deepcopy(DEFAULT_CONFIG)
    customized_fields = customized["categories"]["drawers"]["fields"]
    assert move_field_by_drop(
        customized_fields,
        "black_white",
        "colored",
        place_after=False,
    )
    customized = normalize_config(customized)
    assert [
        field["id"] for field in customized["categories"]["drawers"]["fields"]
    ] == ["black_white", "colored", "edits"]

    snapshot = build_report_snapshot(
        migrated,
        [
            new_entry(
                "drawers",
                "موظف اختبار",
                {"colored": 3, "edits": 2, "black_white": 1},
            ),
            new_entry(
                "technicians",
                "موظف اختبار",
                {
                    "new_measure": 4,
                    "confirmed_measure": 3,
                    "outlets": 2,
                    "cancelled": 1,
                    "new_requests": 5,
                    "notes": "اختبار ترتيب الأعمدة",
                },
            ),
        ],
        date(2026, 8, 15),
        "V13-COLUMNS",
        "1",
    )
    drawer_fields, drawer_rows, _drawer_total = _logical_table_rows(snapshot, "drawers")
    technician_fields, technician_rows, _technician_total = _logical_table_rows(
        snapshot,
        "technicians",
    )
    assert [field["id"] for field in drawer_fields[:3]] == [
        "colored",
        "edits",
        "black_white",
    ]
    assert [field["id"] for field in technician_fields] == [
        "new_measure",
        "confirmed_measure",
        "outlets",
        "cancelled",
        "new_requests",
        "notes",
    ]
    assert drawer_rows[0][2:5] == [3, 2, 1]
    assert technician_rows[0][2:8] == [4, 3, 2, 1, 5, "اختبار ترتيب الأعمدة"]
    export_csv_to_path(snapshot, test_dir / "v13_column_order.csv")
    if OPENPYXL_AVAILABLE:
        export_snapshot_excel_to_path(snapshot, test_dir / "v13_column_order.xlsx")
    print(f"V14 COMPATIBILITY SMOKE TEST OK: {test_dir}")


def run_v15_feature_smoke_test() -> None:
    test_dir = Path(os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))) / f"v15_feature_{uuid.uuid4().hex}"
    test_dir.mkdir(parents=True, exist_ok=True)
    config = normalize_config(copy.deepcopy(DEFAULT_CONFIG))
    fields = config["categories"]["drawers"]["fields"]
    fields.extend([
        {"id":"due","label":"تاريخ الاستحقاق","kind":"date","options":[],"required":True,"report_visible":True,"report_label":""},
        {"id":"time","label":"الوقت","kind":"time","options":[],"report_visible":True,"report_label":""},
        {"id":"ok","label":"تم","kind":"boolean","options":[],"report_visible":True,"report_label":""},
        {"id":"pct","label":"النسبة","kind":"percentage","options":[],"report_visible":True,"report_label":"نسبة الإنجاز"},
        {"id":"amount","label":"المبلغ","kind":"currency","options":[],"report_visible":False,"report_label":""},
        {"id":"phone","label":"الجوال","kind":"phone","options":[],"report_visible":True,"report_label":""},
    ])
    config["report_layout"]["orientation"] = "portrait"
    config["identity"]["report_footer"] = "تقرير مخصص v15"
    config = normalize_config(config)
    assert all(kind in SUPPORTED_FIELD_KINDS for kind in [f["kind"] for f in config["categories"]["drawers"]["fields"]])
    entry = new_entry("drawers", "اختبار", {"colored":2,"edits":1,"black_white":3,"due":"2026-09-02","time":"11:30","ok":"نعم","pct":75.5,"amount":1500.25,"phone":"0500000000"})
    snapshot = build_report_snapshot(config,[entry],date(2026,9,2),"V15","1")
    assert snapshot["report_orientation"] == "portrait"
    assert snapshot["report_footer"] == "تقرير مخصص v15"
    labels = [f["label"] for f in snapshot["rows"][0]["fields"]]
    assert "نسبة الإنجاز" in labels and "المبلغ" not in labels
    db = ReportDatabase(test_dir / "reports.db")
    company_id, template_id = db.ensure_defaults(config)
    rid = db.save_snapshot(snapshot, None, company_id=company_id, template_id=template_id, actor="qa")
    assert db.count() == 1
    assert db.archive_reports([rid], actor="qa") == 1
    assert db.count() == 0
    assert len(db.list_reports(archived_only=True)) == 1
    assert db.restore_archived_reports([rid], actor="qa") == 1
    assert db.count() == 1
    assert db.archive_reports([rid], actor="qa") == 1
    assert db.purge_archived_reports([rid], actor="qa") == 1
    assert db.count(include_archived=True) == 0
    if REPORTLAB_AVAILABLE:
        export_pdf_to_path(snapshot, test_dir / "portrait.pdf")
        assert (test_dir / "portrait.pdf").is_file()
    print(f"V15 FEATURE SMOKE TEST OK: {test_dir}")


def run_template_visual_smoke_test() -> None:
    output_dir = Path(
        os.environ.get("DAILY_WORK_REPORTS_TEST_DIR", str(Path.cwd()))
    ).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    icon_path = Path(__file__).resolve().parent / "assets" / "daily_work_reports_icon.png"
    entries = [
        new_entry(
            "drawers",
            "موظف اختبار",
            {"black_white": 8, "colored": 4, "edits": 2},
        ),
        new_entry(
            "drawers",
            "موظف اختبار",
            {"black_white": 5, "colored": 6, "edits": 1},
        ),
        new_entry(
            "technicians",
            "موظف اختبار",
            {
                "new_measure": 7,
                "confirmed_measure": 3,
                "outlets": 2,
                "notes": "تمت مراجعة الموقع واعتماد مواقع التمديدات مع العميل",
            },
        ),
        new_entry(
            "showroom_employees",
            "موظف اختبار",
            {"drawings_count": 4, "measurements_count": 3},
        ),
    ]
    company_profile = {
        "name": "شركة الأفق للأعمال والتصميم",
        "address": "الرياض - المملكة العربية السعودية",
        "phone": "011 000 0000",
        "tax_number": "300000000000003",
        "logo_path": str(icon_path if icon_path.is_file() else ""),
    }
    for theme_id in REPORT_THEMES:
        config = normalize_config(DEFAULT_CONFIG)
        config["report_theme"] = theme_id
        snapshot = build_report_snapshot(
            config,
            entries,
            date(2026, 7, 30),
            "DWR-0730",
            "18",
            company_profile["name"],
            "مدير التقارير",
            company_profile=company_profile,
            company_id="visual-company",
            template_id=f"builtin_{theme_id}",
            template_name=REPORT_THEMES[theme_id]["label"],
            status="approved",
        )
        export_pdf_to_path(snapshot, output_dir / f"template_{theme_id}.pdf")
        export_images_to_path(snapshot, output_dir / f"template_{theme_id}.png")
        export_snapshot_excel_to_path(
            snapshot,
            output_dir / f"template_{theme_id}.xlsx",
        )
    analytics = {
        "start_date": "2026-07-24",
        "end_date": "2026-07-30",
        "company_id": "visual-company",
        "reports": 7,
        "records": 36,
        "work_total": 284,
        "average_per_report": 40.57,
        "freshness": "2026-07-30T16:45:00",
        "daily": [
            {"label": f"2026-07-{day:02d}", "reports": 1, "records": records, "total": total}
            for day, records, total in (
                (24, 4, 31),
                (25, 5, 38),
                (26, 6, 46),
                (27, 4, 35),
                (28, 7, 55),
                (29, 5, 42),
                (30, 5, 37),
            )
        ],
        "categories": [
            {"label": "أعمال الرسم والتصميم", "records": 14, "total": 118},
            {"label": "أعمال المقاسات", "records": 12, "total": 94},
            {"label": "أعمال الموظفين", "records": 10, "total": 72},
        ],
        "people": [
            {"label": "موظف اختبار", "records": 7, "total": 64},
            {"label": "موظف اختبار", "records": 6, "total": 57},
            {"label": "موظف اختبار", "records": 6, "total": 51},
            {"label": "موظف اختبار", "records": 5, "total": 43},
            {"label": "موظف اختبار", "records": 4, "total": 37},
        ],
        "statuses": [
            {"label": "approved", "reports": 4},
            {"label": "completed", "reports": 2},
            {"label": "draft", "reports": 1},
        ],
    }
    export_management_pdf_to_path(
        analytics,
        company_profile,
        output_dir / "management_visual.pdf",
    )
    export_analytics_excel_to_path(
        analytics,
        company_profile,
        output_dir / "management_visual.xlsx",
    )
    print("TEMPLATE VISUAL SMOKE TEST OK")


def main() -> None:
    global DEFAULT_CONFIG
    if any(arg.endswith("-test") for arg in sys.argv[1:]):
        DEFAULT_CONFIG = make_qa_config()
    if "--self-test" in sys.argv:
        run_self_test()
        return
    if "--database-smoke-test" in sys.argv:
        run_database_smoke_test()
        return
    if "--export-smoke-test" in sys.argv:
        run_export_smoke_test()
        return
    if "--period-export-smoke-test" in sys.argv:
        run_period_export_smoke_test()
        return
    if "--pdf-design-smoke-test" in sys.argv:
        run_pdf_design_smoke_test()
        return
    if "--v9-smoke-test" in sys.argv:
        run_v9_smoke_test()
        return
    if "--v10-smoke-test" in sys.argv:
        run_v10_smoke_test()
        return
    if "--v11-smoke-test" in sys.argv:
        run_v11_smoke_test()
        return
    if "--v12-smoke-test" in sys.argv:
        run_v12_smoke_test()
        return
    if "--v13-smoke-test" in sys.argv or "--v14-smoke-test" in sys.argv or "--v15-smoke-test" in sys.argv:
        run_v13_smoke_test()
        return
    if "--v15-feature-smoke-test" in sys.argv:
        run_v15_feature_smoke_test()
        return
    if "--template-visual-smoke-test" in sys.argv:
        run_template_visual_smoke_test()
        return
    root = tk.Tk()
    if "--ui-smoke-test" not in sys.argv:
        from public_setup_v155 import ensure_initial_configuration
        if not ensure_initial_configuration(
            root, app_data_dir() / "config.json",
            DEFAULT_CONFIG, normalize_config, read_json, atomic_write_json,
        ):
            root.destroy()
            return
    app = KitchensERP(root)
    if "--ui-smoke-test" in sys.argv:
        root.state("normal")
        root.geometry("1100x720")
        root.update()
        assert app.category_combo.cget("values")
        def all_buttons(parent: tk.Misc) -> list[tk.Button]:
            buttons: list[tk.Button] = []
            for widget in parent.winfo_children():
                if isinstance(widget, tk.Button):
                    buttons.append(widget)
                buttons.extend(all_buttons(widget))
            return buttons

        button_labels = {widget.cget("text") for widget in all_buttons(app.actions_frame)}
        def all_menus(parent: tk.Misc) -> list[tk.Menubutton]:
            found = []
            for widget in parent.winfo_children():
                if isinstance(widget, tk.Menubutton):
                    found.append(widget)
                found.extend(all_menus(widget))
            return found

        menu_labels = {widget.cget("text") for widget in all_menus(app.actions_frame)}
        assert "حفظ التقرير" in button_labels
        assert "الرئيسية" in button_labels
        assert "إدارة الموظفين والحقول" in button_labels
        assert "تقرير جديد" in button_labels
        assert {
            "ملف ▾", "إدخال ▾", "تصدير ▾", "التقارير ▾", "تخصيص ▾", "إدارة ▾"
        }.issubset(menu_labels)
        assert app.root.cget("menu")
        assert app.quality_var.get().startswith("جودة التقرير:")
        assert app.actions_frame.winfo_exists()
        homepage = ModernHome(app)
        root.update()
        assert "15.5" in homepage.title()
        homepage.destroy()
        configuration = ConfigurationManager(app)
        root.update()
        assert [
            getattr(widget, "field_id", "")
            for widget in configuration.column_order_strip.winfo_children()
        ] == ["colored", "edits", "black_white"]
        configuration.destroy()
        structure = StructureManager(app)
        root.update()
        assert structure.fields_tree.get_children()
        structure.destroy()
        studio = StructureStudio(app)
        root.update()
        assert studio.categories_list.size() >= 3
        assert studio.fields_tree.get_children()
        studio.destroy()
        designer = ReportDesigner(app)
        root.update()
        assert designer.tree.get_children()
        designer.destroy()
        appearance = AppearanceManager(app)
        root.update()
        assert appearance.mode_var.get() in APPEARANCE_MODES.values()
        appearance.destroy()
        assert all(
            row.winfo_reqwidth() < 1100
            for row in app.actions_frame.winfo_children()
            if isinstance(row, tk.Frame)
        )
        app.show_history()
        app.show_dashboard()
        app.show_backup_manager()
        app.show_company_manager()
        app.show_template_manager()
        app.show_administration()
        app.show_import_profiles_manager()
        root.update()
        for widget in root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                widget.destroy()
        root.destroy()
        print("UI SMOKE TEST OK")
        return
    root.after(420, app.open_modern_home)
    root.mainloop()


if __name__ == "__main__":
    main()
