"""Safe opt-in local data migration from DailyWorkReports 15.3 to 15.6.

Never modifies the source database. No migration occurs without opt-in.
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import tempfile
from pathlib import Path
from contextlib import closing


def migrate_legacy_data(source_dir: Path, target_dir: Path) -> dict[str, int]:
    """Copy a valid legacy database and configuration into a *new, empty* profile.

    Raises if target already has an application profile; never overwrites.
    """
    source_dir = Path(source_dir).resolve()
    target_dir = Path(target_dir).resolve()
    source_db = source_dir / "reports.db"
    source_config = source_dir / "config.json"
    dest_db = target_dir / "reports.db"
    dest_config = target_dir / "config.json"
    if source_dir == target_dir or not source_db.is_file() or not source_config.is_file():
        raise ValueError("مسار بيانات النسخة السابقة غير صالح.")
    if dest_db.exists() or dest_config.exists():
        raise FileExistsError("النسخة الجديدة لديها بيانات بالفعل؛ لن يتم استبدالها.")
    config = json.loads(source_config.read_text(encoding="utf-8-sig"))
    if not isinstance(config, dict) or not isinstance(config.get("categories"), dict):
        raise ValueError("ملف إعدادات النسخة السابقة غير صالح.")

    source_uri = source_db.as_uri() + "?mode=ro"
    with closing(sqlite3.connect(source_uri, uri=True, timeout=20)) as source:
        if source.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ValueError("قاعدة بيانات النسخة السابقة لا تجتاز فحص السلامة.")
        old_count = source.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
        target_dir.mkdir(parents=True, exist_ok=True)
        staging = Path(tempfile.mkdtemp(prefix=".v156-migrate-", dir=target_dir))
        try:
            stage_db = staging / "reports.db"
            with closing(sqlite3.connect(stage_db)) as copy:
                source.backup(copy)
            with closing(sqlite3.connect(stage_db)) as check:
                if check.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise RuntimeError("فشل فحص سلامة النسخة المنسوخة.")
                copied_count = check.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
            if copied_count != old_count:
                raise RuntimeError("عدد التقارير المنسوخة يختلف عن الأصل.")
            stage_conf = staging / "config.json"
            shutil.copy2(source_config, stage_conf)
            json.loads(stage_conf.read_text(encoding="utf-8-sig"))
            if dest_db.exists() or dest_config.exists():
                raise FileExistsError("تم إنشاء بيانات في الوجهة أثناء الترحيل.")
            db_moved = False
            try:
                os.replace(stage_db, dest_db)
                db_moved = True
                os.replace(stage_conf, dest_config)
            except OSError:
                if db_moved and not dest_config.exists():
                    dest_db.unlink(missing_ok=True)
                raise
        finally:
            shutil.rmtree(staging, ignore_errors=True)
    # Preserve a separate copy of the previous assets; source remains untouched.
    source_assets = source_dir / "company_assets"
    target_assets = target_dir / "company_assets"
    if source_assets.is_dir() and not target_assets.exists():
        shutil.copytree(source_assets, target_assets, symlinks=False)
    return {"reports": int(old_count), "config_sections": len(config["categories"])}


def offer_legacy_upgrade(parent, target_dir: Path) -> bool:
    """Prompt only on a truly blank profile. No silent data imports."""
    import tkinter.messagebox as messagebox

    target_dir = Path(target_dir)
    if (target_dir / "reports.db").exists() or (target_dir / "config.json").exists():
        return False
    app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if not app_data:
        return False
    legacy_dir = Path(app_data) / "DailyWorkReports"
    if not (legacy_dir / "reports.db").is_file():
        return False
    if not messagebox.askyesno(
        "ترقية بيانات Daily Work Reports",
        "تم العثور على بيانات نسخة سابقة. هل تريد نسخ التقارير والموظفين "
        "والإعدادات إلى النسخة الجديدة؟\n\n"
        "ستبقى النسخة السابقة وملفاتها الأصلية دون تعديل.",
        parent=parent,
    ):
        return False
    try:
        summary = migrate_legacy_data(legacy_dir, target_dir)
    except (OSError, ValueError, RuntimeError, sqlite3.Error) as exc:
        messagebox.showerror("تعذر نقل البيانات بأمان", str(exc), parent=parent)
        return False
    messagebox.showinfo(
        "اكتملت الترقية",
        f"تم نسخ {summary['reports']} تقرير مع إعدادات الأقسام.\n"
        "بيانات النسخة السابقة محفوظة كما هي.",
        parent=parent,
    )
    return True
