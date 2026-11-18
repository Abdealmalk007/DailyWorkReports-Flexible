"""Regression for KeyError('technicians') with historical private templates.

All real records are accessed read-only, copied to a disposable SQLite QA profile.
Nothing private is printed, committed, or uploaded.
Python 3.13, standard library plus application's dependencies.
"""
from __future__ import annotations
import copy
import importlib.util
import json
import os
from pathlib import Path
import shutil
import sqlite3
import sys
import tempfile
import tkinter as tk
from contextlib import closing

ROOT=Path(__file__).resolve().parents[1]
WIN=ROOT/"windows"
sys.path.insert(0,str(WIN))
spec=importlib.util.spec_from_file_location("dwr1571_qa",WIN/"DailyWorkReports_15.2.py")
assert spec and spec.loader
app=importlib.util.module_from_spec(spec)
sys.modules[spec.name]=app
spec.loader.exec_module(app)
assert app.APP_VERSION=="15.7.1"
assert app.DEFAULT_CONFIG["categories"]=={}  # actual public setup

# Pre-v11 snapshots must never assume a builtin category exists in DEFAULT_CONFIG.
fields=[{"id":"new_measure","label":"Legacy measurements","kind":"number"},
        {"id":"custom_notes","label":"Custom field","kind":"text"}]
before=copy.deepcopy(fields)
out=app.canonical_builtin_fields("technicians",fields,restore_missing=True)
assert out==before
assert app.canonical_builtin_fields("unknown_future_section",fields,restore_missing=True)==before
legacy_config={
 "version":6,
 "identity":{"report_title":"QA"},
 "category_order":["technicians"],
 "categories":{"technicians":{
    "label":"Measurements","people":["Synthetic Person"],"fields":copy.deepcopy(fields)
 }}
}
normalized=app.normalize_config(legacy_config)
assert normalized["categories"]["technicians"]["people"]==["Synthetic Person"]
assert [(f["id"],f["label"],f["kind"]) for f in normalized["categories"]["technicians"]["fields"]]==[(f["id"],f["label"],f["kind"]) for f in before]

source_dir=Path(os.environ.get("LOCALAPPDATA", ""))/"ProTech"/"DailyWorkReportsPrivate"
if (source_dir/"reports.db").is_file() and (source_dir/"config.json").is_file():
    test_private_copy = True
else:
    test_private_copy = False
if test_private_copy:
    with tempfile.TemporaryDirectory(prefix="dwr1571_private_qa_") as directory:
        destination=Path(directory)
        (destination/"config.json").write_bytes((source_dir/"config.json").read_bytes())
        original_db=source_dir/"reports.db"
        uri=original_db.as_uri()+"?mode=ro"
        with closing(sqlite3.connect(uri,uri=True,timeout=20)) as src:
            assert src.execute("PRAGMA integrity_check").fetchone()[0]=="ok"
            count=src.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
            with closing(sqlite3.connect(destination/"reports.db")) as target:
                src.backup(target)
        assert count==65
        os.environ["DAILY_WORK_REPORTS_DATA_DIR"]=str(destination)
        database=app.ReportDatabase(destination/"reports.db")
        templates=database.list_templates()  # reproduces screenshot's exact old failing call
        assert len(templates)>=3
        for template in templates:
            assert isinstance(template.get("config"),dict)
        cfg=json.loads((destination/"config.json").read_text(encoding="utf-8-sig"))
        clean=app.normalize_config(cfg)
        assert sum(len(c.get("people",[])) for c in clean["categories"].values())==28
        root=tk.Tk()
        root.withdraw()
        original_auth=app.KitchensERP._prompt_for_login
        app.KitchensERP._prompt_for_login=lambda self: {
            "id":"QA-TEMP", "username":"QA", "display_name":"QA", "role":"admin"
        }
        try:
            instance=app.KitchensERP(root)
            assert instance.database.count()==65
            assert sum(len(c.get("people",[])) for c in instance.config["categories"].values())==28
            assert len(instance.template_records)>=3
            assert instance.name_combo.winfo_exists()
            print("PASS real historical templates loaded into UI (isolated copy)")
        finally:
            app.KitchensERP._prompt_for_login=original_auth
            root.destroy()
print("PASS old template normalization, private staff count and report data preserved")

if not test_private_copy:
    print('SKIP private-profile integration on environments without local data')
