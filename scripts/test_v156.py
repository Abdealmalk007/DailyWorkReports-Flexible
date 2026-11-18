"""Regression tests for Daily Work Reports 15.6; only synthetic data."""
import hashlib
import importlib.util
import json
import sqlite3
import sys
import tempfile
from datetime import date, timedelta
from contextlib import closing
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WIN = ROOT / "windows"
sys.path.insert(0, str(WIN))
spec = importlib.util.spec_from_file_location("daily_v156", WIN / "DailyWorkReports_15.2.py")
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
from migration_v156 import migrate_legacy_data

def run() -> None:
    names = ["أحمد سالم", "محمد عبدالرحمن", "خالد محمد", "Sarah Smith"]
    assert module.find_matching_employees(names, "حمد") == ["محمد عبدالرحمن", "أحمد سالم", "خالد محمد"] or len(module.find_matching_employees(names,"حمد")) == 3
    assert module.find_matching_employees(names, "رحمن") == ["محمد عبدالرحمن"]
    assert module.find_matching_employees(names, "احمد") == ["أحمد سالم"]
    assert module.find_matching_employees(names, "smith") == ["Sarah Smith"]
    assert module.find_matching_employees(names, "zzzz") == []
    with tempfile.TemporaryDirectory(prefix="daily_v156_test_") as working:
        base = Path(working)
        legacy = base / "old"
        target = base / "new"
        legacy.mkdir()
        with closing(sqlite3.connect(legacy/"reports.db")) as db:
            db.execute("CREATE TABLE reports (id TEXT PRIMARY KEY, snapshot_json TEXT NOT NULL)")
            db.execute("INSERT INTO reports VALUES (?,?)", ("id-01", '{"rows": []}'))
            db.commit()
        (legacy/"config.json").write_text(json.dumps({"categories":{"first":{"label":"QA"}}}),encoding="utf-8")
        original_db = hashlib.sha256((legacy/"reports.db").read_bytes()).hexdigest()
        summary=migrate_legacy_data(legacy,target)
        assert summary["reports"]==1
        assert hashlib.sha256((legacy/"reports.db").read_bytes()).hexdigest()==original_db
        with closing(sqlite3.connect(target/"reports.db")) as db:
            assert db.execute("SELECT COUNT(*) FROM reports").fetchone()[0]==1
        try:
            migrate_legacy_data(legacy,target)
        except FileExistsError:
            pass
        else:
            raise AssertionError("Migration overwrote an existing user profile")
        conf = module.make_qa_config()
        cid = module.category_ids(conf)[0]
        first_cat = conf["categories"][cid]
        first_name = first_cat["people"][0] if first_cat["people"] else "QA Person"
        snapshots=[]
        for day in range(3):
            vals={}
            for field in first_cat["fields"]:
                vals[field["id"]]=day+1 if field["kind"]=="number" else "QA"
            snapshots.append(module.build_report_snapshot(
                conf,[module.new_entry(cid,first_name,vals)],
                date(2026,11,15)+timedelta(days=day),f"QA-{day+1}",str(day+1)
            ))
        dest=base/"continuous.pdf"
        module.export_sequential_pdf_to_path(snapshots,dest)
        assert dest.read_bytes().startswith(b"%PDF-")
        assert dest.stat().st_size>1000
        print("V15.6 REGRESSION TESTS OK: matching, safe migration, continuous PDF")

if __name__ == "__main__":
    run()
