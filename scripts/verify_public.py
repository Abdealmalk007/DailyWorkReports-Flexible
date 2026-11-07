"""GitHub preflight checks for clean user-defined startup."""
from pathlib import Path
import ast,re
ROOT=Path(__file__).resolve().parents[1]
src=(ROOT/"windows/DailyWorkReports_15.2.py").read_text(encoding="utf-8")
java=(ROOT/"android/app/src/main/java/com/protech/dailyreports/MainActivity.java").read_text(encoding="utf-8")
win=(ROOT/"windows/public_setup_v155.py").read_text(encoding="utf-8")
ui=(ROOT/"windows/modern_ui_v153.py").read_text(encoding="utf-8")
head=src[src.index("DEFAULT_CONFIG: dict[str, Any] = {"):src.index("def make_qa_config()")]
for expected in ('"category_order": []','"categories": {}','"report_title": ""','"company_name": ""'):
    assert expected in head,expected
assert '"people": [' not in head and '"fields": [' not in head
assert 'APP_VERSION = "15.5.0"' in src and "DailyWorkReportsFlexible_v155" in src
for expected in ('put(cfg,"category_order",new JSONArray());','put(cfg,"report_title","");','put(cfg,"categories",cats);'):
    assert expected in java,expected
assert 'applicationId = "com.protech.dailyreports.flexible"' in (ROOT/"android/app/build.gradle.kts").read_text(encoding="utf-8")
assert 'android:allowBackup="false"' in (ROOT/"android/app/src/main/AndroidManifest.xml").read_text(encoding="utf-8")
for body in (win,ui,java):
    for entry in ("Abu Omar","protech.2025.2026@gmail.com","966503276927"):
        assert entry in body
for f in ROOT.rglob("*"):
    if not f.is_file():continue
    if set(f.relative_to(ROOT).parts).intersection({".git",".venv","__pycache__","build",".gradle","dist","Output"}):
        continue
    assert f.suffix.lower() not in {".db",".sqlite",".sqlite3",".apk",".exe",".aab",".zip",".jks",".keystore",".log",".csv",".pdf"}
    if f.suffix==".py":ast.parse(f.read_text(encoding="utf-8"),filename=str(f))
    if f.suffix.lower() in {".py",".java",".kts",".xml",".md",".txt",".properties",".yml",".spec",".bat"}:
        body=f.read_text(encoding="utf-8",errors="replace")
        assert not re.search(r"ghp_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9_-]{20,}",body),str(f)
        assert not re.search(r"(?i)C:\\Users\\[^\s\\]+",body),str(f)
print("DWR_PUBLIC_155_PREFLIGHT_PASS")
