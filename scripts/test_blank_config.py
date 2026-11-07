from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/"windows"))
from DailyWorkReports_15_2_compat import main_module
app=main_module()
initial=app.DEFAULT_CONFIG
assert initial["categories"]=={}
assert initial["category_order"]==[]
assert initial["identity"]["report_title"]==""
assert initial["identity"]["company_name"]==""
normalized=app.normalize_config(initial)
assert normalized["categories"]=={}
assert normalized["category_order"]==[]
assert normalized["identity"]["report_title"]==""
print("EMPTY_CONFIG_TEST_PASS")
