"""Synthetic regression for per-employee metrics and calendar decisions."""
from datetime import date
from decimal import Decimal
from pathlib import Path
import sys
import importlib.util

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/"windows"))
from daily_productivity_v157 import employee_productivity, latest_daily_snapshots, period_bounds, format_total

config={"categories":{
    "tech":{"label":"Technical","people":["Alice","Bob"],"fields":[]},
    "draw":{"label":"Design","people":["Charlie"],"fields":[]},
}}
def make_snapshot(day:str,amount:int,name="Alice",dept="tech"):
    return {
        "date":day,"config":config,
        "rows":[{
            "category":dept,"category_label":config["categories"][dept]["label"],"name":name,
            "fields":[
                {"id":"measures","label":"Measurements","kind":"number","value":amount},
                {"id":"drawings","label":"Drawings","kind":"number","value":2},
                {"id":"cost","label":"SAR","kind":"currency","value":"350.50"},
                {"id":"notes","label":"Notes","kind":"text","value":"test"},
            ],
            "total":amount+2,
        }]
    }
class FakeDB:
    def __init__(self):
        self.records=[
            ("old","2026-11-17","2026-11-17T10:00",make_snapshot("2026-11-17",12)),
            ("new","2026-11-17","2026-11-17T12:00",make_snapshot("2026-11-17",8)),
            ("today","2026-11-18","2026-11-18T10:00",make_snapshot("2026-11-18",4,"Bob")),
            ("oct","2026-10-16","2026-10-16T10:00",make_snapshot("2026-10-16",3,"Charlie","draw")),
        ]
    def list_reports(self,start_date=None,end_date=None,**_):
        return [
            {"id":id,"report_date":day,"updated_at":updated,"company_id":"test","saved_at":updated}
            for id,day,updated,_ in self.records
            if (not start_date or day>=start_date) and (not end_date or day<=end_date)
        ]
    def get_report(self,report_id):
        for id,day,updated,snapshot in self.records:
            if report_id==id:
                return {"id":id,"snapshot":snapshot,"company_id":"test","status":"draft"}
        return None

data=FakeDB()
today=date(2026,11,18)
week=employee_productivity(data,config,"week",today)
month=employee_productivity(data,config,"month",today)
all_time=employee_productivity(data,config,"all",today)
assert week["reports"]==2,week["reports"]
assert month["reports"]==2
assert all_time["reports"]==3
persons={x["name"]:x for x in week["employees"]}
assert persons["Alice"]["total"]==Decimal(10)
assert persons["Bob"]["total"]==Decimal(6)
assert persons["Charlie"]["total"]==Decimal(0)
assert persons["Alice"]["days"]==1
assert sum(x["records"] for x in week["employees"])==2
assert len(week["fields"])==6  # 2 days x three numeric/currency categories
assert period_bounds("week",today)[0]=="2026-11-16"
assert period_bounds("month",today)[0]=="2026-11-01"
assert format_total(Decimal("123.40"))=="123.4"
assert all_time["days"]==3

spec=importlib.util.spec_from_file_location("daily_gui_v157",ROOT/"windows"/"DailyWorkReports_15.2.py")
module=importlib.util.module_from_spec(spec)
sys.modules[spec.name]=module
spec.loader.exec_module(module)

class Dummy:
    def __init__(self):
        self.selected_date=today
        self.entries=[]
        self.loaded_report_id=None
        self.database=data
        self.current_company_id="test"
        self.root=None
        self._date_navigation_active=False
        self.draft_store=self
        self.calendar=self
        self.database_status_var=self
        self.loaded=[]
        self.blank=[]
    def clear(self):pass
    def set_date(self,x):self.selected_date=x
    def set(self,x):self.status=x
    def load_snapshot_for_editing(self,report_id,snapshot,metadata):
        self.loaded.append((report_id,snapshot["date"]))
        self.selected_date=date.fromisoformat(snapshot["date"])
    def _prepare_blank_day(self,day):
        self.blank.append(day)
        self.selected_date=day
        self.entries=[]
        self.loaded_report_id=None

stub=Dummy()
module.KitchensERP.on_date_selected(stub,date(2026,11,17))
assert stub.loaded==[("new","2026-11-17")]
module.KitchensERP.on_date_selected(stub,date(2026,11,19))
assert stub.blank==[date(2026,11,19)]
print("V15.7.0 PRODUCTIVITY AND DATE REGRESSION TESTS OK")
