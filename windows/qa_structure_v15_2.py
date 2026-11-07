"""Source + live Tcl/Tk checks for the v15.2 structure workspace.
Run under Windows or xvfb-run on Linux. Uses isolated user data.
"""
from __future__ import annotations

import copy
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import tkinter as tk

from structure_studio import (
    FieldDialog, StructureStudio, active_category_conflicts, clean_choices,
    create_structure_safety_copy, validate_structure,
)
from DailyWorkReports_15_2_compat import main_module

core = main_module()
core.DEFAULT_CONFIG = core.make_qa_config()


class StructurePureTests(unittest.TestCase):
    def test_normalization_and_validation(self):
        config = core.normalize_config(copy.deepcopy(core.DEFAULT_CONFIG))
        config["categories"]["drawers"]["fields"].append({
            "id": "new_select", "label": "حالة الطلب", "kind": "choice",
            "options": ["مكتمل", "جديد"], "required": True,
        })
        self.assertEqual(validate_structure(config), [])
        rebuilt = core.normalize_config(config)
        self.assertEqual([f for f in rebuilt["categories"]["drawers"]["fields"] if f["id"] == "new_select"][0]["options"], ["مكتمل", "جديد"])
        self.assertEqual(clean_choices("جديد، قيد التنفيذ; جديد\nمنجز"), ["جديد", "قيد التنفيذ", "منجز"])

    def test_prevent_draft_category_loss(self):
        before = core.normalize_config(copy.deepcopy(core.DEFAULT_CONFIG))
        after = copy.deepcopy(before)
        del after["categories"]["drawers"]
        after["category_order"].remove("drawers")
        rows = [core.new_entry("drawers", "موظف اختبار", {"colored": 9})]
        self.assertEqual(active_category_conflicts(before, after, rows), {"drawers": 1})
        self.assertEqual(active_category_conflicts(before, after, []), {})


class StructureLiveTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="dwr152_")
        self.original = os.environ.get("DAILY_WORK_REPORTS_DATA_DIR")
        os.environ["DAILY_WORK_REPORTS_DATA_DIR"] = self.temp.name
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = core.KitchensERP(self.root)
        self.root.update()

    def tearDown(self):
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                widget.destroy()
        if self.app._draft_save_job is not None:
            self.root.after_cancel(self.app._draft_save_job)
            self.app._draft_save_job = None
        self.root.destroy()
        if self.original is None:
            os.environ.pop("DAILY_WORK_REPORTS_DATA_DIR", None)
        else:
            os.environ["DAILY_WORK_REPORTS_DATA_DIR"] = self.original
        self.temp.cleanup()

    def test_field_dialog_choice_validation_and_properties(self):
        dialog = FieldDialog(self.root, None, ["ملون"])
        dialog.label_var.set("نوع المهمة")
        dialog.kind_var.set("قائمة منسدلة")
        dialog.options_text.insert("1.0", "جديد\nقيد التنفيذ\nجديد")
        dialog.required_var.set(True)
        dialog.visible_var.set(False)
        dialog.report_label_var.set("حالة المهمة")
        dialog.save()
        result = dialog.result
        self.assertEqual(result["options"], ["جديد", "قيد التنفيذ"])
        self.assertTrue(result["required"])
        self.assertFalse(result["report_visible"])
        self.assertEqual(result["report_label"], "حالة المهمة")

    def test_choice_options_edit_and_reorder(self):
        studio = StructureStudio(self.app)
        studio.category["fields"].append({
            "id": "task_state", "label": "الحالة", "kind": "choice",
            "options": ["جديد", "جاري"], "required": False,
        })
        studio.refresh_all()
        studio.option_field_var.set("الحالة")
        studio.refresh_options()
        with patch("structure_studio.simpledialog.askstring", return_value="مكتمل"):
            studio.add_option()
        self.assertEqual(studio.current_choice_field()["options"], ["جديد", "جاري", "مكتمل"])
        studio.options_list.selection_set(2)
        studio.move_option(-1)
        self.assertEqual(studio.current_choice_field()["options"], ["جديد", "مكتمل", "جاري"])
        studio.destroy()

    def test_workspace_edit_save_reload_and_safety_copy(self):
        studio = StructureStudio(self.app)
        self.root.update()
        self.assertGreater(len(studio.visible_categories), 0)
        with patch("structure_studio.simpledialog.askstring", return_value="المتابعة"):
            studio.add_category()
        self.assertEqual(studio.category["label"], "المتابعة")
        with patch("structure_studio.simpledialog.askstring", return_value="موظف اختبار"):
            studio.add_person()
        self.assertEqual(studio.category["people"], ["موظف اختبار"])
        studio.category["fields"].append({
            "id": "field_test_status", "label": "الحالة", "kind": "choice",
            "options": ["جديد", "منجز"], "required": True,
        })
        studio.refresh_all()
        self.assertEqual(validate_structure(studio.draft), [])
        with patch("structure_studio.messagebox.showinfo"):
            studio.commit()
        self.assertEqual(self.app.config["categories"][self.app.config["category_order"][-1]]["people"], ["موظف اختبار"])
        self.assertTrue(list(Path(self.temp.name, "structure_backups").glob("*.json")))
        self.assertEqual(core.normalize_config(core.read_json(self.app.config_path, {})), self.app.config)

    def test_deleting_field_does_not_erase_draft_values(self):
        row = core.new_entry("drawers", "موظف اختبار", {"colored": 7, "edits": 5})
        self.app.entries.append(row)
        draft = copy.deepcopy(self.app.config)
        draft["categories"]["drawers"]["fields"] = [
            f for f in draft["categories"]["drawers"]["fields"] if f["id"] != "colored"
        ]
        self.assertTrue(self.app.apply_configuration(draft))
        self.assertEqual(self.app.entries[0]["values"]["colored"], 7)
        self.assertEqual(self.app.entries[0]["values"]["edits"], 5)

    def test_draft_recovery_preserves_hidden_field_data(self):
        self.app.entries.append(core.new_entry("drawers", "موظف اختبار", {"colored": 13, "edits": 2}))
        draft = copy.deepcopy(self.app.config)
        draft["categories"]["drawers"]["fields"] = [
            f for f in draft["categories"]["drawers"]["fields"] if f["id"] != "colored"
        ]
        self.assertTrue(self.app.apply_configuration(draft))
        self.app._save_current_draft()
        payload = self.app.draft_store.load()
        self.assertIsNotNone(payload)
        self.assertEqual(payload["raw_entries"][0]["values"]["colored"], 13)
        self.assertNotIn("colored", [f["id"] for f in payload["snapshot"]["rows"][0]["fields"]])
        self.app.entries = []
        with patch.object(core.messagebox, "askyesnocancel", return_value=True):
            self.app._offer_draft_recovery()
        self.assertEqual(self.app.entries[0]["values"]["colored"], 13)

    def test_deleting_category_is_blocked(self):
        self.app.entries.append(core.new_entry("drawers", "موظف اختبار", {"colored": 4}))
        draft = copy.deepcopy(self.app.config)
        del draft["categories"]["drawers"]
        draft["category_order"].remove("drawers")
        with patch.object(core.messagebox, "showwarning"):
            self.assertFalse(self.app.apply_configuration(draft))
        self.assertIn("drawers", self.app.config["categories"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
