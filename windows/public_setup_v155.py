"""Windows first-launch setup: user defines headings and fields; no seeded staff."""
from __future__ import annotations
import copy
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from uuid import uuid4

DEVELOPER = "Abu Omar"
EMAIL = "protech.2025.2026@gmail.com"
PHONE = "+966503276927"

def ensure_initial_configuration(root, config_path: Path, base_config, normalize, read_json, save_json) -> bool:
    """Return True when a nonempty user-configured section is available."""
    if config_path.is_file():
        try:
            if normalize(read_json(config_path, {})).get("categories"):
                return True
        except (OSError, ValueError, TypeError):
            pass

    root.withdraw()
    dialog = tk.Toplevel(root)
    dialog.title("إعداد التقارير لأول مرة | 15.5.0")
    dialog.geometry("700x610")
    dialog.minsize(620, 560)
    dialog.configure(bg="#F0F7F6")
    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    variables = {key: tk.StringVar() for key in ("report_title", "category", "person", "field")}
    field_kind = tk.StringVar(value="number")

    tk.Label(dialog, text="إعداد تقريرك الأول", bg="#1B464B", fg="#FFFFFF",
             font=("Arial", 21, "bold"), pady=20).pack(fill="x")
    tk.Label(dialog, text="لا توجد أسماء أو أقسام أو عناوين أو حقول جاهزة. اكتب ما يناسب عملك.",
             bg="#F0F7F6", fg="#28565B", font=("Arial", 11), pady=12).pack(fill="x")

    frame = tk.Frame(dialog, bg="#F0F7F6")
    frame.pack(fill="x", padx=30, pady=4)
    for label, key in (
        ("عنوان التقرير", "report_title"),
        ("اسم القسم الأول", "category"),
        ("مسمى الموظف أو الجهة", "person"),
        ("عنوان الحقل الأول", "field"),
    ):
        tk.Label(frame, text=label, bg="#F0F7F6", fg="#18494E",
                 font=("Arial", 11, "bold"), anchor="e").pack(fill="x", pady=(8, 2))
        ttk.Entry(frame, textvariable=variables[key], justify="right", font=("Arial", 12)).pack(fill="x")

    tk.Label(frame, text="نوع الحقل (number / text / choice)", bg="#F0F7F6",
             fg="#18494E", font=("Arial", 10), anchor="e").pack(fill="x", pady=(8, 2))
    ttk.Combobox(frame, textvariable=field_kind, state="readonly",
                 values=("number", "text", "choice"), justify="right").pack(fill="x")

    completed = {"ok": False}

    def save():
        vals = {key: variable.get().strip() for key, variable in variables.items()}
        if not all(vals.values()):
            messagebox.showwarning("بيانات مطلوبة", "أكمل العنوان والقسم والمسمى واسم الحقل.", parent=dialog)
            return
        conf = copy.deepcopy(base_config)
        section = "section_" + uuid4().hex[:12]
        conf["identity"]["report_title"] = vals["report_title"]
        conf["category_order"] = [section]
        conf["categories"] = {
            section: {
                "label": vals["category"],
                "person_label": vals["person"],
                "total_label": "الإجمالي",
                "people": [],
                "fields": [{
                    "id": "field_" + uuid4().hex[:12], "label": vals["field"],
                    "kind": field_kind.get(), "options": [], "required": False,
                }],
            },
        }
        try:
            save_json(config_path, normalize(conf))
        except (OSError, ValueError) as error:
            messagebox.showerror("تعذر حفظ الإعدادات", str(error), parent=dialog)
            return
        completed["ok"] = True
        dialog.destroy()

    controls = tk.Frame(dialog, bg="#F0F7F6")
    controls.pack(fill="x", padx=30, pady=(15, 8))
    tk.Button(controls, text="إنشاء النموذج وفتح البرنامج", command=save,
              bg="#228567", fg="#FFFFFF", font=("Arial", 11, "bold"), padx=15,
              pady=9).pack(side="right")
    tk.Button(controls, text="إلغاء", command=dialog.destroy, font=("Arial", 11),
              padx=15, pady=9).pack(side="left")
    tk.Label(dialog, text=f"Developer: {DEVELOPER} | {EMAIL} | {PHONE}",
             bg="#F0F7F6", fg="#56787A", font=("Arial", 9)).pack(side="bottom", pady=10)
    dialog.transient(root)
    dialog.grab_set()
    dialog.lift()
    dialog.focus_force()
    root.wait_window(dialog)
    if completed["ok"]:
        root.deiconify()
    return completed["ok"]
