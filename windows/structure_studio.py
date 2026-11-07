"""ProTech v15.2: transactional RTL workspace for sections, employees and fields.

This module intentionally does not import the main application, which keeps the
Windows/PyInstaller entry points and the existing report/database format stable.
"""
from __future__ import annotations

import copy
import json
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk


NAVY = "#18334B"
PRIMARY = "#27627A"
BLUE = "#3276A7"
GREEN = "#198563"
RED = "#BD4954"
GOLD = "#BA8952"
SURFACE = "#FFFFFF"
PALE = "#EDF3F7"
MUTED = "#607385"
INK = "#193448"
KIND_LABELS = {
    "number": "رقمي", "text": "نصي", "choice": "قائمة منسدلة",
    "date": "تاريخ", "time": "وقت", "boolean": "نعم / لا",
    "percentage": "نسبة مئوية", "currency": "مبلغ مالي", "phone": "هاتف / جوال",
}
LABEL_TO_KIND = {v: k for k, v in KIND_LABELS.items()}


def clean_choices(value: str | list[str]) -> list[str]:
    raw = re.split(r"[,،;؛\n]+", value) if isinstance(value, str) else value
    result: list[str] = []
    seen: set[str] = set()
    for item in raw:
        name = str(item).strip()
        if name and name.casefold() not in seen:
            result.append(name)
            seen.add(name.casefold())
    return result


def active_category_conflicts(old: dict[str, Any], new: dict[str, Any], entries: list[dict[str, Any]]) -> dict[str, int]:
    """Never silently remove a category that already has draft records."""
    removed = set(old.get("categories", {})) - set(new.get("categories", {}))
    return {
        category_id: sum(1 for row in entries if row.get("category") == category_id)
        for category_id in removed
        if any(row.get("category") == category_id for row in entries)
    }


def validate_structure(config: dict[str, Any]) -> list[str]:
    problems: list[str] = []
    categories = config.get("categories", {})
    order = config.get("category_order", [])
    if not isinstance(categories, dict) or not categories or not isinstance(order, list):
        return ["يجب وجود قسم واحد على الأقل."]
    if set(categories) != set(order) or len(order) != len(set(order)):
        problems.append("ترتيب الأقسام غير متطابق مع الأقسام الموجودة.")
    seen_sections: set[str] = set()
    for section_id in order:
        category = categories.get(section_id)
        if not isinstance(category, dict):
            problems.append("يوجد قسم غير صالح.")
            continue
        name = str(category.get("label", "")).strip()
        if not name or name.casefold() in seen_sections:
            problems.append("أسماء الأقسام يجب أن تكون غير فارغة وغير مكررة.")
        seen_sections.add(name.casefold())
        people = category.get("people", [])
        if len({str(p).strip().casefold() for p in people}) != len(people) or any(not str(p).strip() for p in people):
            problems.append(f"قائمة الأشخاص في {name} تحتوي أسماء فارغة أو مكررة.")
        ids: set[str] = set()
        labels: set[str] = set()
        for field in category.get("fields", []):
            field_id = str(field.get("id", "")).strip()
            label = str(field.get("label", "")).strip()
            kind = str(field.get("kind", ""))
            if not field_id or field_id in ids or not label or label.casefold() in labels:
                problems.append(f"الحقول في {name} بها عناوين/معرفات غير صالحة أو مكررة.")
            if kind not in KIND_LABELS:
                problems.append(f"نوع غير مدعوم للحقل {label}.")
            choices = field.get("options", [])
            if kind == "choice" and not clean_choices(choices):
                problems.append(f"القائمة {label} لا تحتوي خيارات.")
            ids.add(field_id)
            labels.add(label.casefold())
    return problems


def create_structure_safety_copy(app: Any) -> Path:
    """Back up config + unsaved rows immediately before committing schema changes."""
    parent = Path(app.data_dir) / "structure_backups"
    parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    destination = parent / f"structure_before_{stamp}.json"
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "version": "15.2",
        "config": copy.deepcopy(app.config),
        "draft_entries": copy.deepcopy(app.entries),
        "note": "نسخة أمان قبل تعديل هيكل التقارير. لا تتضمن قاعدة بيانات التقارير المحفوظة.",
    }
    # Build the file outside the final path; os.replace via Path.replace on same filesystem.
    temp = destination.with_suffix(".tmp")
    try:
        temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(destination)
    finally:
        temp.unlink(missing_ok=True)
    return destination


class FieldDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, original: dict[str, Any] | None, existing_labels: list[str]):
        super().__init__(parent)
        self.title("إضافة حقل" if original is None else "تعديل خصائص الحقل")
        self.geometry("575x485")
        self.minsize(520, 440)
        self.configure(bg=PALE)
        self.transient(parent)
        self.resizable(True, True)
        self.result: dict[str, Any] | None = None
        self.original = original
        self.existing_labels = [s.casefold() for s in existing_labels]
        current = original or {}
        self.label_var = tk.StringVar(value=str(current.get("label", "")))
        self.kind_var = tk.StringVar(value=KIND_LABELS.get(current.get("kind", "text"), "نصي"))
        self.required_var = tk.BooleanVar(value=bool(current.get("required", False)))
        self.visible_var = tk.BooleanVar(value=bool(current.get("report_visible", True)))
        self.report_label_var = tk.StringVar(value=str(current.get("report_label", "")))
        content = tk.Frame(self, bg=SURFACE, padx=20, pady=12)
        content.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(content, text="إعدادات الحقل", bg=SURFACE, fg=NAVY, font=("Arial", 16, "bold")).pack(anchor="e", pady=(0, 12))
        self._labeled_entry(content, "اسم الحقل الظاهر", self.label_var)
        row = tk.Frame(content, bg=SURFACE)
        row.pack(fill="x", pady=5)
        tk.Label(row, text="نوع الحقل", bg=SURFACE, fg=INK, font=("Arial", 10, "bold")).pack(side="right", padx=5)
        combo = ttk.Combobox(row, values=list(KIND_LABELS.values()), textvariable=self.kind_var, state="readonly", justify="right", width=26)
        combo.pack(side="left", fill="x", expand=True)
        self._labeled_entry(content, "عنوان مختلف في التقرير (اختياري)", self.report_label_var)
        tk.Checkbutton(content, text="حقل إلزامي في الإدخال", variable=self.required_var, bg=SURFACE,
                       fg=INK, font=("Arial", 10), anchor="e").pack(fill="x", pady=3)
        tk.Checkbutton(content, text="إظهار الحقل في التقرير والتصدير", variable=self.visible_var,
                       bg=SURFACE, fg=INK, font=("Arial", 10), anchor="e").pack(fill="x", pady=3)
        tk.Label(content, text="خيارات القائمة المنسدلة — خيار في كل سطر أو فاصلة بينها", bg=SURFACE, fg=INK,
                 font=("Arial", 10, "bold"), anchor="e").pack(fill="x", pady=(8, 3))
        self.options_text = tk.Text(content, height=4, wrap="word", font=("Arial", 10))
        self.options_text.pack(fill="both", expand=True)
        self.options_text.insert("1.0", "\n".join(current.get("options", [])))
        foot = tk.Frame(self, bg=PALE)
        foot.pack(fill="x", padx=12, pady=(0, 12))
        self._button(foot, "حفظ الحقل", self.save, GREEN).pack(side="right", padx=4)
        self._button(foot, "إلغاء", self.destroy, MUTED).pack(side="right", padx=4)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.focus_set()

    @staticmethod
    def _button(parent: tk.Misc, text: str, command: Any, background: str) -> tk.Button:
        return tk.Button(parent, text=text, command=command, bg=background, fg="white", relief="flat",
                         font=("Arial", 10, "bold"), padx=15, pady=7, cursor="hand2")

    @staticmethod
    def _labeled_entry(parent: tk.Misc, label: str, variable: tk.StringVar) -> None:
        row = tk.Frame(parent, bg=SURFACE)
        row.pack(fill="x", pady=5)
        tk.Label(row, text=label, bg=SURFACE, fg=INK, font=("Arial", 10, "bold")).pack(side="right", padx=5)
        ttk.Entry(row, textvariable=variable, justify="right").pack(side="left", fill="x", expand=True)

    def save(self) -> None:
        label = self.label_var.get().strip()
        if not label or label.casefold() in self.existing_labels:
            messagebox.showwarning("اسم غير صالح", "أدخل اسمًا فريدًا غير فارغ للحقل.", parent=self)
            return
        kind = LABEL_TO_KIND.get(self.kind_var.get(), "text")
        options = clean_choices(self.options_text.get("1.0", "end")) if kind == "choice" else []
        if kind == "choice" and not options:
            messagebox.showwarning("خيارات مطلوبة", "أدخل خيارًا واحدًا على الأقل.", parent=self)
            return
        self.result = {
            "id": self.original["id"] if self.original else f"field_{uuid.uuid4().hex[:12]}",
            "label": label,
            "kind": kind,
            "options": options,
            "required": bool(self.required_var.get()),
            "report_visible": bool(self.visible_var.get()),
            "report_label": self.report_label_var.get().strip(),
        }
        self.destroy()


class StructureStudio(tk.Toplevel):
    """One safe place to manage section labels, people, fields and dropdowns."""

    def __init__(self, app: Any):
        super().__init__(app.root)
        self.app = app
        self.baseline = copy.deepcopy(app.config)
        self.draft = copy.deepcopy(app.config)
        self.category_id = self.draft["category_order"][0]
        self.title("مركز تخصيص النظام | ProTech v15.3")
        self.geometry("1260x790")
        self.minsize(1000, 680)
        self.configure(bg=PALE)
        self.transient(app.root)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self.refresh_all()
        self.bind("<Control-s>", lambda _event: self.commit())
        self.bind("<Escape>", lambda _event: self.close())
        self.grab_set()

    @property
    def category(self) -> dict[str, Any]:
        return self.draft["categories"][self.category_id]

    def _button(self, parent: tk.Misc, text: str, callback: Any, color: str = BLUE, *, side: str = "right") -> tk.Button:
        button = tk.Button(parent, text=text, command=callback, bg=color, fg="white",
                           activebackground=NAVY, activeforeground="white", relief="flat", cursor="hand2",
                           font=("Arial", 10, "bold"), padx=12, pady=9)
        button.pack(side=side, padx=3, pady=3)
        return button

    def _build(self) -> None:
        banner = tk.Frame(self, bg=NAVY, height=98)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="مركز تخصيص وإدارة النظام", bg=NAVY, fg="white",
                 font=("Arial", 21, "bold")).pack(anchor="e", padx=20, pady=(11, 0))
        tk.Label(banner, text="الأقسام   •   الموظفون والفنيون   •   الحقول والأعمدة   •   القوائم المنسدلة",
                 bg=NAVY, fg="#B9D4E4", font=("Arial", 11)).pack(anchor="e", padx=20, pady=2)

        main = tk.Frame(self, bg=PALE)
        main.pack(fill="both", expand=True, padx=12, pady=10)
        side = tk.Frame(main, bg=SURFACE, width=268, bd=1, relief="solid")
        side.pack(side="right", fill="y", padx=(10, 0))
        side.pack_propagate(False)
        tk.Label(side, text="الأقسام", bg=SURFACE, fg=NAVY, font=("Arial", 15, "bold"),
                 anchor="e").pack(fill="x", padx=12, pady=(12, 2))
        self.category_search = tk.StringVar()
        self._search(side, self.category_search, "بحث عن قسم...", self.refresh_categories)
        self.categories_list = tk.Listbox(side, justify="right", font=("Arial", 11), bd=0, highlightthickness=0,
                                         selectmode="browse", activestyle="none", exportselection=False)
        self.categories_list.pack(fill="both", expand=True, padx=9, pady=5)
        self.categories_list.bind("<<ListboxSelect>>", self.on_category_select)
        self.categories_list.bind("<Double-1>", lambda _event: self.edit_category())
        buttons = tk.Frame(side, bg=SURFACE)
        buttons.pack(fill="x", padx=5, pady=4)
        self._button(buttons, "+ قسم", self.add_category, GREEN)
        self._button(buttons, "تعديل", self.edit_category)
        self._button(buttons, "حذف", self.delete_category, RED)
        move = tk.Frame(side, bg=SURFACE)
        move.pack(fill="x", padx=5, pady=(0, 10))
        self._button(move, "↑ نقل لأعلى", lambda: self.move_category(-1), NAVY)
        self._button(move, "↓ نقل لأسفل", lambda: self.move_category(1), NAVY)

        workspace = tk.Frame(main, bg=PALE)
        workspace.pack(side="left", fill="both", expand=True)
        self.context_var = tk.StringVar()
        tk.Label(workspace, textvariable=self.context_var, bg=PALE, fg=NAVY,
                 font=("Arial", 14, "bold"), anchor="e").pack(fill="x", padx=3, pady=(0, 7))
        style = ttk.Style(self)
        style.configure("Modern.TNotebook", background=PALE, borderwidth=0)
        style.configure("Modern.TNotebook.Tab", font=("Arial", 11, "bold"), padding=(20, 11))
        self.tabs = ttk.Notebook(workspace, style="Modern.TNotebook")
        self.tabs.pack(fill="both", expand=True)
        self.people_tab = tk.Frame(self.tabs, bg=SURFACE)
        self.fields_tab = tk.Frame(self.tabs, bg=SURFACE)
        self.options_tab = tk.Frame(self.tabs, bg=SURFACE)
        self.tabs.add(self.people_tab, text="  الموظفون والأسماء  ")
        self.tabs.add(self.fields_tab, text="  الحقول والأعمدة  ")
        self.tabs.add(self.options_tab, text="  خيارات القوائم  ")
        self._people_ui()
        self._fields_ui()
        self._options_ui()

        footer = tk.Frame(self, bg=SURFACE, bd=1, relief="solid")
        footer.pack(fill="x", padx=12, pady=(0, 11))
        self.dirty_var = tk.StringVar()
        tk.Label(footer, textvariable=self.dirty_var, bg=SURFACE, fg=MUTED, font=("Arial", 10),
                 anchor="e").pack(side="left", padx=12)
        self._button(footer, "حفظ وتطبيق", self.commit, GREEN)
        self._button(footer, "تجاهل التغييرات", self.close, MUTED)

    @staticmethod
    def _search(parent: tk.Misc, variable: tk.StringVar, hint: str, command: Any) -> None:
        frame = tk.Frame(parent, bg=SURFACE)
        frame.pack(fill="x", padx=10, pady=(6, 4))
        tk.Label(frame, text="بحث", bg=SURFACE, fg=MUTED).pack(side="right", padx=3)
        ttk.Entry(frame, textvariable=variable, justify="right").pack(side="left", fill="x", expand=True)
        variable.trace_add("write", lambda *_: command())

    def _people_ui(self) -> None:
        self.people_search = tk.StringVar()
        self._search(self.people_tab, self.people_search, "بحث عن موظف", self.refresh_people)
        self.people_count_var = tk.StringVar()
        tk.Label(self.people_tab, textvariable=self.people_count_var, bg=SURFACE, fg=MUTED,
                 anchor="e", font=("Arial", 10)).pack(fill="x", padx=12)
        area = tk.Frame(self.people_tab, bg=SURFACE)
        area.pack(fill="both", expand=True, padx=12, pady=8)
        self.people_list = tk.Listbox(area, justify="right", font=("Arial", 12), bd=1, relief="solid",
                                      selectmode="browse", exportselection=False)
        self.people_list.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(area, command=self.people_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.people_list.configure(yscrollcommand=scrollbar.set)
        self.people_list.bind("<Double-1>", lambda _event: self.edit_person())
        bar = tk.Frame(self.people_tab, bg=SURFACE)
        bar.pack(fill="x", padx=10, pady=(0, 10))
        self._button(bar, "+ إضافة موظف", self.add_person, GREEN)
        self._button(bar, "تعديل الاسم", self.edit_person)
        self._button(bar, "إزالة من القائمة", self.delete_person, RED)
        self._button(bar, "↑", lambda: self.move_person(-1), NAVY)
        self._button(bar, "↓", lambda: self.move_person(1), NAVY)
        tk.Label(self.people_tab, text="إزالة الموظف من الاختيارات لا تحذف سجلات التقارير السابقة.",
                 bg=SURFACE, fg=MUTED, font=("Arial", 9)).pack(anchor="e", padx=13, pady=(0, 9))

    def _fields_ui(self) -> None:
        self.fields_search = tk.StringVar()
        self._search(self.fields_tab, self.fields_search, "بحث عن حقل", self.refresh_fields)
        area = tk.Frame(self.fields_tab, bg=SURFACE)
        area.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("label", "kind", "required", "visible", "count")
        self.fields_tree = ttk.Treeview(area, columns=cols, show="headings", selectmode="browse", height=10)
        for col, title, width in (("label", "اسم العمود", 210), ("kind", "النوع", 125),
                                  ("required", "إلزامي", 68), ("visible", "بالتقرير", 80),
                                  ("count", "الخيارات", 65)):
            self.fields_tree.heading(col, text=title)
            self.fields_tree.column(col, width=width, anchor="e" if col == "label" else "center")
        self.fields_tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(area, command=self.fields_tree.yview)
        scroll.pack(side="right", fill="y")
        self.fields_tree.configure(yscrollcommand=scroll.set)
        self.fields_tree.bind("<Double-1>", lambda _event: self.edit_field())
        bar = tk.Frame(self.fields_tab, bg=SURFACE)
        bar.pack(fill="x", padx=10, pady=(0, 8))
        self._button(bar, "+ إضافة حقل", self.add_field, GREEN)
        self._button(bar, "تعديل الخصائص", self.edit_field)
        self._button(bar, "حذف الحقل", self.delete_field, RED)
        self._button(bar, "↑", lambda: self.move_field(-1), NAVY)
        self._button(bar, "↓", lambda: self.move_field(1), NAVY)
        tk.Label(self.fields_tab, text="يمكن تحديد نوع الحقل وإلزاميته وظهوره في التصدير؛ الحذف لا يمس التقارير المحفوظة.",
                 bg=SURFACE, fg=MUTED, font=("Arial", 9)).pack(anchor="e", padx=13, pady=(0, 9))

    def _options_ui(self) -> None:
        tk.Label(self.options_tab, text="اختر حقلًا من نوع قائمة منسدلة ثم أضف خياراته أو عدّلها أو أعد ترتيبها.",
                 bg=SURFACE, fg=MUTED, anchor="e", font=("Arial", 10)).pack(fill="x", padx=12, pady=13)
        row = tk.Frame(self.options_tab, bg=SURFACE)
        row.pack(fill="x", padx=12)
        tk.Label(row, text="القائمة المنسدلة", bg=SURFACE, fg=INK, font=("Arial", 10, "bold")).pack(side="right", padx=8)
        self.option_field_var = tk.StringVar()
        self.option_field_combo = ttk.Combobox(row, textvariable=self.option_field_var, state="readonly", justify="right", width=35)
        self.option_field_combo.pack(side="right", padx=5)
        self.option_field_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_options())
        self.options_list = tk.Listbox(self.options_tab, justify="right", font=("Arial", 12),
                                       bd=1, relief="solid", selectmode="browse", exportselection=False)
        self.options_list.pack(fill="both", expand=True, padx=12, pady=12)
        self.options_list.bind("<Double-1>", lambda _event: self.edit_option())
        bar = tk.Frame(self.options_tab, bg=SURFACE)
        bar.pack(fill="x", padx=10, pady=(0, 12))
        self._button(bar, "+ خيار", self.add_option, GREEN)
        self._button(bar, "تعديل", self.edit_option)
        self._button(bar, "حذف", self.delete_option, RED)
        self._button(bar, "↑", lambda: self.move_option(-1), NAVY)
        self._button(bar, "↓", lambda: self.move_option(1), NAVY)

    def _dirty(self) -> None:
        dirty = self.draft != self.baseline
        self.dirty_var.set("● تعديلات غير محفوظة — Ctrl+S للحفظ" if dirty else "✓ لا توجد تغييرات غير محفوظة")
        self.context_var.set(f"{self.category['label']}    |    {len(self.category['people'])} اسم    |    {len(self.category['fields'])} حقل")

    def refresh_all(self) -> None:
        self.refresh_categories()
        self.refresh_people()
        self.refresh_fields()
        self.refresh_option_fields()
        self._dirty()

    def refresh_categories(self) -> None:
        selection = self.category_id
        text = self.category_search.get().strip().casefold()
        self.visible_categories = [cat for cat in self.draft["category_order"]
                                   if text in self.draft["categories"][cat]["label"].casefold()]
        self.categories_list.delete(0, "end")
        for cat in self.visible_categories:
            self.categories_list.insert("end", self.draft["categories"][cat]["label"])
        if selection in self.visible_categories:
            self.categories_list.selection_set(self.visible_categories.index(selection))

    def on_category_select(self, _event: Any = None) -> None:
        choice = self.categories_list.curselection()
        if not choice:
            return
        selected = self.visible_categories[int(choice[0])]
        if selected == self.category_id:
            return
        self.category_id = selected
        self.refresh_people()
        self.refresh_fields()
        self.refresh_option_fields()
        self._dirty()

    def add_category(self) -> None:
        name = simpledialog.askstring("إضافة قسم", "اسم القسم الجديد:", parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {s["label"].casefold() for s in self.draft["categories"].values()}:
            messagebox.showwarning("اسم القسم", "اسم القسم مطلوب ويجب ألا يكون مكررًا.", parent=self)
            return
        new_id = f"category_{uuid.uuid4().hex[:12]}"
        self.draft["categories"][new_id] = {"label": name, "person_label": "اسم الموظف", "total_label": "الإجمالي", "people": [], "fields": []}
        self.draft["category_order"].append(new_id)
        self.category_search.set("")
        self.category_id = new_id
        self.refresh_all()

    def edit_category(self) -> None:
        category = self.category
        name = simpledialog.askstring("تعديل القسم", "اسم القسم:", initialvalue=category["label"], parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {v["label"].casefold() for key, v in self.draft["categories"].items() if key != self.category_id}:
            messagebox.showwarning("اسم القسم", "اسم القسم فارغ أو مكرر.", parent=self)
            return
        person_label = simpledialog.askstring("مسمى الأشخاص", "عنوان عمود الأشخاص (مثال: الموظف أو الفني):", initialvalue=category["person_label"], parent=self)
        if person_label is None:
            return
        total_label = simpledialog.askstring("الإجمالي", "عنوان عمود الإجمالي:", initialvalue=category["total_label"], parent=self)
        if total_label is None:
            return
        category.update(label=name, person_label=person_label.strip() or "الاسم", total_label=total_label.strip() or "الإجمالي")
        self.refresh_all()

    def delete_category(self) -> None:
        if len(self.draft["category_order"]) < 2:
            messagebox.showwarning("غير مسموح", "يجب الإبقاء على قسم واحد على الأقل.", parent=self)
            return
        entries = [row for row in self.app.entries if row.get("category") == self.category_id]
        if entries:
            messagebox.showwarning("قسم مستخدم", f"القسم مرتبط بـ {len(entries)} سجل في التقرير الجاري. احفظ التقرير أو انقل سجلاته أولًا حتى لا تختفي.", parent=self)
            return
        if not messagebox.askyesno("حذف القسم", f"حذف قسم «{self.category['label']}» من الإدخال القادم؟\nلن تُحذف التقارير القديمة.", parent=self):
            return
        self.draft["category_order"].remove(self.category_id)
        del self.draft["categories"][self.category_id]
        self.category_id = self.draft["category_order"][0]
        self.refresh_all()

    def move_category(self, direction: int) -> None:
        order = self.draft["category_order"]
        index = order.index(self.category_id)
        target = index + direction
        if 0 <= target < len(order):
            order[index], order[target] = order[target], order[index]
            self.refresh_all()

    def refresh_people(self) -> None:
        text = self.people_search.get().strip().casefold()
        self.visible_people = [(i, name) for i, name in enumerate(self.category["people"]) if text in name.casefold()]
        self.people_list.delete(0, "end")
        for _i, name in self.visible_people:
            self.people_list.insert("end", name)
        self.people_count_var.set(f"{len(self.category['people'])} اسم — {self.category.get('person_label', 'الاسم')}")
        self._dirty()

    def selected_person_index(self) -> int | None:
        selected = self.people_list.curselection()
        return self.visible_people[int(selected[0])][0] if selected else None

    def add_person(self) -> None:
        name = simpledialog.askstring("إضافة اسم", "اسم الموظف / الشخص الجديد:", parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {s.casefold() for s in self.category["people"]}:
            messagebox.showwarning("اسم غير صالح", "الاسم مطلوب ويجب ألا يكون مكررًا في القسم.", parent=self)
            return
        self.category["people"].append(name)
        self.refresh_people()

    def edit_person(self) -> None:
        index = self.selected_person_index()
        if index is None:
            messagebox.showinfo("تحديد", "حدد موظفًا أولًا.", parent=self)
            return
        name = simpledialog.askstring("تعديل الاسم", "الاسم الجديد:", initialvalue=self.category["people"][index], parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {s.casefold() for i, s in enumerate(self.category["people"]) if i != index}:
            messagebox.showwarning("اسم غير صالح", "الاسم مطلوب ويجب ألا يكون مكررًا.", parent=self)
            return
        self.category["people"][index] = name
        self.refresh_people()

    def delete_person(self) -> None:
        index = self.selected_person_index()
        if index is None:
            return
        name = self.category["people"][index]
        if messagebox.askyesno("إزالة موظف", f"إزالة «{name}» من قائمة الاختيار؟\nلن تُحذف السجلات السابقة المسجلة باسمه.", parent=self):
            del self.category["people"][index]
            self.refresh_people()

    def move_person(self, direction: int) -> None:
        index = self.selected_person_index()
        if index is None:
            return
        destination = index + direction
        names = self.category["people"]
        if 0 <= destination < len(names):
            names[index], names[destination] = names[destination], names[index]
            self.people_search.set("")
            self.refresh_people()
            self.people_list.selection_set(destination)

    def refresh_fields(self) -> None:
        text = self.fields_search.get().strip().casefold()
        previous = self.fields_tree.selection()
        self.fields_tree.delete(*self.fields_tree.get_children())
        for field in self.category["fields"]:
            if text not in field["label"].casefold():
                continue
            self.fields_tree.insert("", "end", iid=field["id"], values=(
                field["label"], KIND_LABELS.get(field["kind"], ""),
                "نعم" if field.get("required") else "لا", "نعم" if field.get("report_visible", True) else "لا",
                len(field.get("options", [])) if field["kind"] == "choice" else "—",
            ))
        if previous and self.fields_tree.exists(previous[0]):
            self.fields_tree.selection_set(previous[0])
        self._dirty()

    def selected_field(self) -> dict[str, Any] | None:
        chosen = self.fields_tree.selection()
        return next((field for field in self.category["fields"] if field["id"] == chosen[0]), None) if chosen else None

    def add_field(self) -> None:
        dialog = FieldDialog(self, None, [field["label"] for field in self.category["fields"]])
        self.wait_window(dialog)
        if dialog.result:
            self.category["fields"].append(dialog.result)
            self.refresh_fields()
            self.fields_tree.selection_set(dialog.result["id"])
            self.refresh_option_fields()

    def edit_field(self) -> None:
        field = self.selected_field()
        if not field:
            messagebox.showinfo("تحديد", "حدد حقلًا أولًا.", parent=self)
            return
        dialog = FieldDialog(self, field, [item["label"] for item in self.category["fields"] if item["id"] != field["id"]])
        self.wait_window(dialog)
        if dialog.result:
            if dialog.result["kind"] != field["kind"] and any(
                field["id"] in row.get("values", {}) for row in self.app.entries if row.get("category") == self.category_id
            ):
                messagebox.showwarning("حماية البيانات", "لا يمكن تغيير نوع هذا الحقل بسبب وجود قيم له في التقرير الجاري. أنشئ حقلًا جديدًا بالنوع المطلوب.", parent=self)
                return
            field.update(dialog.result)
            self.refresh_fields()
            self.refresh_option_fields()

    def delete_field(self) -> None:
        field = self.selected_field()
        if not field:
            return
        count = sum(1 for row in self.app.entries if row.get("category") == self.category_id and field["id"] in row.get("values", {}))
        suffix = f"\nيوجد {count} سجل حالي يحتوي قيمة لهذا الحقل؛ ستظل القيم محفوظة داخليًا ولن تظهر في التقرير الجديد." if count else ""
        if messagebox.askyesno("حذف حقل", f"إزالة عمود «{field['label']}» من تصميم الإدخال والتقارير الجديدة؟{suffix}\nالتقارير السابقة لا تتغير.", parent=self):
            self.category["fields"].remove(field)
            self.refresh_fields()
            self.refresh_option_fields()

    def move_field(self, direction: int) -> None:
        field = self.selected_field()
        if field is None:
            return
        fields = self.category["fields"]
        index = fields.index(field)
        destination = index + direction
        if 0 <= destination < len(fields):
            fields[index], fields[destination] = fields[destination], fields[index]
            self.fields_search.set("")
            self.refresh_fields()
            self.fields_tree.selection_set(field["id"])

    def refresh_option_fields(self) -> None:
        choices = [field for field in self.category["fields"] if field["kind"] == "choice"]
        self.choice_ids_by_label = {field["label"]: field["id"] for field in choices}
        self.option_field_combo.configure(values=list(self.choice_ids_by_label))
        if self.option_field_var.get() not in self.choice_ids_by_label:
            self.option_field_var.set(choices[0]["label"] if choices else "")
        self.refresh_options()

    def current_choice_field(self) -> dict[str, Any] | None:
        choice_id = self.choice_ids_by_label.get(self.option_field_var.get())
        return next((field for field in self.category["fields"] if field["id"] == choice_id), None)

    def refresh_options(self) -> None:
        self.options_list.delete(0, "end")
        field = self.current_choice_field()
        if field:
            for option in field["options"]:
                self.options_list.insert("end", option)
        self._dirty()

    def _selected_option_index(self) -> int | None:
        selected = self.options_list.curselection()
        return int(selected[0]) if selected else None

    def add_option(self) -> None:
        field = self.current_choice_field()
        if not field:
            messagebox.showinfo("لا توجد قائمة", "أنشئ حقلًا من نوع «قائمة منسدلة» في تبويب الحقول أولًا.", parent=self)
            return
        name = simpledialog.askstring("خيار جديد", "اسم الخيار:", parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {s.casefold() for s in field["options"]}:
            messagebox.showwarning("خيار غير صالح", "الخيار فارغ أو مكرر.", parent=self)
            return
        field["options"].append(name)
        self.refresh_options()
        self.refresh_fields()

    def edit_option(self) -> None:
        field = self.current_choice_field()
        index = self._selected_option_index()
        if field is None or index is None:
            return
        name = simpledialog.askstring("تعديل خيار", "النص الجديد:", initialvalue=field["options"][index], parent=self)
        if name is None:
            return
        name = name.strip()
        if not name or name.casefold() in {s.casefold() for i, s in enumerate(field["options"]) if i != index}:
            messagebox.showwarning("خيار غير صالح", "الخيار فارغ أو مكرر.", parent=self)
            return
        field["options"][index] = name
        self.refresh_options()
        self.options_list.selection_set(index)

    def delete_option(self) -> None:
        field = self.current_choice_field()
        index = self._selected_option_index()
        if field is None or index is None:
            return
        if len(field["options"]) <= 1:
            messagebox.showwarning("لا يمكن", "أضف خيارًا بديلًا أولًا؛ القائمة يجب أن تحتوي خيارًا واحدًا على الأقل.", parent=self)
            return
        if messagebox.askyesno("حذف خيار", f"حذف خيار «{field['options'][index]}» من الاختيارات القادمة؟\nالقيم التاريخية لا تتغير.", parent=self):
            del field["options"][index]
            self.refresh_options()
            self.refresh_fields()

    def move_option(self, direction: int) -> None:
        field = self.current_choice_field()
        index = self._selected_option_index()
        if field is None or index is None:
            return
        target = index + direction
        if 0 <= target < len(field["options"]):
            field["options"][index], field["options"][target] = field["options"][target], field["options"][index]
            self.refresh_options()
            self.options_list.selection_set(target)

    def commit(self) -> None:
        if self.draft == self.baseline:
            self.destroy()
            return
        errors = validate_structure(self.draft)
        if errors:
            messagebox.showwarning("تعذر الحفظ", "\n".join(errors[:5]), parent=self)
            return
        conflicts = active_category_conflicts(self.app.config, self.draft, self.app.entries)
        if conflicts:
            names = [self.app.config["categories"][key]["label"] for key in conflicts]
            messagebox.showwarning("حماية التقرير الجاري", "لا يمكن حذف أقسام فيها سجلات حالية: " + "، ".join(names), parent=self)
            return
        try:
            backup_path = create_structure_safety_copy(self.app)
        except OSError as exc:
            messagebox.showerror("تعذر إنشاء نسخة الأمان", f"لن تُطبّق التعديلات دون نسخة أمان:\n{exc}", parent=self)
            return
        result = self.app.apply_configuration(copy.deepcopy(self.draft))
        if result is False:
            return
        messagebox.showinfo("تم تطبيق التعديلات", f"تم تحديث القوائم والحقول دون تعديل التقارير المحفوظة.\nنسخة الأمان:\n{backup_path}", parent=self)
        self.destroy()

    def close(self) -> None:
        if self.draft != self.baseline and not messagebox.askyesno("تجاهل التعديلات", "هناك تغييرات لم تُحفظ. هل تريد إغلاق النافذة دون تطبيقها؟", parent=self):
            return
        self.destroy()
