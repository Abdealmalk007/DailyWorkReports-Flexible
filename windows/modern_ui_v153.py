"""Modern v15.3 navigation surface. No database migrations or new persistent settings."""
import tkinter as tk
from tkinter import ttk

NAVY = "#183E43"
TEAL = "#25646B"
PALE = "#F2F7F6"
WHITE = "#FFFFFF"
INK = "#20383D"
MUTED = "#657E81"
GREEN = "#228567"
BORDER = "#D8E7E4"

class ModernHome(tk.Toplevel):
    """Non-destructive navigation home: all actions delegate to existing ERP commands."""
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("الواجهة الرئيسية • ProTech 15.5.0")
        self.geometry("1050x760")
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self.minsize(760, 560)
        self.configure(bg=PALE)
        self.transient(app.root)
        self._build()
        self.lift()
        self.focus_set()
        self.bind("<Escape>", lambda _e: self.destroy())

    def _card(self, parent, label, value, hint, color):
        frame = tk.Frame(parent, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(side="right", fill="both", expand=True, padx=6)
        tk.Label(frame, text=label, bg=WHITE, fg=MUTED, font=("Arial", 11), anchor="e").pack(fill="x", padx=16, pady=(16, 4))
        tk.Label(frame, text=str(value), bg=WHITE, fg=color, font=("Arial", 24, "bold"), anchor="e").pack(fill="x", padx=16)
        tk.Label(frame, text=hint, bg=WHITE, fg=MUTED, font=("Arial", 9), anchor="e").pack(fill="x", padx=16, pady=(0, 16))
        return frame

    def _action(self, parent, title, hint, command, color=TEAL):
        shell = tk.Frame(parent, bg=WHITE, highlightthickness=1, highlightbackground=BORDER)
        shell.pack(fill="x", pady=5)
        b = tk.Button(shell, text=title, command=lambda: self._invoke(command), bg=color, fg=WHITE,
                      activebackground=NAVY, activeforeground=WHITE, borderwidth=0, relief="flat",
                      font=("Arial", 11, "bold"), padx=16, pady=11, cursor="hand2", width=23, anchor="e")
        b.pack(side="right", padx=8, pady=8)
        tk.Label(shell, text=hint, bg=WHITE, fg=MUTED, font=("Arial", 10), anchor="e",
                 justify="right").pack(side="right", fill="x", expand=True, padx=14)
        return b

    def _invoke(self, cb):
        self.destroy()
        self.app.root.after_idle(cb)

    def _build(self):
        banner=tk.Frame(self,bg=NAVY,height=122); banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="مرحبًا بك في تقارير العمل اليومية", bg=NAVY, fg=WHITE,
                 font=("Arial",22,"bold"), anchor="e").pack(fill="x", padx=28, pady=(18,1))
        tk.Label(banner,text="ProTech Modern UI 15.5.0  |  إدارة أبسط، وصول أسرع، وبيانات محفوظة",
                 bg=NAVY,fg="#C7E5E1",font=("Arial",10),anchor="e").pack(fill="x",padx=28,pady=(0,12))
        outer=tk.Frame(self,bg=PALE);outer.pack(fill="both",expand=True,padx=20,pady=15)
        try: db_count=self.app.database.count(self.app.current_company_id)
        except Exception: db_count=0
        try: active=len(self.app.entries)
        except Exception: active=0
        cats=self.app.config.get("categories", {})
        staff=len({p for c in cats.values() for p in c.get("people",[]) if str(p).strip()})
        stats=tk.Frame(outer,bg=PALE);stats.pack(fill="x",pady=(0,12))
        self._card(stats,"التقارير المحفوظة",db_count,"من قاعدة البيانات الحالية",TEAL)
        self._card(stats,"سجلات التقرير",active,"في المسودة الجارية",GREEN)
        self._card(stats,"الموظفون",staff,"أضف الموظفين من إدارة الأقسام" if staff == 0 else "أسماء فريدة مسجلة",TEAL)
        self._card(stats,"الأقسام",len(cats),"قابلة للتخصيص",GREEN)
        if staff == 0:
            tk.Label(outer, text="لا توجد أسماء موظفين مسبقة. أضف موظفيك من (الموظفون والأقسام).", bg=PALE, fg=TEAL, font=("Arial", 10), anchor="e").pack(fill="x", pady=(0,8))
        tk.Label(outer,text="الإجراءات السريعة",bg=PALE,fg=INK,font=("Arial",15,"bold"),
                 anchor="e").pack(fill="x",pady=(5,8))
        area=tk.Frame(outer,bg=PALE);area.pack(fill="both",expand=True)
        right=tk.Frame(area,bg=PALE);right.pack(side="right",fill="both",expand=True,padx=(6,0))
        left=tk.Frame(area,bg=PALE);left.pack(side="left",fill="both",expand=True,padx=(0,6))
        self._action(right,"＋ إدخال سجل","إضافة العمل إلى التقرير الجاري",self.app.add_or_update_entry,GREEN)
        self._action(right,"💾 حفظ التقرير","حفظ بيانات العمل في قاعدة البيانات",self.app.save_to_history,TEAL)
        self._action(right,"👥 الموظفون والأقسام","إضافة / تعديل / إزالة أسماء وأقسام",self.app.open_structure_studio)
        self._action(left,"📋 سجل التقارير","بحث واستعراض التقارير السابقة",self.app.show_history)
        self._action(left,"📊 المؤشرات","لوحة إحصائيات العمل",self.app.show_dashboard)
        self._action(left,"⚙ إعدادات المظهر","خيارات الألوان وحجم الخط",self.app.show_appearance_manager)
        tk.Label(outer,text="Developer: Abu Omar | protech.2025.2026@gmail.com | +966503276927",
                 bg=PALE,fg=TEAL,font=("Arial",10,"bold"),anchor="e").pack(fill="x",pady=(3,3))
        tk.Label(outer,text="اختصارات: Ctrl+S للحفظ  |  Ctrl+N لتقرير جديد  |  Ctrl+F للبحث",
                 bg=PALE,fg=MUTED,font=("Arial",9),anchor="e").pack(fill="x",pady=(12,2))
