import os
import datetime
import flet as ft
from openpyxl import load_workbook

# ---------- Storage (inside the app's private folder) ----------
DATA_DIR = os.getenv("FLET_APP_STORAGE_DATA") or os.path.dirname(os.path.abspath(__file__))
os.makedirs(DATA_DIR, exist_ok=True)

def dp(name):
    return os.path.join(DATA_DIR, name)

XLSX_PATH = dp("shifts.xlsx")
CRED_FILE = dp("user_credentials.txt")
SETTINGS_FILE = dp("app_settings.txt")
SHIFT_STATE_FILE = dp("last_shift_answered.txt")

SHEET_NAME = "جدول الشفتات 2027"
CURRENT_YEAR = 2027
PROMPT_FROM = datetime.date(CURRENT_YEAR, 1, 1)
PROMPT_EVERY_HOURS = 24

# ---------- Colors ----------
C_BLUE = "#2563EB"
C_PURPLE = "#7C3AED"
C_SLATE = "#334155"
C_AMBER = "#D97706"
C_GREEN = "#059669"
C_RED = "#DC2626"

THEMES = {
    "light": {"bg": "#F8FAFC", "card_bg": "#FFFFFF", "text": "#0F172A",
              "sub_text": "#64748B", "btn_bg": "#E2E8F0", "btn_fg": "#0F172A",
              "primary": "#2563EB", "field": "#FFFFFF", "border": "#CBD5E1"},
    "dark": {"bg": "#0F172A", "card_bg": "#1E293B", "text": "#F8FAFC",
             "sub_text": "#94A3B8", "btn_bg": "#334155", "btn_fg": "#F8FAFC",
             "primary": "#38BDF8", "field": "#0F172A", "border": "#475569"},
    "liquid_glass": {"bg": "#030712", "card_bg": "#0B1528", "text": "#F0F9FF",
                     "sub_text": "#7DD3FC", "btn_bg": "#112240", "btn_fg": "#38BDF8",
                     "primary": "#0284C7", "field": "#071120", "border": "#1E3A8A"},
}

# ---------- Texts (no Arabic shaping needed in Flet) ----------
TEXTS = {
    "en": {
        "main_menu": "Main Menu", "settings": "Settings",
        "record_btn": "Record Shift", "payroll_btn": "Calculate Payroll",
        "exit": "Exit", "back": "Back", "later": "Later",
        "welcome": "Welcome!", "create_creds": "Create account credentials:",
        "username": "Username:", "password": "Password:",
        "save_creds": "Save Credentials", "login": "Login",
        "show_pass": "Show Password",
        "pass_lock": "Enable Password Lock",
        "change_creds": "Change Credentials",
        "new_user": "New Username:", "new_pass": "New Password:",
        "update_creds": "Update Credentials", "language": "Language:",
        "record_title": "Record Shift", "date": "Date:",
        "select_shift": "Select Shift:", "morning": "Morning (1st)",
        "evening": "Evening (2nd)", "night": "Night (3rd)",
        "vacation": "Vacation",
        "prompt_title": "What's your shift today?",
        "payroll_title": "Calculate Payroll",
        "start": "Start Date (MM-DD):", "end": "End Date (MM-DD):",
        "rate2": "2nd Shift Rate:", "rate3": "3rd Shift Rate:",
        "base": "Base Salary:", "quality": "Quality Bonus:",
        "seniority": "Seniority Bonus:", "calc": "Calculate Report",
        "report_title": "Payroll Report", "from_date": "From:", "to_date": "To:",
        "total_days": "Total Days:", "n_morning": "Morning (1st):",
        "n_evening": "Evening (2nd):", "n_night": "Night (3rd):",
        "n_vac": "Vacations:", "shift_bonus": "Shift Bonus:", "net_pay": "NET PAY:",
        "theme_mode": "Theme Style:",
        "import_xlsx": "Import Excel File", "export_xlsx": "Export Excel File",
        "no_file": "Import your Excel file first (Settings).",
        "bad_login": "Invalid login!", "empty_fields": "Fields cannot be empty!",
        "saved": "Saved!", "today_missing": "Today not found in sheet!",
        "bad_date": "Invalid date (MM-DD).", "start_after_end": "Start > End!",
    },
    "ar": {
        "main_menu": "القائمة الرئيسية", "settings": "الإعدادات",
        "record_btn": "تسجيل الشفت", "payroll_btn": "حساب المرتب",
        "exit": "خروج", "back": "رجوع", "later": "لاحقًا",
        "welcome": "أهلاً بك!", "create_creds": "إنشاء بيانات الدخول:",
        "username": "اسم المستخدم:", "password": "كلمة المرور:",
        "save_creds": "حفظ البيانات", "login": "دخول",
        "show_pass": "إظهار كلمة المرور",
        "pass_lock": "تفعيل قفل كلمة المرور",
        "change_creds": "تغيير بيانات الدخول",
        "new_user": "اسم مستخدم جديد:", "new_pass": "كلمة مرور جديدة:",
        "update_creds": "تحديث البيانات", "language": "اللغة:",
        "record_title": "تسجيل الشفت", "date": "التاريخ:",
        "select_shift": "اختر الشفت:", "morning": "صباحي (الأول)",
        "evening": "مسائي (الثاني)", "night": "ليلي (الثالث)",
        "vacation": "إجازة",
        "prompt_title": "شفتك إيه النهاردة؟",
        "payroll_title": "حساب المرتب",
        "start": "تاريخ البداية (شهر-يوم):", "end": "تاريخ النهاية (شهر-يوم):",
        "rate2": "سعر الشفت الثاني:", "rate3": "سعر الشفت الثالث:",
        "base": "المرتب الأساسي:", "quality": "حافز الجودة:",
        "seniority": "حافز الأقدمية:", "calc": "احسب التقرير",
        "report_title": "تقرير المرتب", "from_date": "من:", "to_date": "إلى:",
        "total_days": "إجمالي الأيام:", "n_morning": "صباحي:",
        "n_evening": "مسائي:", "n_night": "ليلي:",
        "n_vac": "الإجازات:", "shift_bonus": "بدل الشفتات:", "net_pay": "صافي المرتب:",
        "theme_mode": "نمط المظهر:",
        "import_xlsx": "استيراد ملف الإكسل", "export_xlsx": "تصدير ملف الإكسل",
        "no_file": "استورد ملف الإكسل الأول (من الإعدادات).",
        "bad_login": "بيانات الدخول غلط!", "empty_fields": "الخانات فاضية!",
        "saved": "تم الحفظ!", "today_missing": "النهاردة مش موجود في الشيت!",
        "bad_date": "تاريخ غلط (شهر-يوم).", "start_after_end": "البداية بعد النهاية!",
    },
}

# ---------- App state + settings ----------
STATE = {"theme": "liquid_glass", "require_password": True, "lang": "ar"}

def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        if len(lines) >= 1 and lines[0] in THEMES: STATE["theme"] = lines[0]
        if len(lines) >= 2: STATE["require_password"] = lines[1].lower() == "true"
        if len(lines) >= 3 and lines[2] in TEXTS: STATE["lang"] = lines[2]
    except Exception:
        pass

def save_settings():
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            f.write(f"{STATE['theme']}\n{STATE['require_password']}\n{STATE['lang']}")
    except Exception:
        pass

def tr(key):
    return TEXTS[STATE["lang"]].get(key, key)

def theme():
    return THEMES[STATE["theme"]]

def is_rtl():
    return STATE["lang"] == "ar"

def date_label():
    d = datetime.date.today().strftime("%Y-%m-%d")
    return f"{tr('date')} {d}"

# ---------- Credentials ----------
def has_creds():
    return os.path.exists(CRED_FILE)

def read_creds():
    try:
        with open(CRED_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        return (lines[0] if len(lines) > 0 else ""), (lines[1] if len(lines) > 1 else "")
    except Exception:
        return "", ""

def write_creds(u, p):
    with open(CRED_FILE, "w", encoding="utf-8") as f:
        f.write(f"{u}\n{p}")

# ---------- Daily prompt state ----------
def last_answered():
    try:
        with open(SHIFT_STATE_FILE, "r", encoding="utf-8") as f:
            return datetime.datetime.fromisoformat(f.read().strip())
    except Exception:
        return None

def mark_answered():
    try:
        with open(SHIFT_STATE_FILE, "w", encoding="utf-8") as f:
            f.write(datetime.datetime.now().isoformat(timespec="seconds"))
    except Exception:
        pass

def prompt_is_due():
    last = last_answered()
    if last is None:
        return True
    return datetime.datetime.now() - last >= datetime.timedelta(hours=PROMPT_EVERY_HOURS)

# ---------- Excel logic (returns results, no popups) ----------
def load_excel():
    """returns (wb, sheet) or (None, error_key)"""
    if not os.path.exists(XLSX_PATH):
        return None, "no_file"
    try:
        wb = load_workbook(XLSX_PATH, data_only=False)
        return wb, wb[SHEET_NAME]
    except Exception as e:
        return None, f"Excel error: {e}"

def record_shift(val, col):
    """returns (ok, message)"""
    today = datetime.date.today()
    wb, sheet = load_excel()
    if wb is None:
        return False, tr(sheet) if sheet in TEXTS["en"] else sheet
    target = next((r for r in range(2, sheet.max_row + 1)
                   if sheet.cell(row=r, column=3).value
                   and str(sheet.cell(row=r, column=3).value).startswith(str(today))), None)
    if not target:
        return False, tr("today_missing")
    sheet.cell(row=target, column=col).value = val
    try:
        wb.save(XLSX_PATH)
    except Exception as e:
        return False, f"Save error: {e}"
    mark_answered()
    return True, tr("saved")

def calc_payroll(start_txt, end_txt, rates):
    """rates: dict with r2, r3, base, quality, seniority (floats). returns (data, error)"""
    try:
        s = datetime.datetime.strptime(f"{CURRENT_YEAR}-{start_txt.strip()}", "%Y-%m-%d").date()
        e = datetime.datetime.strptime(f"{CURRENT_YEAR}-{end_txt.strip()}", "%Y-%m-%d").date()
    except ValueError:
        return None, tr("bad_date")
    if s > e:
        return None, tr("start_after_end")
    wb, sheet = load_excel()
    if wb is None:
        return None, tr(sheet) if sheet in TEXTS["en"] else sheet
    c = {"total": 0, "1st": 0, "2nd": 0, "3rd": 0, "vac": 0}
    for r in range(2, sheet.max_row + 1):
        val = sheet.cell(row=r, column=3).value
        if val and isinstance(val, (datetime.datetime, datetime.date)):
            d = val.date() if isinstance(val, datetime.datetime) else val
            if s <= d <= e:
                c["total"] += 1
                s_val = str(sheet.cell(row=r, column=4).value or "").strip()
                v_val = str(sheet.cell(row=r, column=5).value or "").strip().lower()
                if v_val == "vacation": c["vac"] += 1
                elif s_val in c: c[s_val] += 1
    b2 = c["2nd"] * rates["r2"]
    b3 = c["3rd"] * rates["r3"]
    net = rates["base"] + rates["quality"] + rates["seniority"] + b2 + b3
    return {"start": s, "end": e, "total": c["total"], "m1": c["1st"], "m2": c["2nd"],
            "m3": c["3rd"], "vac": c["vac"], "b2": b2, "b3": b3, "base": rates["base"],
            "quality": rates["quality"], "seniority": rates["seniority"], "net": net}, None


VIEWS = {}

def view(name):
    def deco(fn):
        VIEWS[name] = fn
        return fn
    return deco


class App:
    def __init__(self, page):
        self.page = page
        self.current = ""
        self.logged_in = False
        self.report = None

    def show(self, name):
        self.current = name
        t = theme()
        self.page.theme_mode = ft.ThemeMode.LIGHT if STATE["theme"] == "light" else ft.ThemeMode.DARK
        self.page.bgcolor = t["bg"]
        self.page.clean()
        self.page.add(ft.SafeArea(content=VIEWS[name](self), expand=True))
        self.page.update()

    def toast(self, msg):
        sb = ft.SnackBar(ft.Text(msg))
        try:
            self.page.show_dialog(sb)
        except AttributeError:
            self.page.overlay.append(sb)
            sb.open = True
            self.page.update()

    async def exit(self, e=None):
        r = self.page.window.destroy()
        if hasattr(r, "__await__"):
            await r


# ---------- UI helpers ----------
def label(text, size=14, bold=False, key="text"):
    return ft.Text(
        text, size=size, color=theme()[key],
        weight=ft.FontWeight.BOLD if bold else ft.FontWeight.NORMAL,
        text_align=ft.TextAlign.RIGHT if is_rtl() else ft.TextAlign.LEFT,
    )


def col(controls, spacing=8):
    return ft.Column(controls, spacing=spacing,
                     horizontal_alignment=ft.CrossAxisAlignment.STRETCH)


def card(controls, spacing=8):
    t = theme()
    return ft.Container(
        content=col(controls, spacing),
        bgcolor=t["card_bg"], border_radius=16,
        border=ft.Border.all(1, t["border"]),
        padding=ft.Padding.all(14),
    )


def button(text, on_click, bg=None, fg="#FFFFFF", height=48, radius=14,
           icon=None, size=14, expand=False):
    t = theme()
    items = [ft.Text(text, size=size, weight=ft.FontWeight.BOLD, color=fg)]
    if icon:
        ic = ft.Icon(icon, color=fg, size=size + 6)
        items = items + [ic] if is_rtl() else [ic] + items
    return ft.Container(
        content=ft.Row(items, alignment=ft.MainAxisAlignment.CENTER, spacing=8),
        bgcolor=bg or t["primary"], height=height, border_radius=radius,
        alignment=ft.Alignment.CENTER, on_click=on_click, ink=True, expand=expand,
    )


def field(password=False, value="", on_submit=None, keyboard=None):
    t = theme()
    return ft.TextField(
        value=value, password=password, on_submit=on_submit,
        keyboard_type=keyboard,
        color=t["text"], bgcolor=t["field"], border_color=t["border"],
        focused_border_color=t["primary"], border_radius=12,
        text_align=ft.TextAlign.RIGHT if is_rtl() else ft.TextAlign.LEFT,
    )


def switch_row(text, value, on_change):
    t = theme()
    sw = ft.Switch(value=value, on_change=on_change, active_color=t["primary"])
    items = [ft.Container(label(text, 14, True), expand=True), sw]
    if is_rtl():
        items.reverse()
    return ft.Row(items, alignment=ft.MainAxisAlignment.SPACE_BETWEEN)


def screen(app, title, body, back=None, back_key="back", exit_btn=False, subtitle=None):
    t = theme()
    head = [label(title, 22, True)]
    if subtitle:
        head.append(label(subtitle, 13, False, "sub_text"))
    bottom = []
    if back:
        bottom.append(button(tr(back_key), lambda e: app.show(back),
                             bg=t["btn_bg"], fg=t["btn_fg"], height=46, expand=True))
    if exit_btn:
        bottom.append(button(tr("exit"), app.exit, bg=C_RED, height=46, expand=True))
    if is_rtl():
        bottom.reverse()
    main_col = ft.Column(head + body, spacing=10, scroll=ft.ScrollMode.AUTO, expand=True,
                         horizontal_alignment=ft.CrossAxisAlignment.STRETCH)
    rows = [main_col]
    if bottom:
        rows.append(ft.Row(bottom, spacing=10))
    return ft.Container(content=ft.Column(rows, expand=True),
                        padding=ft.Padding.all(16), expand=True)


# ---------- Setup page ----------
@view("setup")
def v_setup(app):
    u = field()
    p = field(password=True)

    def save(e):
        if not u.value.strip() or not p.value.strip():
            app.toast(tr("empty_fields"))
            return
        write_creds(u.value.strip(), p.value.strip())
        app.toast(tr("saved"))
        app.show("menu")

    p.on_submit = save
    body = [card([
        label(tr("username"), 14, True), u,
        label(tr("password"), 14, True), p,
        button(tr("save_creds"), save, height=46),
    ], spacing=10)]
    return screen(app, tr("welcome"), body, subtitle=tr("create_creds"))


# ---------- Login page ----------
@view("login")
def v_login(app):
    u = field()
    p = field(password=True)
    attempts = {"n": 3}

    def toggle(e):
        p.password = not e.control.value
        p.update()

    async def check(e):
        cu, cp = read_creds()
        if u.value.strip() == cu and p.value.strip() == cp:
            app.show("menu")
            return
        attempts["n"] -= 1
        app.toast(f"{tr('bad_login')} ({max(attempts['n'], 0)})")
        if attempts["n"] <= 0:
            await app.exit()

    p.on_submit = check
    body = [card([
        label(tr("username"), 14, True), u,
        label(tr("password"), 14, True), p,
        switch_row(tr("show_pass"), False, toggle),
        button(tr("login"), check, height=46),
    ], spacing=10)]
    return screen(app, tr("login"), body)


# ---------- TEMPORARY menu (Part 3 replaces it) ----------
@view("menu")
def v_menu_temp(app):
    return screen(app, tr("main_menu"), [label("Part 3 coming...", 14, False, "sub_text")])


# ---------- App entry ----------
def main(page: ft.Page):
    load_settings()
    page.title = "Payroll Manager"
    page.padding = 0
    app = App(page)
    if not has_creds():
        app.show("setup")
    elif not STATE["require_password"]:
        app.show("menu")
    else:
        app.show("login")



import asyncio

SNOOZE_HOURS = 3


def maybe_ask(app):
    now = datetime.datetime.now()
    if not app.logged_in or app.current != "menu":
        return
    if now.date() < PROMPT_FROM or not prompt_is_due():
        return
    last = getattr(app, "last_prompt", None)
    if last and now - last < datetime.timedelta(hours=SNOOZE_HOURS):
        return
    app.last_prompt = now
    app.show("prompt")


async def ask_soon(app):
    await asyncio.sleep(0.4)
    maybe_ask(app)


async def periodic(app):
    while True:
        await asyncio.sleep(30 * 60)
        maybe_ask(app)


def shift_buttons(app, on_done=None, height=48):
    def pick(val, column):
        def handler(e):
            ok, msg = record_shift(val, column)
            app.toast(msg)
            if ok and on_done:
                on_done()
        return handler

    return [
        button(tr("morning"), pick("1st", 4), bg=C_BLUE, icon=ft.Icons.WB_SUNNY, height=height, size=15),
        button(tr("evening"), pick("2nd", 4), bg=C_PURPLE, icon=ft.Icons.NIGHTLIGHT_ROUND, height=height, size=15),
        button(tr("night"), pick("3rd", 4), bg=C_SLATE, icon=ft.Icons.DARK_MODE, height=height, size=15),
        button(tr("vacation"), pick("vacation", 5), bg=C_AMBER, icon=ft.Icons.FREE_BREAKFAST, height=height, size=15),
    ]


@view("menu")
def v_menu(app):
    t = theme()
    app.logged_in = True
    app.page.run_task(ask_soon, app)
    if not getattr(app, "loop_started", False):
        app.loop_started = True
        app.page.run_task(periodic, app)

    settings_btn = ft.Container(
        content=button(tr("settings"), lambda e: app.show("settings"),
                       bg=t["btn_bg"], fg=t["btn_fg"], height=36, radius=12,
                       icon=ft.Icons.SETTINGS, size=12),
        width=120)
    top = ft.Row([settings_btn],
                 alignment=ft.MainAxisAlignment.START if is_rtl() else ft.MainAxisAlignment.END)
    body = [
        top,
        label(date_label(), 13, False, "sub_text"),
        button(tr("record_btn"), lambda e: app.show("record"), bg=C_BLUE,
               height=64, radius=18, icon=ft.Icons.CALENDAR_MONTH, size=16),
        button(tr("payroll_btn"), lambda e: app.show("salary"), bg=C_GREEN,
               height=64, radius=18, icon=ft.Icons.PAYMENTS, size=16),
        button(tr("exit"), app.exit, bg=C_RED, height=48, radius=18),
    ]
    return screen(app, tr("main_menu"), body)


@view("record")
def v_record(app):
    body = [card(
        [label(tr("select_shift"), 15, True)] + shift_buttons(app), spacing=10)]
    return screen(app, tr("record_title"), body, back="menu", exit_btn=True,
                  subtitle=date_label())


@view("prompt")
def v_prompt(app):
    body = [card(shift_buttons(app, on_done=lambda: app.show("menu"), height=52), spacing=10)]
    return screen(app, tr("prompt_title"), body, back="menu", back_key="later",
                  subtitle=date_label())


import io

TEXTS["en"].update({
    "imported": "Excel file imported!",
    "invalid_xlsx": "Invalid file, or sheet name not found.",
    "exported": "Saved to:", "cancelled": "Cancelled.",
    "file_ok": "Excel file: loaded", "file_none": "Excel file: not imported yet",
    "creds_updated": "Credentials updated!", "bad_number": "Invalid number",
})
TEXTS["ar"].update({
    "imported": "تم استيراد ملف الإكسل!",
    "invalid_xlsx": "الملف غلط، أو اسم الشيت مش موجود.",
    "exported": "اتحفظ في:", "cancelled": "اتلغى.",
    "file_ok": "ملف الإكسل: موجود", "file_none": "ملف الإكسل: لسه ماتستوردش",
    "creds_updated": "تم تحديث البيانات!", "bad_number": "رقم غلط",
})

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩٫", "0123456789.")


def norm(txt):
    return (txt or "").translate(AR_DIGITS).strip()


SALARY_KEYS = ["start", "end", "rate2", "rate3", "base", "quality", "seniority"]
RATE_NAMES = {"rate2": "r2", "rate3": "r3", "base": "base",
              "quality": "quality", "seniority": "seniority"}


def parse_rates(vals):
    nums = {}
    for k, name in RATE_NAMES.items():
        txt = vals[k].replace(",", "")
        try:
            nums[name] = float(txt) if txt else 0.0
        except ValueError:
            return None, k
    return nums, None


@view("salary")
def v_salary(app):
    saved = getattr(app, "salary_vals", {})
    fields = {}
    rows = []
    for k in SALARY_KEYS:
        kb = None if k in ("start", "end") else ft.KeyboardType.NUMBER
        f = field(value=saved.get(k, ""), keyboard=kb)
        f.dense = True
        f.text_align = ft.TextAlign.CENTER
        fields[k] = f
        items = [ft.Container(label(tr(k), 14), expand=True), ft.Container(f, width=130)]
        if is_rtl():
            items.reverse()
        rows.append(ft.Row(items, alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                           vertical_alignment=ft.CrossAxisAlignment.CENTER))

    def calculate(e):
        vals = {k: norm(fields[k].value) for k in SALARY_KEYS}
        app.salary_vals = vals
        nums, bad = parse_rates(vals)
        if bad:
            app.toast(f"{tr('bad_number')}: {tr(bad)}")
            return
        data, err = calc_payroll(vals["start"], vals["end"], nums)
        if err:
            app.toast(err)
            return
        app.report = data
        app.show("report")

    fields["seniority"].on_submit = calculate
    body = [card(rows, spacing=8), button(tr("calc"), calculate, height=48)]
    return screen(app, tr("payroll_title"), body, back="menu", exit_btn=True)


T_BG = "#0A100C"
T_GREEN = "#4ADE80"
T_DIM = "#6B8F77"
T_GOLD = "#FACC15"
T_LINE = "#1E3326"
T_BOX = "#10261A"


def rep_row(key, value, color=T_GREEN, big=False):
    size = 17 if big else 14
    k = ft.Text(tr(key), size=size, color=color if big else T_DIM)
    v = ft.Text(value, size=size, color=color, weight=ft.FontWeight.BOLD,
                font_family="monospace")
    items = [v, k] if is_rtl() else [k, v]
    return ft.Row(items, alignment=ft.MainAxisAlignment.SPACE_BETWEEN)


def rep_sep():
    return ft.Container(height=1, bgcolor=T_LINE)


@view("report")
def v_report(app):
    d = app.report
    if not d:
        return screen(app, tr("report_title"), [], back="salary")
    net_box = ft.Container(
        content=rep_row("net_pay", f'{d["net"]:.2f}', T_GOLD, True),
        bgcolor=T_BOX, border=ft.Border.all(1, T_GREEN),
        border_radius=8, padding=ft.Padding.all(10))
    lines = [
        ft.Text("$ payroll --report", size=14, color=T_GREEN,
                weight=ft.FontWeight.BOLD, font_family="monospace"),
        rep_sep(),
        rep_row("from_date", str(d["start"])),
        rep_row("to_date", str(d["end"])),
        rep_sep(),
        rep_row("total_days", str(d["total"])),
        rep_row("n_morning", str(d["m1"])),
        rep_row("n_evening", f'{d["m2"]}  (+{d["b2"]:.2f})'),
        rep_row("n_night", f'{d["m3"]}  (+{d["b3"]:.2f})'),
        rep_row("n_vac", str(d["vac"])),
        rep_sep(),
        rep_row("base", f'{d["base"]:.2f}'),
        rep_row("quality", f'{d["quality"]:.2f}'),
        rep_row("seniority", f'{d["seniority"]:.2f}'),
        rep_row("shift_bonus", f'{d["b2"] + d["b3"]:.2f}'),
        rep_sep(),
        net_box,
        ft.Text("[ok] report generated", size=12, color=T_DIM, font_family="monospace"),
    ]
    term = ft.Container(
        content=ft.Column(lines, spacing=7,
                          horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
        bgcolor=T_BG, border=ft.Border.all(1, T_LINE),
        border_radius=12, padding=ft.Padding.all(12))
    return screen(app, tr("report_title"), [term], back="salary", exit_btn=True)


def validate_xlsx_bytes(data):
    try:
        wb = load_workbook(io.BytesIO(data))
        wb[SHEET_NAME]
        return True
    except Exception:
        return False


def get_picker(app):
    p = getattr(app, "picker", None)
    if p is None:
        p = ft.FilePicker()
        app.page.services.append(p)
        app.picker = p
    return p


@view("settings")
def v_settings(app):
    t = theme()

    def toggle_lock(e):
        STATE["require_password"] = bool(e.control.value)
        save_settings()

    def set_theme(name):
        def h(e):
            STATE["theme"] = name
            save_settings()
            app.show("settings")
        return h

    def set_lang(code):
        def h(e):
            STATE["lang"] = code
            save_settings()
            app.show("settings")
        return h

    def seg(text, active, handler):
        return button(text, handler,
                      bg=t["primary"] if active else t["btn_bg"],
                      fg="#FFFFFF" if active else t["btn_fg"],
                      height=40, radius=12, size=12, expand=True)

    theme_row = ft.Row([
        seg("Light", STATE["theme"] == "light", set_theme("light")),
        seg("Dark", STATE["theme"] == "dark", set_theme("dark")),
        seg("Liquid Glass", STATE["theme"] == "liquid_glass", set_theme("liquid_glass")),
    ], spacing=6)
    lang_row = ft.Row([
        seg("English", STATE["lang"] == "en", set_lang("en")),
        seg("العربية", STATE["lang"] == "ar", set_lang("ar")),
    ], spacing=6)

    nu = field()
    npw = field(password=True)

    def update_creds(e):
        u, p = nu.value.strip(), npw.value.strip()
        if not u or not p:
            app.toast(tr("empty_fields"))
            return
        write_creds(u, p)
        nu.value = ""
        npw.value = ""
        nu.update()
        npw.update()
        app.toast(tr("creds_updated"))

    async def do_import(e):
        files = await get_picker(app).pick_files(
            allow_multiple=False, with_data=True,
            file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["xlsx"])
        if not files:
            return
        data = files[0].bytes
        if not data and files[0].path:
            with open(files[0].path, "rb") as f:
                data = f.read()
        if not data or not validate_xlsx_bytes(data):
            app.toast(tr("invalid_xlsx"))
            return
        with open(XLSX_PATH, "wb") as f:
            f.write(data)
        app.toast(tr("imported"))
        app.show("settings")

    async def do_export(e):
        if not os.path.exists(XLSX_PATH):
            app.toast(tr("no_file"))
            return
        with open(XLSX_PATH, "rb") as f:
            data = f.read()
        path = await get_picker(app).save_file(
            file_name="shifts_2027.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["xlsx"],
            src_bytes=data)
        app.toast(f"{tr('exported')} {path}" if path else tr("cancelled"))

    status = tr("file_ok") if os.path.exists(XLSX_PATH) else tr("file_none")
    body = [
        card([
            switch_row(tr("pass_lock"), STATE["require_password"], toggle_lock),
            label(tr("theme_mode"), 14, True), theme_row,
            label(tr("language"), 14, True), lang_row,
        ], spacing=10),
        card([
            label(tr("change_creds"), 14, True),
            label(tr("new_user"), 13, False, "sub_text"), nu,
            label(tr("new_pass"), 13, False, "sub_text"), npw,
            button(tr("update_creds"), update_creds, height=44),
        ], spacing=8),
        card([
            label(status, 13, False, "sub_text"),
            button(tr("import_xlsx"), do_import, bg=C_GREEN, icon=ft.Icons.UPLOAD_FILE, height=46),
            button(tr("export_xlsx"), do_export, bg=C_BLUE, icon=ft.Icons.DOWNLOAD, height=46),
        ], spacing=8),
    ]
    return screen(app, tr("settings"), body, back="menu")


ft.run(main)
