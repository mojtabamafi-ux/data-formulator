# -*- coding: utf-8 -*-
"""تعریف‌های ثابتِ مدل مالی و عملیاتی کارخانهٔ بازیافت باتری سرب–اسیدی.

همهٔ گزارش‌ها و کنترل‌ها به‌جای آدرسِ عددیِ ثابت، با «کلیدِ» درج‌شده در ستونِ B
هر سطر آدرس‌دهی می‌شوند؛ پس اگر سطرها جابه‌جا شوند، ابزارها همچنان درست کار می‌کنند.
"""

import os

# ------------------------------------------------------------------ مسیرها
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL_PATH = os.path.join(ROOT, "Battery_Recycling_Financial_Model.xlsx")
OUT_DIR = os.path.join(ROOT, "build", "out")

# ------------------------------------------------------- ابعاد و لایه‌بندی
C0 = 4          # نخستین ستونِ داده (D)
NY = 13         # تعداد دوره‌ها: ۰ (ساخت) تا ۱۲
NS = 6          # تعداد ایستگاه‌ها
NP = 4          # تعداد محصولات
NCAP = 11       # طبقاتِ سرمایه‌گذاری
NSCEN = 3       # سناریوها

INPUT_SHEETS = ["IN_MACRO", "IN_TECH", "IN_SALES", "IN_COST", "IN_WC", "IN_FIN", "IN_SCEN"]
ENGINE_SHEETS = ["ENG_SCEN", "ENG_OPS", "ENG_ALLOC", "ENG_COST", "ENG_WC",
                 "ENG_CAPEX", "ENG_FCF", "ENG_DCF", "ENG_LEV", "ENG_SENS"]
OUTPUT_SHEETS = ["OUT_DASH", "OUT_BANK", "OUT_DEBT", "OUT_SENS", "OUT_INVEST"]
SERVICE_SHEETS = ["00_MAP", "99_CHECK"]

SHEET_ORDER = (SERVICE_SHEETS[:1] + INPUT_SHEETS + ENGINE_SHEETS +
               OUTPUT_SHEETS + SERVICE_SHEETS[1:])

LAYER = {}
for _s in INPUT_SHEETS:
    LAYER[_s] = 0
for _s in ENGINE_SHEETS:
    LAYER[_s] = 1
for _s in OUTPUT_SHEETS + SERVICE_SHEETS:
    LAYER[_s] = 2

SCENARIOS = ["پایه", "خوش\u200cبینانه", "بدبینانه"]

# حدِ قراردادیِ DSCR برای کنترلِ شمارهٔ ۸ در شیت 99_CHECK (مفروضِ بانکیِ متعارف)
DSCR_MIN_CONTRACT = 1.2

# ستونِ NPV در موتورِ حساسیت: C0 + 3 (ستون‌های کلید/محورها) + NY (FCF) + NY (ضریب تنزیل)
SENS_NPV_COL = C0 + 3 + 2 * NY


# --------------------------------------------------- کشفِ کلیدهای ستونِ B
def discover_keys(wb):
    """{(نام شیت, کلید): شماره سطر} بر پایهٔ کدهای ستونِ B (و ستونِ A برای ENG_SENS)."""
    keys = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            v = row[1].value
            if isinstance(v, str) and v.strip() and v.strip() != "موتور محاسبات":
                keys[(ws.title, v.strip())] = row[0].row
            if ws.title == "ENG_SENS":
                a = row[0].value
                if isinstance(a, str) and "|" in a:
                    keys[(ws.title, a.strip())] = row[0].row
    return keys


# ------------------------------------------------------------- تعریف گزارش
# قالبِ هر آیتم:
#   ("key", شیت, کلید, شماره دوره, قالب)
#   ("row", شیت, شماره سطر, قالب)              → ستونِ D
#   ("sens", کلیدِ سطرِ ماتریس, قالب)          → ENG_SENS
#   ("ratio", آیتمِ صورت, آیتمِ مخرج, قالب)
# قالب‌ها: num | num1 | num2 | pct | scen

REPORT = [
    ("عملیات", [
        ("key", "ENG_OPS", "OPS.intake", 5, "num"),
        ("key", "ENG_OPS", "OPS.prod_total", 5, "num"),
        ("key", "ENG_OPS", "OPS.sold_tons", 5, "num"),
        ("key", "ENG_OPS", "OPS.caputil", 5, "pct"),
        ("key", "ENG_OPS", "OPS.bottleneck", 5, "pct"),
        ("key", "ENG_OPS", "OPS.kwh", 5, "num"),
        ("key", "ENG_OPS", "OPS.price", 5, "num"),
    ]),
    ("درآمد و هزینه", [
        ("key", "ENG_OPS", "OPS.rev", 5, "num"),
        ("key", "ENG_COST", "EC.dpc", 5, "num"),
        ("key", "ENG_COST", "EC.opc", 5, "num"),
        ("key", "ENG_COST", "EC.fpc", 5, "num"),
        ("key", "ENG_COST", "EC.rm", 5, "num"),
        ("key", "ENG_COST", "EC.cost_per_ton", 5, "num1"),
        ("key", "ENG_COST", "EC.ebitda", 5, "num"),
        ("key", "ENG_COST", "EC.ebit", 5, "num"),
        ("key", "ENG_FCF", "FCF.ni", 5, "num"),
        ("ratio", ("key", "ENG_COST", "EC.ebitda", 5, "num"),
         ("key", "ENG_OPS", "OPS.rev", 5, "num"), "pct"),
    ]),
    ("دارایی‌ها و استهلاک", [
        ("key", "ENG_CAPEX", "CAPEX.cap_total", 0, "num"),
        ("key", "ENG_CAPEX", "CAPEX.cap_total", 1, "num"),
        ("key", "ENG_CAPEX", "CAPEX.dep_total", 5, "num"),
        ("key", "ENG_CAPEX", "CAPEX.nbv", 12, "num"),
        ("key", "ENG_CAPEX", "CAPEX.repl_total", 8, "num"),
        ("key", "ENG_CAPEX", "CAPEX.idc", 1, "num"),
    ]),
    ("جریان نقد و تأمین مالی", [
        ("key", "ENG_FCF", "FCF.fcff", 5, "num"),
        ("key", "ENG_FCF", "FCF.fcff", 0, "num"),
        ("key", "ENG_FCF", "FCF.debt_close", 1, "num"),
        ("key", "ENG_FCF", "FCF.principal", 5, "num"),
        ("key", "ENG_FCF", "FCF.int_lt", 5, "num"),
        ("key", "ENG_FCF", "FCF.dscr", 5, "num2"),
        ("key", "ENG_FCF", "FCF.dscr", 3, "num2"),
        ("key", "ENG_FCF", "FCF.cash", 5, "num"),
        ("key", "ENG_FCF", "FCF.bs_check", 5, "num1"),
        ("key", "ENG_FCF", "FCF.bs_check", 12, "num1"),
        ("key", "ENG_WC", "WC.ccc", 5, "num1"),
        ("key", "ENG_WC", "WC.dwc", 5, "num"),
    ]),
    ("ارزش‌گذاری", [
        ("key", "ENG_DCF", "DCF.wacc", 5, "pct"),
        ("row", "ENG_DCF", 26, "num"),   # NPV
        ("row", "ENG_DCF", 27, "pct"),   # IRR
        ("row", "ENG_DCF", 28, "pct"),   # MIRR
        ("row", "ENG_DCF", 29, "pct"),   # IRR سهامداران
        ("row", "ENG_DCF", 30, "num1"),  # بازپرداخت ساده
        ("row", "ENG_DCF", 31, "num1"),  # بازپرداخت تنزیلی
        ("row", "ENG_DCF", 32, "num2"),  # شاخص سودآوری
        ("key", "ENG_DCF", "DCF.tv", 12, "num"),
        ("row", "ENG_DCF", 33, "num"),   # ارزش بنگاه
    ]),
    ("اهرم", [
        ("key", "ENG_LEV", "LEV.dol", 5, "num2"),
        ("key", "ENG_LEV", "LEV.dfl", 5, "num2"),
        ("key", "ENG_LEV", "LEV.dtl", 5, "num2"),
        ("key", "ENG_LEV", "LEV.be_input", 5, "num"),
        ("key", "ENG_LEV", "LEV.roe", 5, "pct"),
        ("key", "ENG_LEV", "LEV.roic", 5, "pct"),
    ]),
    ("سناریوها", [
        ("row", "ENG_SCEN", 68, "num"),    # NPV پایه (مدل سایه)
        ("row", "ENG_SCEN", 102, "num"),   # خوش‌بینانه
        ("row", "ENG_SCEN", 136, "num"),   # بدبینانه
        ("row", "ENG_SCEN", 69, "pct"),    # IRR پایه
        ("row", "ENG_SCEN", 103, "pct"),
        ("row", "ENG_SCEN", 137, "pct"),
        ("row", "ENG_SCEN", 72, "num2"),   # حداقل DSCR پایه
        ("row", "ENG_SCEN", 140, "num2"),  # حداقل DSCR بدبینانه
    ]),
    ("حساسیت", [
        ("sens", "FX|2|2", "num"),   # مبنا (بدون انحراف)
        ("sens", "FX|2|0", "num"),   # ارز −۲۰٪
        ("sens", "FX|2|4", "num"),   # ارز +۲۰٪
        ("sens", "FX|4|2", "num"),   # نرخ بهره +۳٪
        ("sens", "PR|2|0", "num"),   # قیمت −۱۵٪
        ("sens", "PR|2|4", "num"),   # قیمت +۱۵٪
    ]),
    ("خروجی‌ها", [
        ("row", "IN_SCEN", 8, "scen"),   # سناریوی فعال
        ("row", "OUT_DASH", 11, "num"),  # NPV روی داشبورد
        ("row", "OUT_DASH", 12, "pct"),  # IRR روی داشبورد
    ]),
]

REPORT_LABELS = {
    "OPS.intake": "ورودی باتری (تن)",
    "OPS.prod_total": "تولید محصولات (تن)",
    "OPS.sold_tons": "فروش محصولات (تن)",
    "OPS.caputil": "نرخ استفاده از ظرفیت",
    "OPS.bottleneck": "بارگذاری گلوگاه",
    "OPS.kwh": "برق مصرفی (kWh)",
    "OPS.price": "قیمت سرب خالص",
    "OPS.rev": "درآمد خالص",
    "EC.dpc": "DPC",
    "EC.opc": "OPC",
    "EC.fpc": "FPC",
    "EC.rm": "مواد اولیه",
    "EC.cost_per_ton": "بهای تمام‌شده هر تن",
    "EC.ebitda": "EBITDA",
    "EC.ebit": "EBIT",
    "FCF.ni": "سود خالص",
    "CAPEX.cap_total": "جمع CAPEX",
    "CAPEX.dep_total": "استهلاک",
    "CAPEX.nbv": "ارزش دفتری",
    "CAPEX.repl_total": "جایگزینی",
    "CAPEX.idc": "IDC",
    "FCF.fcff": "FCFF",
    "FCF.debt_close": "مانده بدهی پایان سال",
    "FCF.principal": "بازپرداخت اصل",
    "FCF.int_lt": "هزینه سود",
    "FCF.dscr": "DSCR",
    "FCF.cash": "نقد پایان سال",
    "FCF.bs_check": "کنترل تراز",
    "WC.ccc": "CCC",
    "WC.dwc": "تغییرات سرمایه در گردش",
    "DCF.wacc": "WACC",
    "DCF.tv": "ارزش پایانی",
    "LEV.dol": "DOL",
    "LEV.dfl": "DFL",
    "LEV.dtl": "DTL",
    "LEV.be_input": "نقطه سر‌به‌سر (تن)",
    "LEV.roe": "ROE",
    "LEV.roic": "ROIC",
}

ROW_LABELS = {
    ("ENG_DCF", 26): "NPV",
    ("ENG_DCF", 27): "IRR",
    ("ENG_DCF", 28): "MIRR",
    ("ENG_DCF", 29): "IRR سهامداران",
    ("ENG_DCF", 30): "دوره بازپرداخت ساده",
    ("ENG_DCF", 31): "دوره بازپرداخت تنزیلی",
    ("ENG_DCF", 32): "شاخص سودآوری",
    ("ENG_DCF", 33): "ارزش بنگاه",
    ("ENG_SCEN", 68): "NPV پایه (مدل سایه)",
    ("ENG_SCEN", 102): "NPV خوش‌بینانه",
    ("ENG_SCEN", 136): "NPV بدبینانه",
    ("ENG_SCEN", 69): "IRR پایه",
    ("ENG_SCEN", 103): "IRR خوش‌بینانه",
    ("ENG_SCEN", 137): "IRR بدبینانه",
    ("ENG_SCEN", 72): "حداقل DSCR پایه",
    ("ENG_SCEN", 140): "حداقل DSCR بدبینانه",
    ("IN_SCEN", 8): "سناریوی فعال",
    ("OUT_DASH", 11): "داشبورد: NPV",
    ("OUT_DASH", 12): "داشبورد: IRR",
}

SENS_LABELS = {
    "FX|2|2": "NPV مبنا (بدون انحراف)",
    "FX|2|0": "NPV ارز −۲۰٪",
    "FX|2|4": "NPV ارز +۲۰٪",
    "FX|4|2": "NPV نرخ بهره +۳٪",
    "PR|2|0": "NPV قیمت −۱۵٪",
    "PR|2|4": "NPV قیمت +۱۵٪",
}

# ------------------------------------------------------ کنترل‌های 99_CHECK
# (شماره سطر، شرح، نوعِ قاعده، مقدارِ هدف)
#   eq  : برابری (با تلورانس) | ge : بزرگ‌تر/مساوی | le : کوچک‌تر/مساوی | gt : بزرگ‌ترِ خالص
CHECKS = [
    (6,  "جمع سهم‌های جرمی باتری برابر یک", "eq", 1.0),
    (7,  "جمع سهم‌های محصول (سرب خالص + آلیاژ) برابر یک", "eq", 1.0),
    (8,  "جمع برنامه تزریق CAPEX هر طبقه برابر یک", "eq", 1.0),
    (9,  "بیشترین قدرمطلقِ کنترل ترازنامه برابر صفر", "eq", 0.0),
    (10, "تراز جرمی: ضایعات نامنفی", "ge", 0.0),
    (11, "بارگذاری گلوگاه حداکثر ۱۰۰٪", "le", 1.0),
    (12, "مانده نقد هیچ دوره‌ای منفی نیست", "ge", 0.0),
    (13, "حداقل DSCR دوره بازپرداخت بالاتر از حد قراردادی", "ge", DSCR_MIN_CONTRACT),
    (14, "WACC بزرگ‌تر از نرخ رشد دائمی (روش گوردون)", "gt", 0.0),
    (15, "ارزش دفتری دارایی‌ها نامنفی", "ge", 0.0),
]

# ------------------------------------------------------------ تست‌های آبشاری
# هر مورد: (شرح، وصله‌ها، ...)؛ وصله‌ها پیش از سنجه‌ها اعمال و سپس برگردانده
# می‌شوند. قالبِ وصله:
#   ("key", شیت, کلید, دوره/None, مقدار)   → دورهٔ None یعنی ستونِ نخست (مقدارِ واحد)
#   ("cell", شیت, سطر, ستون, مقدار)        → برای خانه‌های جدولیِ بدون کلید
STD_PROBES = [
    ("row", "ENG_DCF", 26, "num"),            # NPV مدل اصلی
    ("row", "ENG_DCF", 27, "pct"),            # IRR مدل اصلی
    ("row", "OUT_DASH", 11, "num"),           # NPV روی داشبورد
    ("row", "OUT_DASH", 12, "pct"),           # IRR روی داشبورد
    ("key", "ENG_OPS", "OPS.intake", 5, "num"),
    ("key", "ENG_OPS", "OPS.prod_total", 5, "num"),
    ("key", "ENG_COST", "EC.ebitda", 5, "num"),
    ("key", "ENG_FCF", "FCF.fcff", 5, "num"),
    ("row", "ENG_SCEN", 68, "num"),           # NPV پایه (مدل سایه)
    ("row", "ENG_SCEN", 102, "num"),          # NPV خوش‌بینانه
    ("row", "ENG_SCEN", 136, "num"),          # NPV بدبینانه
]

CASCADE = [
    ("سناریوی فعال: خوش‌بینانه", [("row", "IN_SCEN", 8, 4, 2)]),
    ("سناریوی فعال: بدبینانه", [("row", "IN_SCEN", 8, 4, 3)]),
    ("نرخ ارزِ سال مبنا: ۵۰۰→۶۰۰ هزار", [("key", "IN_MACRO", "MAC.fx", 0, 600000)]),
    ("LME سال مبنا: ۲۱۰۰→۲۶۰۰", [("key", "IN_MACRO", "MAC.lme", 0, 2600)]),
    ("ظرفیت اسمی ورودی: ۴۰→۴۸ هزار تن", [("key", "IN_TECH", "TECH.cap_in", None, 48000)]),
    ("سهم تسهیلات: ۶۰٪→۴۵٪", [("key", "IN_FIN", "FIN.debt_share", None, 0.45)]),
    ("نرخ مالیات: ۲۰٪→۲۵٪", [("key", "IN_MACRO", "MAC.tax", None, 0.25)]),
    ("ظرفیت ایستگاه S3: ۱۷→۷ هزار تن", [("cell", "IN_TECH", 36, 6, 7000)]),
    ("ظرفیت ایستگاه S1: ۴۲→۳۰ هزار تن", [("cell", "IN_TECH", 36, 4, 30000)]),
    ("دسترسی ایستگاه S2: ۰.۹۰→۰.۷۰", [("cell", "IN_TECH", 38, 5, 0.70)]),
]
