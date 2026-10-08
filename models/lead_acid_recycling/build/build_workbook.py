# -*- coding: utf-8 -*-
"""ساختِ کاملِ فایل مدل از روی کد (بدون وابستگی به فایلِ موجود).

اجرا:  python build/build_workbook.py [مسیر خروجی] [--only=input|engine|output]
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook  # noqa: E402

import common  # noqa: E402
import inputs  # noqa: E402
import spec  # noqa: E402


def build(path=None, only=None):
    """ساختِ کتاب کار. only می‌تواند 'input'، 'engine' یا 'output' باشد."""
    path = path or os.path.join(spec.OUT_DIR, "built.xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)

    b = common.Book(start_year=inputs.START_YEAR)
    inputs.build(b)

    # لایه‌های موتور و خروجی به‌تدریج افزوده می‌شوند؛ اگر ماژولی هنوز نوشته
    # نشده باشد، ساخت ادامه می‌یابد تا بتوان مرحله‌به‌مرحله مقایسه کرد.
    done = []
    if only != "input":
        for mod in ("engine_a", "engine_b", "outputs"):
            try:
                m = __import__(mod)
            except ImportError:
                continue
            m.build(b)
            done.append(mod)
    if "outputs" not in done:
        print("  (هنوز ساخته‌نشده: %s)" % "، ".join(
            x for x in ("engine_a", "engine_b", "outputs") if x not in done))

    # پنهان کردنِ شیت‌های موتور (طبق نقشه: لایهٔ موتور پنهان است)
    for name in spec.ENGINE_SHEETS:
        if name in b.wb.sheetnames:
            b.wb[name].sheet_state = "hidden"

    complete = all(x in done for x in ("engine_a", "engine_b", "outputs"))
    left = b.save(path, strict=complete)
    if left:
        print("  ⚠ %d ارجاعِ حل‌نشده (شیت‌های ساخته‌نشده)" % left)
    wb = load_workbook(path)
    nform = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith("="))
    print("✓ ساخته شد: %s" % path)
    print("  شیت‌ها: %d | سلول‌های فرمولی: %d | کلیدها: %d"
          % (len(wb.sheetnames), nform, len(b.keys)))
    return path


if __name__ == "__main__":
    out = None
    only = None
    for arg in sys.argv[1:]:
        if arg.startswith("--only="):
            only = arg.split("=", 1)[1]
        elif not arg.startswith("--"):
            out = arg
    build(out, only)
