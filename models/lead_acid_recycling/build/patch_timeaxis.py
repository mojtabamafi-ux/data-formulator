# -*- coding: utf-8 -*-
"""تک‌سلول‌سازیِ محورِ زمان: تبدیلِ سطرهای ۷ و ۸ (شمارهٔ دوره و سالِ تقویمی)
در شیت‌های موتور و خروجی از «عدد ثابت» به «ارجاع به IN_MACRO».

چرا: قاعدهٔ مدل می‌گوید در لایهٔ محاسبات/خروجی هیچ عددِ سخت‌کدشده‌ای نباشد؛
ضمن این‌که کپی شدنِ سالِ تقویمی در ۱۲ شیت، ریسکِ ناهماهنگی دارد. منبعِ معتبر
یکی است: IN_MACRO (سطرِ MAC.t و MAC.year).

اجرا:  python build/patch_timeaxis.py [--apply]
بدونِ --apply فقط بررسی می‌کند و چیزی نمی‌نویسد.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import spec

SRC = "IN_MACRO"


def main(apply=False):
    path = spec.MODEL_PATH
    wb = load_workbook(path)
    src = wb[SRC]
    keys = spec.discover_keys(wb)
    rt = keys.get((SRC, "MAC.t"))
    ry = keys.get((SRC, "MAC.year"))
    if not rt or not ry:
        print("✗ ردیف‌های MAC.t / MAC.year در IN_MACRO یافت نشد")
        return 1
    t_vals = [src.cell(rt, spec.C0 + i).value for i in range(spec.NY)]
    y_vals = [src.cell(ry, spec.C0 + i).value for i in range(spec.NY)]
    print("محورِ زمان در IN_MACRO: دوره %s … %s | سال %s … %s"
          % (t_vals[0], t_vals[-1], y_vals[0], y_vals[-1]))

    targets, mism, done, skipped = [], [], 0, 0
    for ws in wb.worksheets:
        if spec.LAYER.get(ws.title, 2) == 0:      # شیت‌های ورودی باید بی‌فرمول بمانند
            continue
        cells = []
        for i in range(spec.NY):
            col = spec.C0 + i
            ct, cy = ws.cell(7, col), ws.cell(8, col)
            if not (isinstance(ct.value, (int, float)) and isinstance(cy.value, (int, float))):
                continue
            if ct.value != t_vals[i] or cy.value != y_vals[i]:
                mism.append((ws.title, ct.coordinate, ct.value, t_vals[i]))
                continue
            cells.append((ct, cy, col))
        if not cells:
            continue
        targets.append((ws.title, len(cells)))
        if apply:
            for ct, cy, col in cells:
                L = get_column_letter(col)
                ct.value = "='%s'!%s$%d" % (SRC, L, rt)
                cy.value = "='%s'!%s$%d" % (SRC, L, ry)
                done += 1
        else:
            skipped += len(cells)

    print("شیت‌های مشمول (%d):" % len(targets))
    for sh, n in targets:
        print("   %-12s %d ستون" % (sh, n))
    if mism:
        print("✗ ناهماهنگی با IN_MACRO (%d مورد):" % len(mism))
        for m in mism[:8]:
            print("   %s!%s مقدار %s در برابر %s" % m)
        return 1
    if apply:
        wb.save(path)
        print("✓ %d سلول به ارجاع تبدیل شد و فایل ذخیره گردید" % done)
    else:
        print("· حالتِ بررسی: %d سلول بدون تغییر (برای اعمال، --apply بدهید)" % skipped)
    return 0


if __name__ == "__main__":
    sys.exit(main(apply="--apply" in sys.argv))
