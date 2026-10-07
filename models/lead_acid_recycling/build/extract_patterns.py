# -*- coding: utf-8 -*-
"""استخراجِ الگوی نمادینِ فرمول‌ها (ابزارِ توسعهٔ مولد).

فرمولِ ستونِ D هر سطر را به الگویی تبدیل می‌کند که در آن:
  · ارجاع به سطری که کلید دارد   →  {شیت.کلید[t]}   (t = شمارهٔ دوره)
  · ارجاع به سطرِ بدون کلید       →  {شیت!r<سطر>c<ستون>[t]}
و سپس بررسی می‌کند آیا همهٔ دوره‌ها از همان الگو پیروی می‌کنند (با جایگزینیِ
حرفِ ستون). خروجی فشرده است تا بتوان آن را مستقیماً در کدِ مولد بازنویسی کرد.

اجرا:  python build/extract_patterns.py ENG_CAPEX [سطر_آغاز] [سطر_پایان]
"""

import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import spec

REF = re.compile(r"'([A-Z0-9_]+)'!\$([A-Z]{1,3})\$?(\d+)")
WB = load_workbook(spec.MODEL_PATH)


def main(sheet, r0=1, r1=None):
    wb = load_workbook(spec.MODEL_PATH)
    ws = wb[sheet]
    keys = spec.discover_keys(wb)
    rows = {}
    for (sh, k), r in keys.items():
        if sh == sheet:
            rows.setdefault(r, k)
    r1 = r1 or ws.max_row

    for r in range(r0, r1 + 1):
        key = ws.cell(r, 2).value
        lab = ws.cell(r, 1).value
        d = ws.cell(r, spec.C0).value
        if d is None:
            if lab or key:
                print("%-4d %-18s %-38s ·" % (r, str(key or "")[:18], str(lab or "")[:38]))
            continue
        if not (isinstance(d, str) and d.startswith("=")):
            print("%-4d %-18s %-38s = %s" % (r, str(key or "")[:18], str(lab or "")[:38], d))
            continue
        pat = pattern(d, sheet, rows)
        print("%-4d %-18s %-34s %s" % (r, str(key or "")[:18], str(lab or "")[:34], pat))


def shift_col(f, t):
    """الگوی یک فرمولِ دوره‌ای با جابه‌جاییِ ستون به t"""
    tgt = get_column_letter(spec.C0 + t)
    return re.sub(r"(\$?)([A-Z]{1,3})(\$?\d+)",
                  lambda m: m.group(1) + tgt + m.group(3)
                  if m.group(2) != tgt else m.group(0), f)


REV = {}
for (_sh, _k), _r in spec.discover_keys(WB).items():
    REV.setdefault((_sh, _r), _k)


def pattern(f, sheet=None, rows=None):
    """جایگزینی ارجاع‌ها با نامِ نمادین و شمارهٔ دوره"""
    def sub(m):
        sh, col, row = m.group(1), m.group(2), int(m.group(3))
        t = ord(col[-1]) - ord(get_column_letter(spec.C0)[-1]) if len(col) == 1 else None
        k = REV.get((sh, row))
        tag = ("%s.%s" % (sh, k)) if k else ("%s!r%d" % (sh, row))
        return "{%s[%s]}" % (tag, t if t is not None else col)
    return REF.sub(sub, f)


if __name__ == "__main__":
    sh = sys.argv[1]
    a = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    bb = int(sys.argv[3]) if len(sys.argv) > 3 else None
    main(sh, a, bb)
