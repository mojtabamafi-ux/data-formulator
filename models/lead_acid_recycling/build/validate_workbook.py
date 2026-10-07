# -*- coding: utf-8 -*-
"""اعتبارسنجِ ایستای مدل: ارجاع‌های نامعتبر، یک‌طرفه بودنِ جریان داده،
نبودِ دور ارجاعی و نبودِ فرمول/عددِ سخت‌کدشده در لایهٔ ورودی.

اجرا:  python build/validate_workbook.py [مسیر فایل]
خروجِ برنامه: ۰ = سالم، ۱ = وجودِ خطا
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook

from mini_excel import tokenize, XLError, Evaluator  # noqa: E402
import spec  # noqa: E402

TOL = 1e-6


def dynamic_cycles(path):
    """بررسیِ پویای دورِ ارجاعی: ارزیابیِ واقعیِ همهٔ سلول‌های فرمولی.

    برخلافِ تحلیلِ ایستا، این بررسی بازه‌های پویا (مانند INDEX(...):INDEX(...))
    را هم درست می‌سنجد؛ هر سلولی که ارزیابی‌اش با «دور ارجاعی» روبه‌رو شود
    گزارش می‌شود. خروجی: فهرستِ (آدرس، پیام)
    """
    ev = Evaluator(path)
    wb = ev.wb
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                try:
                    ev.get(ws.title, c.row, c.column)
                except XLError as e:
                    if "دور" in str(e):
                        bad.append(("%s!%s" % (ws.title, c.coordinate), str(e)))
    return bad, ev


def col_index(txt):
    n = 0
    for ch in txt.replace("$", "").upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _ref_parts(m, sheet, inherit=None):
    """(شیت، سطر، ستون) یک ارجاع؛ پیشوندِ شیت اگر نباشد به ارث می‌رسد"""
    sh = m.group("sheet1") or m.group("sheet2") or inherit or sheet
    return sh, int(m.group("row").replace("$", "")), col_index(m.group("col"))


def extract_refs(text, sheet):
    """همهٔ ارجاع‌های یک فرمول به‌صورت (شیت، سطر، ستون).

    نکته: در بازه‌ای مانند 'ENG_FCF'!$D$40:$P$40 بخشِ دوم پیشوند ندارد،
    بنابراین شیت از ارجاعِ نخست به ارث می‌رسد (همان رفتارِ ارزیاب).
    """
    toks = tokenize(text)
    out, i = [], 0
    while i < len(toks):
        kind, m = toks[i]
        if kind != "ref":
            i += 1
            continue
        sh, rr, rc = _ref_parts(m, sheet)
        out.append((sh, rr, rc))
        if i + 2 < len(toks) and toks[i + 1][0] == "colon" and toks[i + 2][0] == "ref":
            out.append(_ref_parts(toks[i + 2][1], sheet, inherit=sh))
            i += 2
        i += 1
    return out


def extract_ranges(text, sheet):
    """استخراجِ بازه‌ها برای بررسیِ پوششِ محدوده (دو ارجاعِ پشت‌سرهم با دونقطه)"""
    toks = tokenize(text)
    out, i = [], 0
    while i < len(toks):
        kind, m = toks[i]
        if kind == "ref":
            sh, r1, c1 = _ref_parts(m, sheet)
            r2, c2 = r1, c1
            if i + 2 < len(toks) and toks[i + 1][0] == "colon" and toks[i + 2][0] == "ref":
                _, r2, c2 = _ref_parts(toks[i + 2][1], sheet, inherit=sh)
                i += 2
            out.append((sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
        i += 1
    return out


def main(path=None):
    path = path or spec.MODEL_PATH
    wb = load_workbook(path, data_only=False)
    sheets = set(wb.sheetnames)

    n_form = 0
    invalid = []        # ارجاع‌های نامعتبر
    flow = []           # خطاهای جریانِ داده
    input_formulas = []
    hardcoded = []      # عددِ سخت‌کدشده در موتور/خروجی
    graph = {}          # گرافِ وابستگی برای کشفِ دور

    for ws in wb.worksheets:
        sh = ws.title
        layer = spec.LAYER.get(sh, 2)
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None or not isinstance(v, str) or not v.startswith("="):
                    if (layer >= 1 and c.column >= spec.C0 and isinstance(v, (int, float))
                            and not isinstance(v, bool)):
                        hardcoded.append("%s!%s" % (sh, c.coordinate))
                    continue
                n_form += 1
                if layer == 0:
                    input_formulas.append("%s!%s" % (sh, c.coordinate))
                try:
                    refs = extract_refs(v[1:], sh)
                    rngs = extract_ranges(v[1:], sh)
                except XLError as e:
                    invalid.append("%s!%s → %s" % (sh, c.coordinate, e.code))
                    continue

                deps = []
                for (rs, rr, rc) in refs:
                    if rs not in sheets:
                        invalid.append("%s!%s → شیت ناشناخته %s" % (sh, c.coordinate, rs))
                        continue
                    wsx = wb[rs]
                    if rr > wsx.max_row or rc > wsx.max_column:
                        invalid.append("%s!%s → %s!R%dC%d خارج از محدوده (حداکثر R%dC%d)"
                                       % (sh, c.coordinate, rs, rr, rc,
                                          wsx.max_row, wsx.max_column))
                        continue
                    tl = spec.LAYER.get(rs, 2)
                    if tl > layer:
                        flow.append("%s!%s → %s (از لایهٔ %d به %d)"
                                    % (sh, c.coordinate, rs, layer, tl))
                    deps.append((rs, rr, rc))
                # گرافِ ایستا فقط از ارجاع‌های تک‌سلولی ساخته می‌شود؛ بازه‌ها
                # گسترش نمی‌یابند، زیرا بازه‌های پویا (INDEX(...):INDEX(...)) که
                # مبدأ را هم در بر می‌گیرند، دورِ کاذب می‌سازند. پوششِ واقعیِ
                # این حالت‌ها در بررسیِ پویا با ارزیاب انجام می‌شود.
                graph[(sh, c.row, c.column)] = deps
                for (rs, r1, c1, r2, c2) in rngs:
                    if rs in sheets and (r2 > wb[rs].max_row or c2 > wb[rs].max_column):
                        invalid.append("%s!%s → %s!R%dC%d:R%dC%d خارج از محدوده"
                                       % (sh, c.coordinate, rs, r1, c1, r2, c2))

    # ------------------------------------------------------- کشفِ دور ارجاعی
    cycles = []
    WHITE, GREY, BLACK = 0, 1, 2
    color = {}

    def dfs(node, stack):
        color[node] = GREY
        stack.append(node)
        for dep in graph.get(node, ()):
            # فقط یال‌هایی که به سلولِ فرمولی می‌روند می‌توانند دور بسازند
            if dep not in graph:
                continue
            st = color.get(dep, WHITE)
            if st == GREY:
                i = stack.index(dep)
                cycles.append(" ← ".join("%s!%s%d" % (s, chr(64 + cc) if cc < 27 else "?", rr)
                                         for (s, rr, cc) in stack[i:] + [dep]))
            elif st == WHITE:
                dfs(dep, stack)
        stack.pop()
        color[node] = BLACK

    sys.setrecursionlimit(100000)
    for node in list(graph.keys()):
        if color.get(node, WHITE) == WHITE:
            dfs(node, [])

    # ------------------------------------------- کشفِ دور ارجاعی به‌صورتِ پویا
    dyn, ev = dynamic_cycles(path)

    # ----------------------------------------------------------------- گزارش
    print("=" * 72)
    print("اعتبارسنجی ایستای مدل")
    print("=" * 72)
    print("  فایل: %s" % os.path.relpath(path, spec.ROOT))
    print("  تعداد شیت‌ها: %d | تعداد سلول‌های فرمولی: %d" % (len(wb.sheetnames), n_form))
    print("  ارجاع‌های نامعتبر: %d" % len(invalid))
    for x in invalid[:10]:
        print("      ⚠ %s" % x)
    print("  خطاهای جریان داده: %d" % len(flow))
    for x in flow[:10]:
        print("      ⚠ %s" % x)
    print("  فرمول در شیت‌های ورودی: %d" % len(input_formulas))
    for x in input_formulas[:10]:
        print("      ⚠ %s" % x)
    print("  دورهای ارجاعی (تحلیل ایستا): %d" % len(cycles))
    for x in cycles[:10]:
        print("      ⚠ %s" % x)
    print("  دورهای ارجاعی (ارزیابی پویا): %d" % len(dyn))
    for x in dyn[:10]:
        print("      ⚠ %s → %s" % x)
    print("  بیشینهٔ عمقِ زنجیرهٔ ارجاع: %d" % ev.max_depth)
    print("  عدد سخت‌کدشده در لایهٔ موتور/خروجی (اطلاعاتی): %d" % len(hardcoded))
    for x in hardcoded[:10]:
        print("      · %s" % x)
    ok = not invalid and not flow and not input_formulas and not cycles and not dyn
    print("-" * 72)
    print("نتیجه: %s" % ("مدل سالم است ✅" if ok else "مدل نیازمند اصلاح است ❌"))
    print("=" * 72)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else None))
