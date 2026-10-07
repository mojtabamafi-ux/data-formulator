# -*- coding: utf-8 -*-
"""بازمحاسبهٔ کاملِ مدل و سنجشِ انطباق با مقدارهای مرجع.

سه مرحله:
  ۱) ارزیابیِ همهٔ سلول‌های فرمولی با mini_excel و شمارشِ خطاها؛
  ۲) خواندنِ سنجه‌های تعریف‌شده در spec.REPORT و مقایسه با مرجع (baseline.json)؛
  ۳) اجرای تست‌های آبشاری (تغییرِ یک ورودی و دیدنِ اثرِ آن تا NPV/IRR).

اجرا:
  python build/recalc_check.py            مقایسه با مرجع
  python build/recalc_check.py --record   نوشتنِ مرجعِ جدید
خروجیِ برنامه: ۰ = انطباق کامل، ۱ = انحراف
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mini_excel import Evaluator, XLError  # noqa: E402
import spec  # noqa: E402

BASELINE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "baseline.json")
REL_TOL = 1e-9
ABS_TOL = 1e-6


# ----------------------------------------------------------------- ابزارها
def fmt(v, kind="num"):
    if isinstance(v, XLError):
        return "خطا %s" % v.code
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return str(v)
    if kind == "pct":
        return "%.3f%%" % (v * 100)
    if kind == "num1":
        return "{:,.1f}".format(v)
    if kind == "num2":
        return "{:,.2f}".format(v)
    if kind == "scen":
        return "%s (%s)" % (v, spec.SCENARIOS[int(v) - 1] if 0 < v <= len(spec.SCENARIOS) else "؟")
    return "{:,.0f}".format(v)


def item_id(it):
    if it[0] == "key":
        return "key:%s|%s|%s" % (it[1], it[2], it[3])
    if it[0] == "row":
        return "row:%s|%s" % (it[1], it[2])
    if it[0] == "sens":
        return "sens:%s" % it[1]
    return "ratio:%s/%s" % (item_id(it[1]), item_id(it[2]))


def item_label(it):
    if it[0] == "key":
        return spec.REPORT_LABELS.get(it[2], it[2])
    if it[0] == "row":
        return spec.ROW_LABELS.get((it[1], it[2]), "%s ردیف %d" % (it[1], it[2]))
    if it[0] == "sens":
        return spec.SENS_LABELS.get(it[1], it[1])
    return "نسبت: %s / %s" % (item_label(it[1]), item_label(it[2]))


def item_kind(it):
    return it[-1]


def item_value(ev, keys, it):
    kind = it[0]
    if kind == "key":
        _, sh, key, t, _f = it
        return ev.get(sh, keys[(sh, key)], spec.C0 + t)
    if kind == "row":
        _, sh, row, _f = it
        return ev.get(sh, row, spec.C0)
    if kind == "sens":
        _k, key, _f = it
        return ev.get("ENG_SENS", keys[("ENG_SENS", key)], spec.SENS_NPV_COL)
    if kind == "ratio":
        _k, num, den, _f = it
        a = item_value(ev, keys, num)
        b = item_value(ev, keys, den)
        if isinstance(a, XLError):
            return a
        if not isinstance(b, (int, float)) or isinstance(b, bool) or b == 0:
            return XLError("#DIV/0!")
        return a / b
    raise ValueError(it)


def cells_of(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    yield ws.title, c.row, c.column


# ------------------------------------------------------------------ مرحلهٔ ۱
def stage_recompute(ev):
    hist = {}
    samples = {}
    n = 0
    for sh, r, c in cells_of(ev.wb):
        n += 1
        try:
            ev.get(sh, r, c)
        except XLError as e:
            code = e.code
            hist[code] = hist.get(code, 0) + 1
            samples.setdefault(code, "%s!R%dC%d" % (sh, r, c))
    return {"n": n, "hist": hist, "samples": samples, "max_depth": ev.max_depth}


# ------------------------------------------------------------------ مرحلهٔ ۲
def stage_metrics(ev, keys):
    out = {}
    for _section, items in spec.REPORT:
        for it in items:
            try:
                v = item_value(ev, keys, it)
            except XLError as e:
                v = "ERR:" + e.code
            except KeyError as e:
                v = "ERR:کلید ناموجود %s" % e
            out[item_id(it)] = v if isinstance(v, str) else float(v)
    return out


# ------------------------------------------------------------------ مرحلهٔ ۳
def stage_cascade(ev, keys):
    res = {}
    for title, patches in spec.CASCADE:
        backup = []
        try:
            for p in patches:
                if p[0] == "key":
                    _, sh, key, t, val = p
                    row, col = keys[(sh, key)], spec.C0 + (t or 0)
                elif p[0] == "row":
                    _, sh, row, col, val = p
                else:
                    _, sh, row, col, val = p
                backup.append((sh, row, col, ev.wb[sh].cell(row, col).value))
                ev.set_cell(sh, row, col, val)
            ev.cache.clear()
            prob = {}
            for it in spec.STD_PROBES:
                try:
                    v = item_value(ev, keys, it)
                except XLError as e:
                    v = "ERR:" + e.code
                prob[item_id(it)] = v if isinstance(v, str) else float(v)
            res[title] = prob
        finally:
            for sh, row, col, val in backup:
                ev.set_cell(sh, row, col, val)
            ev.cache.clear()
    return res


def stage_invariants(ev, keys):
    """هم‌ارزیِ NPV در چهار مسیر: موتورِ اصلی، داشبورد، مدلِ سایهٔ سناریو،
    و مرکزِ هر دو ماتریس حساسیت."""
    vals = {
        "ENG_DCF (مدل اصلی)": ev.get("ENG_DCF", 26, spec.C0),
        "OUT_DASH (داشبورد)": ev.get("OUT_DASH", 11, spec.C0),
        "ENG_SCEN (مدل سایه پایه)": ev.get("ENG_SCEN", 68, spec.C0),
        "ENG_SENS ماتریس نرخ×ارز": ev.get("ENG_SENS", keys[("ENG_SENS", "FX|2|2")], spec.SENS_NPV_COL),
        "ENG_SENS ماتریس نرخ×قیمت": ev.get("ENG_SENS", keys[("ENG_SENS", "PR|2|2")], spec.SENS_NPV_COL),
    }
    ref = vals["ENG_DCF (مدل اصلی)"]
    inv = {}
    for k, v in vals.items():
        inv[k] = {"value": float(v) if isinstance(v, (int, float)) and not isinstance(v, XLError) else str(v),
                  "ok": isinstance(v, (int, float)) and isinstance(ref, (int, float))
                        and abs(v - ref) <= ABS_TOL + REL_TOL * abs(ref)}
    return inv


# --------------------------------------------------------------------- مقایسه
def compare(label, base, cur, log):
    diffs = []
    for k, bv in base.items():
        cv = cur.get(k)
        if isinstance(bv, str) or isinstance(cv, str):
            if bv != cv:
                diffs.append((k, bv, cv))
            continue
        if cv is None:
            diffs.append((k, bv, "ناموجود"))
            continue
        if abs(cv - bv) > ABS_TOL + REL_TOL * abs(bv):
            diffs.append((k, bv, cv))
    for k in cur:
        if k not in base:
            diffs.append((k, "ناموجود", cur[k]))
    return diffs


def main(record=False):
    ev = Evaluator(spec.MODEL_PATH)
    keys = spec.discover_keys(ev.wb)

    print("=" * 78)
    print("بازمحاسبهٔ کاملِ مدل و سنجشِ انطباق")
    print("=" * 78)
    print("فایل: %s" % os.path.relpath(spec.MODEL_PATH, spec.ROOT))

    rep = stage_recompute(ev)
    print("\n۱) ارزیابیِ همهٔ سلول‌ها")
    print("   تعداد سلول‌های فرمولی: %d | بیشینهٔ عمق زنجیره: %d"
          % (rep["n"], rep["max_depth"]))
    if rep["hist"]:
        for code, n in sorted(rep["hist"].items(), key=lambda kv: -kv[1]):
            print("   خطا %-10s %4d مورد (نمونه: %s)" % (code, n, rep["samples"][code]))
    else:
        print("   هیچ خطایی یافت نشد")

    print("\n۲) سنجه‌های کلیدی")
    metrics = stage_metrics(ev, keys)
    for section, items in spec.REPORT:
        print("   ── %s" % section)
        for it in items:
            v = metrics[item_id(it)]
            print("      %-28s %s" % (item_label(it), fmt(v, item_kind(it))))

    print("\n۳) هم‌ارزیِ NPV در مسیرهای مختلف")
    inv = stage_invariants(ev, keys)
    for k, d in inv.items():
        print("   %-30s %18s  %s" % (k, "{:,.0f}".format(d["value"])
                                     if isinstance(d["value"], float) else d["value"],
                                     "✓" if d["ok"] else "✗"))

    print("\n۴) تست‌های آبشاری (تغییرِ یک ورودی → اثر تا خروجی)")
    casc = stage_cascade(ev, keys)
    for title, prob in casc.items():
        npv = prob.get("row:ENG_DCF|26")
        irr = prob.get("row:ENG_DCF|27")
        dash = prob.get("row:OUT_DASH|11")
        print("   ▸ %s" % title)
        print("       NPV: %s   |   IRR: %s   |   داشبورد NPV: %s"
              % (fmt(npv), fmt(irr, "pct"), fmt(dash)))
        extra = {k: v for k, v in prob.items() if k.startswith("row:ENG_SCEN")}
        if extra:
            print("       مدل سایه — " + " · ".join(
                "%s: %s" % (spec.ROW_LABELS.get(("ENG_SCEN", int(k.split("|")[1])), k),
                            fmt(v)) for k, v in sorted(extra.items())))

    new = {"metrics": metrics, "cascade": casc, "errors": rep["hist"]}
    if record:
        with open(BASELINE, "w", encoding="utf-8") as f:
            json.dump(new, f, ensure_ascii=False, indent=1, sort_keys=True)
        print("\n✓ مرجعِ جدید در %s نوشته شد" % os.path.relpath(BASELINE, spec.ROOT))
        return 0

    if not os.path.exists(BASELINE):
        print("\n! مرجعی برای مقایسه نیست؛ نخست --record اجرا شود")
        return 1

    with open(BASELINE, encoding="utf-8") as f:
        base = json.load(f)

    print("\n۵) انطباق با مرجع")
    ok = True
    d1 = compare("سنجه", base.get("metrics", {}), metrics, None)
    for k, bv, cv in d1:
        ok = False
        print("   ✗ سنجه %s: مرجع %s ← اکنون %s" % (k, bv, cv))
    print("   سنجه‌ها: %s (%d مورد بررسی شد)"
          % ("منطبق ✓" if not d1 else "منحرف ✗", len(base.get("metrics", {}))))

    d2 = []
    for title, prob in base.get("cascade", {}).items():
        d2 += [(title + " / " + k, bv, cv) for k, bv, cv in
               compare(title, prob, casc.get(title, {}), None)]
    for k, bv, cv in d2:
        ok = False
        print("   ✗ آبشار %s: مرجع %s ← اکنون %s" % (k, bv, cv))
    print("   تست‌های آبشاری: %s (%d مورد)"
          % ("منطبق ✓" if not d2 else "منحرف ✗", len(base.get("cascade", {}))))

    be, ce = base.get("errors", {}), rep["hist"]
    if be != ce:
        ok = False
        print("   ✗ الگوی خطاها تغییر کرده: مرجع %s ← اکنون %s" % (be, ce))
    else:
        print("   الگوی خطاها: منطبق ✓ (%d موردِ موردانتظار)" % sum(ce.values()))

    print("-" * 78)
    print("نتیجه: %s" % ("مدل با مرجع منطبق است ✅" if ok else "انحراف از مرجع ❌"))
    print("=" * 78)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(record="--record" in sys.argv))
