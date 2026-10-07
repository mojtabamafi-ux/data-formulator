# -*- coding: utf-8 -*-
"""لایهٔ موتور — بخشِ ب: دارایی‌ها و هزینه.

  ENG_CAPEX : برنامهٔ تزریق سرمایه، سودِ سرمایه‌گذاری حین ساخت (IDC)، جایگزینی
              تجهیزات، استهلاکِ خطِ مستقیمِ سقف‌دار به‌تفکیک طبقه و ارزش دفتری.
  ENG_ALLOC : تخصیص هزینهٔ مراکز خدماتی به ایستگاه‌ها.
  ENG_COST  : ماتریس اقلام هزینه (DPC / OPC / FPC) و بهای تمام‌شده.
  ENG_WC    : سرمایه در گردش و چرخهٔ تبدیل نقد.
  ENG_FCF   : صورت سود و زیان، جریان نقد، ترازنامه و شاخص‌های پوشش.
  ENG_DCF   : نرخ تنزیل، ارزش فعلی، ارزش پایانی و شاخص‌های بازدهی.
  ENG_LEV   : اهرم عملیاتی/مالی و نقطهٔ سر‌به‌سر.
  ENG_SENS  : ماتریس‌های دوبعدی حساسیت.
"""

import common
import inputs
import spec

COL_ACTIVE = 7          # ستونِ G در ENG_SCEN: مقدارِ فعالِ ضرایب سناریو


def scen_idx(b, tag, t):
    """شاخصِ قیمتیِ ENG_SCEN (سطرهای بدون کلید)"""
    return b.tagref("ENG_SCEN", tag, t)


def scen_coef(b, key):
    """ضریبِ فعالِ سناریو از ستونِ G"""
    return b.cellref("ENG_SCEN", b.row_of("ENG_SCEN", key), COL_ACTIVE)


# ==================================================================== ENG_CAPEX
def eng_capex(b):
    sh = b.sheet("ENG_CAPEX", "چرخهٔ عمر دارایی‌ها و CAPEX",
                 "برنامه تزریق سرمایه‌گذاری، جایگزینی تجهیزات، استهلاک و ارزش دفتری "
                 "دارایی‌ها.", "engine")
    S = "ENG_CAPEX"

    def R(key, t=0):
        return b.ref(S, key, t)

    def CAP(j, k):
        """ستونِ k از جدولِ طبقهٔ j در IN_FIN (0=مبلغ کل، 1=سهم ارزی، 2=عمر مفید،
        3=دورهٔ جایگزینی، 4=درصد جایگزینی، 5=ارزش اسقاط، 6..8=تزریق t=0..2)"""
        code = inputs.CAPEX_ITEMS[j][0]
        return b.cellref("IN_FIN", b.row_of("IN_FIN", code), spec.C0 + k)

    # تعدیلِ ریالی/ارزیِ هر طبقه: بخشِ ریالی با تورم و بخشِ ارزی با شاخصِ ارز
    def adj(j, t):
        return "((1-%s)*%s+%s*%s)" % (CAP(j, 1), scen_idx(b, "idx_infl", t),
                                      CAP(j, 1), scen_idx(b, "idx_fx", t))

    # ------------------------------------------- الف) برنامه تزریق سرمایه
    sh.section("الف) برنامه تزریق سرمایه‌گذاری ثابت")
    for j, (_code, name, *_rest) in enumerate(inputs.CAPEX_ITEMS):
        key = "CAPEX.cap.%d" % j
        sh.add_row(key, "سرمایه‌گذاری %s" % name, "میلیون ریال", "num")

        def inject(t, j=j):
            tt = b.ref("IN_MACRO", "MAC.t", t)
            return "(IF(%s=0,%s,IF(%s=1,%s,IF(%s=2,%s,0))))" % (
                tt, CAP(j, 6), tt, CAP(j, 7), tt, CAP(j, 8))

        b.frow(S, key, lambda t, j=j, inject=inject: "=%s*%s*%s*%s"
               % (CAP(j, 0), adj(j, t), inject(t), scen_coef(b, "CAPEX")))

    sh.add_row("CAPEX.cap_total", "جمع سرمایه‌گذاری ثابت", "میلیون ریال", "num")
    b.frow(S, "CAPEX.cap_total", lambda t: "=%s"
           % "+".join(R("CAPEX.cap.%d" % j, t) for j in range(spec.NCAP)))

    row = sh.r
    sh.add_row("CAPEX.cap_cum", "مانده تجمعی تزریق سرمایه‌گذاری", "میلیون ریال", "num")
    b.mark(S, "CAPEX.cap_cum", row)
    b.frow_at(S, row, lambda t: "=" + R("CAPEX.cap_total", 0) if t == 0
              else "=" + R("CAPEX.cap_cum", t - 1) + "+" + R("CAPEX.cap_total", t))

    sh.add_row("CAPEX.fxsh", "سهم ارزیِ وزنیِ CAPEX", "درصد", "pct")
    num = "+".join("(%s*%s)" % (CAP(j, 0), CAP(j, 1)) for j in range(spec.NCAP))
    den = "+".join(CAP(j, 0) for j in range(spec.NCAP))
    b.frow(S, "CAPEX.fxsh", lambda t: "=IFERROR((%s)/(%s),0)" % (num, den))

    # ------------------------------------------- ب) سود سرمایه‌گذاری حین ساخت
    sh.gap()
    sh.section("ب) سود سرمایه‌گذاری حین ساخت (IDC)")
    sh.add_row("CAPEX.debt_open", "مانده تسهیلات در ابتدای دوره (دوره ساخت)",
               "میلیون ریال", "num")
    b.frow(S, "CAPEX.debt_open", lambda t:
           "=%s*(0)" % b.ref("IN_FIN", "FIN.debt_share", 0) if t == 0
           else "=%s*(%s+%s)" % (b.ref("IN_FIN", "FIN.debt_share", 0),
                                 R("CAPEX.cap_cum", t - 1), R("CAPEX.idc_cum", t - 1)))
    sh.add_row("CAPEX.idc", "IDC (سرمایه‌ای‌شده)", "میلیون ریال", "num")
    b.frow(S, "CAPEX.idc", lambda t: "=IF(%s=1,%s*(%s+%s*%s/2),0)"
           % (b.ref("IN_MACRO", "MAC.constr_flg", t), b.ref("IN_FIN", "FIN.kd", 0),
              R("CAPEX.debt_open", t), b.ref("IN_FIN", "FIN.debt_share", 0),
              R("CAPEX.cap_total", t)))
    row = sh.r
    sh.add_row("CAPEX.idc_cum", "مانده تجمعی IDC", "میلیون ریال", "num")
    b.mark(S, "CAPEX.idc_cum", row)
    b.frow_at(S, row, lambda t: "=" + R("CAPEX.idc", 0) if t == 0
              else "=" + R("CAPEX.idc_cum", t - 1) + "+" + R("CAPEX.idc", t))

    # ------------------------------------------- ج) جایگزینی دارایی‌ها
    sh.gap()
    sh.section("ج) جایگزینی دارایی‌ها")
    for j, (_code, name, *_rest) in enumerate(inputs.CAPEX_ITEMS):
        key = "CAPEX.repl.%d" % j
        sh.add_row(key, "جایگزینی %s" % name, "میلیون ریال", "num")
        b.frow(S, key, lambda t, j=j: "=IF(OR(%s<=0,%s<=0,%s=0),0,IF(AND(MOD(%s,%s)=0,%s<=%s),"
                                      "%s*%s*%s*%s,0))"
               % (CAP(j, 2), CAP(j, 3), b.ref("IN_MACRO", "MAC.t", t),
                  b.ref("IN_MACRO", "MAC.t", t), CAP(j, 3),
                  b.ref("IN_MACRO", "MAC.t", t), b.ref("IN_MACRO", "MAC.horizon", 0),
                  CAP(j, 0), CAP(j, 4), adj(j, t), scen_coef(b, "CAPEX")))
    sh.add_row("CAPEX.repl_total", "جمع جایگزینی دارایی‌ها", "میلیون ریال", "num")
    b.frow(S, "CAPEX.repl_total", lambda t: "=%s"
           % "+".join(R("CAPEX.repl.%d" % j, t) for j in range(spec.NCAP)))

    # ------------------------------------------- د) دارایی‌ها و استهلاک
    sh.gap()
    sh.section("د) دارایی‌های ثابت و استهلاک")
    sh.add_row("CAPEX.add", "اضافات دارایی (CAPEX + جایگزینی + IDC)", "میلیون ریال", "num")
    b.frow(S, "CAPEX.add", lambda t: "=%s+%s+%s"
           % (R("CAPEX.cap_total", t), R("CAPEX.repl_total", t), R("CAPEX.idc", t)))

    row = sh.r
    sh.add_row("CAPEX.gross_open", "ناخالص دارایی‌های ثابت - افتتاحیه", "میلیون ریال", "num")
    b.mark(S, "CAPEX.gross_open", row)
    b.frow_at(S, row, lambda t: "=0" if t == 0 else "=" + R("CAPEX.gross", t - 1))
    sh.add_row("CAPEX.gross", "ناخالص دارایی‌های ثابت - اختتامیه", "میلیون ریال", "num")
    b.frow(S, "CAPEX.gross", lambda t: "=%s+%s"
           % (R("CAPEX.gross_open", t), R("CAPEX.add", t)))

    for j, (_code, name, *_rest) in enumerate(inputs.CAPEX_ITEMS):
        key = "CAPEX.grossc.%d" % j
        sh.add_row(key, "ناخالص طبقهٔ %s" % name, "میلیون ریال", "num")
        # ناخالصِ هر طبقه تجمعی است تا استهلاکِ خطِ مستقیم روی ماندهٔ همان طبقه
        # محاسبه شود و ارزش دفتری هیچ‌گاه منفی نشود.
        b.frow(S, key, lambda t, j=j, key=key:
               "=(%s+%s)" % (R("CAPEX.cap.%d" % j, t), R("CAPEX.repl.%d" % j, t))
               if t == 0 else
               "=" + R(key, t - 1) + "+(%s+%s)"
               % (R("CAPEX.cap.%d" % j, t), R("CAPEX.repl.%d" % j, t)))

    for j, (_code, name, *_rest) in enumerate(inputs.CAPEX_ITEMS):
        key = "CAPEX.dep.%d" % j
        sh.add_row(key, "استهلاک %s" % name, "میلیون ریال", "num")
        b.frow(S, key, lambda t, j=j: "=IF(%s<=0,0,MAX(0,MIN(%s/%s,%s-%s)))"
               % (CAP(j, 2), R("CAPEX.grossc.%d" % j, t), CAP(j, 2),
                  R("CAPEX.grossc.%d" % j, t),
                  "0" if t == 0 else R("CAPEX.adepc.%d" % j, t - 1)))

    for j, (_code, name, *_rest) in enumerate(inputs.CAPEX_ITEMS):
        key = "CAPEX.adepc.%d" % j
        row = sh.r
        sh.add_row(key, "استهلاک انباشتهٔ طبقهٔ %s" % name, "میلیون ریال", "num")
        b.mark(S, key, row)
        b.frow_at(S, row, lambda t, j=j: "=" + R("CAPEX.dep.%d" % j, 0) if t == 0
                  else "=" + R("CAPEX.adepc.%d" % j, t - 1) + "+" + R("CAPEX.dep.%d" % j, t))

    sh.add_row("CAPEX.dep_total", "جمع استهلاک سالانه", "میلیون ریال", "num")
    b.frow(S, "CAPEX.dep_total", lambda t: "=%s"
           % "+".join(R("CAPEX.dep.%d" % j, t) for j in range(spec.NCAP)))

    row = sh.r
    sh.add_row("CAPEX.accumdep_open", "استهلاک انباشته - افتتاحیه", "میلیون ریال", "num")
    b.mark(S, "CAPEX.accumdep_open", row)
    b.frow_at(S, row, lambda t: "=0" if t == 0 else "=" + R("CAPEX.accumdep", t - 1))
    sh.add_row("CAPEX.accumdep", "استهلاک انباشته - اختتامیه", "میلیون ریال", "num")
    b.frow(S, "CAPEX.accumdep", lambda t: "=%s+%s"
           % (R("CAPEX.accumdep_open", t), R("CAPEX.dep_total", t)))
    sh.add_row("CAPEX.nbv", "ارزش دفتری خالص (NBV)", "میلیون ریال", "num")
    b.frow(S, "CAPEX.nbv", lambda t: "=%s-%s" % (R("CAPEX.gross", t), R("CAPEX.accumdep", t)))
    sh.add_row("CAPEX.salvage", "ارزش اسقاط برآوردی در پایان افق", "میلیون ریال", "num")
    b.frow(S, "CAPEX.salvage", lambda t: "=IF(%s<%s,0,%s)"
           % (b.ref("IN_MACRO", "MAC.t", t), b.ref("IN_MACRO", "MAC.horizon", 0),
              "+".join("(%s*%s*%s)" % (CAP(j, 0), CAP(j, 5), scen_idx(b, "idx_infl", t))
                       for j in range(spec.NCAP))))
    return sh


def build(b):
    eng_capex(b)
