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


# ==================================================================== ENG_ALLOC
def eng_alloc(b):
    sh = b.sheet("ENG_ALLOC", "تخصیص مراکز هزینه (تولیدی و خدماتی)",
                 "تخصیص هزینه‌های مراکز خدماتی به ایستگاه‌های ۱ تا ۶ بر اساس محرک‌های هزینه.",
                 "engine")
    S = "ENG_ALLOC"
    NS = spec.NS
    r_thr = b.row_of("IN_TECH", "THROUGHPUT")
    r_wat = b.row_of("IN_TECH", "WATER")
    drv = b.rng_at("IN_TECH", r_thr, spec.C0, r_wat, spec.C0 + NS - 1)   # $D$47:$I$52
    lab = b.rng_at("IN_TECH", r_thr, 2, r_wat, 2)                       # $B$47:$B$52

    def SC(i, k):
        """ستونِ k از ردیفِ مرکزِ خدماتی i در IN_COST (1=محرک، 2=هزینهٔ پایه)"""
        code = inputs.SERVICE_CENTERS[i][0]
        return b.cellref("IN_COST", b.row_of("IN_COST", code), spec.C0 + k)

    # ------------------------------------------- الف) هزینه مستقیم مراکز
    sh.section("الف) هزینه مستقیم مراکز خدماتی")
    for i, (_code, name, _drv, _amt, _beh) in enumerate(inputs.SERVICE_CENTERS):
        sh.add_row("ALC.cost.%d" % (i + 1), "هزینه مرکز %d: %s" % (i + 1, name),
                   "میلیون ریال", "num")
        b.frow(S, "ALC.cost.%d" % (i + 1), lambda t, i=i: "=%s*%s*%s"
               % (SC(i, 2), scen_idx(b, "idx_infl", t),
                  b.ref("IN_MACRO", "MAC.ops_flg", t)))

    # ------------------------------------------- ب) جمع محرک و وزن تخصیص
    sh.gap()
    sh.section("ب) جمع محرک و وزن تخصیص هر مرکز")
    for i in range(len(inputs.SERVICE_CENTERS)):
        key = "ALC.dtot.%d" % (i + 1)
        sh.add_row(key, "جمع محرک مرکز %d" % (i + 1), "─", "num")
        # محرک‌ها عددِ طراحی‌اند و به دوره وابسته نیستند → تنها یک سلول
        b.fcell(S, sh.r - 1, spec.C0, "=SUM(INDEX(%s,MATCH(%s,%s,0),0))"
                % (drv, SC(i, 1), lab))

    for i in range(len(inputs.SERVICE_CENTERS)):
        key = "ALC.w.%d" % (i + 1)
        row = sh.r
        sh.add_row(key, "وزن تخصیص مرکز %d به ایستگاه‌ها" % (i + 1), "درصد", "num3")
        # فقط ستون‌های ایستگاه‌ها پر می‌شود (این ردیف سریِ زمانی نیست)
        for j in range(NS):
            b.fcell(S, row, spec.C0 + j,
                    "=IFERROR(INDEX(%s,MATCH(%s,%s,0),%d)/%s,0)"
                    % (drv, SC(i, 1), lab, j + 1,
                       b.ref(S, "ALC.dtot.%d" % (i + 1), 0)))

    # ------------------------------------------- ج) ماتریس تخصیص
    sh.gap()
    sh.section("ج) ماتریس تخصیص (مرکز × ایستگاه × سال)")
    matrix_rows = {}
    for i in range(len(inputs.SERVICE_CENTERS)):
        for j in range(NS):
            row = sh.r
            sh.add_row(None, "مرکز %d ← ایستگاه %d" % (i + 1, j + 1), "میلیون ریال", "num")
            matrix_rows[(i, j)] = row
            b.frow_at(S, row, lambda t, i=i, j=j: "=%s*%s"
                      % (b.ref(S, "ALC.cost.%d" % (i + 1), t),
                         b.cellref(S, b.row_of(S, "ALC.w.%d" % (i + 1)), spec.C0 + j)))

    # ------------------------------------------- د) جمع تخصیص‌یافته به هر ایستگاه
    sh.gap()
    sh.section("د) جمع تخصیص‌یافته به هر ایستگاه")
    from openpyxl.utils import get_column_letter
    for j in range(NS):
        sh.add_row("ALC.station.%d" % j, "سهم ایستگاه %d از مراکز خدماتی" % (j + 1),
                   "میلیون ریال", "num")
        b.frow(S, "ALC.station.%d" % j, lambda t, j=j: "=%s" % "+".join(
            "%s%d" % (get_column_letter(spec.C0 + t),
                      matrix_rows[(i, j)])
            for i in range(len(inputs.SERVICE_CENTERS))))
    return sh


# ====================================================================== ENG_WC
def eng_wc(b):
    sh = b.sheet("ENG_WC", "سرمایه در گردش و چرخه تبدیل نقد",
                 "مانده موجودی‌ها، دریافتنی‌ها و پرداختنی‌ها و استنتاج چرخه تبدیل "
                 "نقد (CCC).", "engine")
    S = "ENG_WC"

    def R(key, t=0):
        return b.ref(S, key, t)

    def WC(key):
        return b.ref("IN_WC", key, 0)

    def dpc_opc(t):
        return "(%s+%s)" % (b.ref("ENG_COST", "EC.dpc", t), b.ref("ENG_COST", "EC.opc", t))

    sh.section("الف) اجزای سرمایه در گردش")
    sh.add_row("WC.inv_rm", "موجودی مواد اولیه", "میلیون ریال", "num")
    b.frow(S, "WC.inv_rm", lambda t: "=%s/365*%s"
           % (b.ref("ENG_COST", "EC.rm", t), WC("WC.dio_rm")))
    for key, lab, wk in [("WC.inv_wip", "موجودی کالای در جریان ساخت", "WC.dio_wip"),
                         ("WC.inv_fg", "موجودی محصول نهایی", "WC.dio_fg")]:
        sh.add_row(key, lab, "میلیون ریال", "num")
        b.frow(S, key, lambda t, wk=wk: "=%s/365*%s" % (dpc_opc(t), WC(wk)))
    sh.add_row("WC.inv_cons", "موجودی مواد مصرفی و قطعات یدکی", "میلیون ریال", "num")
    b.frow(S, "WC.inv_cons", lambda t: "=(%s+%s)/365*%s"
           % (b.ref("ENG_COST", "EC.tot.1", t), b.ref("ENG_COST", "EC.tot.9", t),
              WC("WC.dio_cons")))
    sh.add_row("WC.inv", "جمع موجودی‌ها", "میلیون ریال", "num")
    b.frow(S, "WC.inv", lambda t: "=%s" % "+".join(
        R(k, t) for k in ("WC.inv_rm", "WC.inv_wip", "WC.inv_fg", "WC.inv_cons")))
    sh.add_row("WC.ar", "حساب‌های دریافتنی", "میلیون ریال", "num")
    b.frow(S, "WC.ar", lambda t: "=%s/365*%s"
           % (b.ref("ENG_OPS", "OPS.rev", t), WC("WC.dso")))
    sh.add_row("WC.ap", "حساب‌های پرداختنی", "میلیون ریال", "num")
    b.frow(S, "WC.ap", lambda t: "=%s/365*%s" % (dpc_opc(t), WC("WC.dpo")))
    sh.add_row("WC.tax_pay", "مالیات قابل پرداخت", "میلیون ریال", "num")
    b.frow(S, "WC.tax_pay", lambda t: "=%s/365*%s"
           % (b.ref("ENG_COST", "EC.tax_op", t), WC("WC.dpo_tax")))
    sh.add_row("WC.nwc", "خالص سرمایه در گردش (NWC)", "میلیون ریال", "num")
    b.frow(S, "WC.nwc", lambda t: "=%s+%s-%s-%s"
           % (R("WC.inv", t), R("WC.ar", t), R("WC.ap", t), R("WC.tax_pay", t)))
    sh.add_row("WC.dwc", "تغییرات سرمایه در گردش (ΔWC)", "میلیون ریال", "num")
    b.frow(S, "WC.dwc", lambda t: "=" + R("WC.nwc", 0) if t == 0
           else "=" + R("WC.nwc", t) + "-" + R("WC.nwc", t - 1))
    sh.add_row("WC.stloan", "تسهیلات کوتاه‌مدت (مانده)", "میلیون ریال", "num")
    b.frow(S, "WC.stloan", lambda t: "=MAX(0,%s*%s)" % (R("WC.nwc", t), WC("WC.wc_loan_share")))
    sh.add_row("WC.dstloan", "تغییرات تسهیلات کوتاه‌مدت", "میلیون ریال", "num")
    b.frow(S, "WC.dstloan", lambda t: "=" + R("WC.stloan", 0) if t == 0
           else "=" + R("WC.stloan", t) + "-" + R("WC.stloan", t - 1))

    sh.gap()
    sh.section("ب) چرخه تبدیل نقد (CCC)")
    sh.add_row("WC.dio", "دوره گردش موجودی (DIO)", "روز", "num1")
    b.frow(S, "WC.dio", lambda t: "=IFERROR(%s/(%s/365),0)" % (R("WC.inv", t), dpc_opc(t)))
    sh.add_row("WC.dso_calc", "دوره وصول مطالبات (DSO)", "روز", "num1")
    b.frow(S, "WC.dso_calc", lambda t: "=IFERROR(%s/(%s/365),0)"
           % (R("WC.ar", t), b.ref("ENG_OPS", "OPS.rev", t)))
    sh.add_row("WC.dpo_calc", "دوره پرداخت بدهی‌ها (DPO)", "روز", "num1")
    b.frow(S, "WC.dpo_calc", lambda t: "=IFERROR(%s/(%s/365),0)" % (R("WC.ap", t), dpc_opc(t)))
    sh.add_row("WC.ccc", "چرخه تبدیل نقد (CCC)", "روز", "num1")
    b.frow(S, "WC.ccc", lambda t: "=%s+%s-%s"
           % (R("WC.dio", t), R("WC.dso_calc", t), R("WC.dpo_calc", t)))
    sh.add_row("WC.cash_opex", "هزینه نقدی سالانه (مبنای حداقل نقد)", "میلیون ریال", "num")
    b.frow(S, "WC.cash_opex", lambda t: "=%s+%s"
           % (b.ref("ENG_COST", "EC.var_total", t), b.ref("ENG_COST", "EC.fixed_cash", t)))
    sh.add_row("WC.min_cash", "حداقل موجودی نقدی", "میلیون ریال", "num")
    b.frow(S, "WC.min_cash", lambda t: "=MAX(0,%s)/365*%s"
           % (R("WC.cash_opex", t), WC("WC.cashdays")))
    return sh


# ===================================================================== ENG_DCF
def eng_dcf(b):
    sh = b.sheet("ENG_DCF", "محاسبات DCF و ارزش‌گذاری",
                 "ساخت WACC، تنزیل FCFF، محاسبهٔ NPV، IRR، MIRR، دوره بازپرداشت "
                 "و ارزش بنگاه.", "engine")
    S = "ENG_DCF"
    LT = spec.NY - 1

    def R(key, t=0):
        return b.ref(S, key, t)

    sh.section("الف) ساختار سرمایه و نرخ تنزیل")
    sh.add_row("DCF.ke", "بازده مورد انتظار سهامداران (Ke)", "درصد", "pct")
    b.frow(S, "DCF.ke", lambda t: "=%s+%s*%s+%s"
           % (b.ref("ENG_SCEN", "SCEN.rf", t), b.ref("IN_MACRO", "MAC.beta", 0),
              b.ref("IN_MACRO", "MAC.erp", 0), b.ref("IN_MACRO", "MAC.crp", 0)))
    sh.add_row("DCF.kd_at", "هزینه بدهی پس از مالیات (Kd×(1−t))", "درصد", "pct")
    b.frow(S, "DCF.kd_at", lambda t: "=(%s+%s)*(1-%s)"
           % (b.ref("IN_FIN", "FIN.kd", 0), scen_coef(b, "KD"),
              b.ref("IN_MACRO", "MAC.tax", 0)))
    sh.add_row("DCF.we", "وزن حقوق صاحبان سهام", "درصد", "pct")
    b.frow(S, "DCF.we", lambda t: "=1-%s" % b.ref("IN_FIN", "FIN.debt_share", 0))
    sh.add_row("DCF.wd", "وزن بدهی", "درصد", "pct")
    b.frow(S, "DCF.wd", lambda t: "=%s" % b.ref("IN_FIN", "FIN.debt_share", 0))
    sh.add_row("DCF.wacc", "میانگین موزون هزینه سرمایه (WACC)", "درصد", "pct")
    b.frow(S, "DCF.wacc", lambda t: "=%s*%s+%s*%s+%s"
           % (R("DCF.we", t), R("DCF.ke", t), R("DCF.wd", t), R("DCF.kd_at", t),
              scen_coef(b, "WACC")))

    row = sh.r
    sh.add_row("DCF.df", "ضریب تنزیل", "ضریب", "num3")
    b.mark(S, "DCF.df", row)
    b.frow_at(S, row, lambda t: "=1" if t == 0
              else "=" + R("DCF.df", t - 1) + "/(1+" + R("DCF.wacc", t) + ")")

    sh.gap()
    sh.section("ب) تنزیل جریان‌های نقدی و ارزش پایانی")
    sh.add_row("DCF.fcff", "جریان نقد آزاد پروژه (FCFF)", "میلیون ریال", "num")
    b.frow(S, "DCF.fcff", lambda t: "=" + b.ref("ENG_FCF", "FCF.fcff", t))
    sh.add_row("DCF.pv", "ارزش فعلی FCFF", "میلیون ریال", "num")
    b.frow(S, "DCF.pv", lambda t: "=%s*%s" % (R("DCF.fcff", t), R("DCF.df", t)))
    row = sh.r
    sh.add_row("DCF.pv_cum", "ارزش فعلی تجمعی", "میلیون ریال", "num")
    b.mark(S, "DCF.pv_cum", row)
    b.frow_at(S, row, lambda t: "=" + R("DCF.pv", 0) if t == 0
              else "=" + R("DCF.pv_cum", t - 1) + "+" + R("DCF.pv", t))
    sh.add_row("DCF.tv", "ارزش پایانی (Terminal Value)", "میلیون ریال", "num")
    b.frow(S, "DCF.tv", lambda t:
           "=IF(%s<%s,0,IF(%s=1,MAX(0,%s)*(1+%s)/MAX(0.005,%s-%s),%s*%s))" % (
               b.ref("IN_MACRO", "MAC.t", t), b.ref("IN_MACRO", "MAC.horizon", 0),
               b.ref("IN_MACRO", "MAC.tv_method", 0),
               b.ref("ENG_FCF", "FCF.fcff_rep", LT), b.ref("IN_MACRO", "MAC.g_term", 0),
               R("DCF.wacc", LT), b.ref("IN_MACRO", "MAC.g_term", 0),
               b.ref("ENG_FCF", "FCF.ebitda", LT), b.ref("IN_MACRO", "MAC.exit_mult", 0)))
    sh.add_row("DCF.pv_tv", "ارزش فعلی ارزش پایانی", "میلیون ریال", "num")
    b.frow(S, "DCF.pv_tv", lambda t: "=%s*%s" % (R("DCF.tv", t), R("DCF.df", t)))
    sh.add_row("DCF.fcf_tv", "جریان قابل بازدهی (FCFF + ارزش پایانی)", "میلیون ریال", "num")
    b.frow(S, "DCF.fcf_tv", lambda t: "=%s+%s" % (R("DCF.fcff", t), R("DCF.tv", t)))

    # ------------------------------------------- ج) شاخص‌های ارزش‌گذاری
    sh.gap()
    sh.section("ج) شاخص‌های ارزش‌گذاری")
    rng_pv = b.rng(S, "DCF.pv", 0, LT)
    rng_pvtv = b.rng(S, "DCF.pv_tv", 0, LT)
    rng_tv = b.rng(S, "DCF.fcf_tv", 0, LT)
    rng_pvc = b.rng(S, "DCF.pv_cum", 0, LT)
    rng_wacc = b.rng(S, "DCF.wacc", 0, LT)
    rng_cum = b.rng("ENG_FCF", "FCF.fcff_cum", 0, LT)
    rng_t = b.rng("IN_MACRO", "MAC.t", 0, LT)

    def payback(cum):
        n = "COUNTIF(%s,\"<0\")" % cum
        return ('=IF(%s=0,0,IF(%s>=13,NA(),INDEX(%s,%s)+(0-INDEX(%s,%s))'
                '/(INDEX(%s,%s+1)-INDEX(%s,%s))))'
                % (n, n, rng_t, n, cum, n, cum, n, cum, n))

    specs = [("DCF.npv", "ارزش فعلی خالص (NPV)", "میلیون ریال", "num",
              "=SUM(%s)+SUM(%s)" % (rng_pv, rng_pvtv)),
             ("DCF.irr", "نرخ بازده داخلی پروژه (IRR)", "درصد", "pct",
              "=IFERROR(IRR(%s),NA())" % rng_tv),
             ("DCF.mirr", "نرخ بازده داخلی تعدیل‌شده (MIRR)", "درصد", "pct",
              "=IFERROR(MIRR(%s,(%s+%s),AVERAGE(%s)),NA())"
              % (rng_tv, b.ref("IN_FIN", "FIN.kd", 0), scen_coef(b, "KD"), rng_wacc)),
             ("DCF.irr_eq", "نرخ بازده داخلی سهامداران (IRR روی FCFE)", "درصد", "pct",
              "=IFERROR(IRR(%s),NA())" % b.rng("ENG_FCF", "FCF.fcfe", 0, LT)),
             ("DCF.pay", "دوره بازپرداشت ساده (سال)", "سال", "num1", payback(rng_cum)),
             ("DCF.pay_d", "دوره بازپرداشت تنزیلی (سال)", "سال", "num1", payback(rng_pvc)),
             ("DCF.pi", "شاخص سودآوری (PI)", "ضریب", "num2",
              '=IFERROR((SUMIF(%s,">0")+SUM(%s))/ABS(SUMIF(%s,"<0")),NA())'
              % (rng_pv, rng_pvtv, rng_pv)),
             ("DCF.ev", "ارزش بنگاه (EV) - ارزش فعلی عملیات", "میلیون ریال", "num",
              "=SUM(%s)-%s+SUM(%s)" % (rng_pv, R("DCF.pv", 0), rng_pvtv))]
    for tag, lab, unit, fmt, f in specs:
        row = sh.r
        sh.add_row(None, lab, unit, fmt)
        b.mark(S, tag, row)
        b.fcell(S, row, spec.C0, f)

    sh.add_row("DCF.eq_value", "ارزش حقوق صاحبان سهام (EV − بدهی خالص)",
               "میلیون ریال", "num")
    b.frow(S, "DCF.eq_value", lambda t: "=%s-(%s-%s)"
           % (b.tagref(S, "DCF.ev", 0), b.ref("ENG_FCF", "FCF.debt_close", t),
              b.ref("ENG_FCF", "FCF.cash", t)))
    return sh


# ====================================================================== ENG_LEV
def eng_lev(b):
    sh = b.sheet("ENG_LEV", "تحلیل اهرم عملیاتی و مالی",
                 "درجه اهرم عملیاتی و مالی، نقطه سر‌به‌سر و اثر تغییر حجم فروش و "
                 "ساختار سرمایه بر سودآوری.", "engine")
    S = "ENG_LEV"

    def R(key, t=0):
        return b.ref(S, key, t)

    def F(key, t=0):
        return b.ref("ENG_FCF", key, t)

    def C(key, t=0):
        return b.ref("ENG_COST", key, t)

    sh.section("الف) حاشیه مشارکت و نقطه سر‌به‌سر")
    sh.add_row("LEV.rev", "درآمد خالص", "میلیون ریال", "num")
    b.frow(S, "LEV.rev", lambda t: "=" + F("FCF.rev", t))
    sh.add_row("LEV.var", "هزینه‌های متغیر", "میلیون ریال", "num")
    b.frow(S, "LEV.var", lambda t: "=" + C("EC.var_total", t))
    sh.add_row("LEV.cm", "حاشیه مشارکت", "میلیون ریال", "num")
    b.frow(S, "LEV.cm", lambda t: "=%s-%s" % (R("LEV.rev", t), R("LEV.var", t)))
    sh.add_row("LEV.cm_ratio", "نسبت حاشیه مشارکت", "درصد", "pct")
    b.frow(S, "LEV.cm_ratio", lambda t: "=IFERROR(%s/%s,0)" % (R("LEV.cm", t), R("LEV.rev", t)))
    sh.add_row("LEV.cm_ton", "حاشیه مشارکت هر تن", "میلیون ریال/تن", "num1")
    b.frow(S, "LEV.cm_ton", lambda t: "=IFERROR(%s/%s,0)"
           % (R("LEV.cm", t), b.ref("ENG_OPS", "OPS.sold_tons", t)))
    sh.add_row("LEV.fixed_cash", "هزینه ثابت نقدی", "میلیون ریال", "num")
    b.frow(S, "LEV.fixed_cash", lambda t: "=" + C("EC.fixed_cash", t))
    sh.add_row("LEV.be_ton", "نقطه سر‌به‌سر (تن محصول)", "تن", "num")
    b.frow(S, "LEV.be_ton", lambda t: "=IFERROR(%s/%s,0)"
           % (R("LEV.fixed_cash", t), R("LEV.cm_ton", t)))
    sh.add_row("LEV.be_input", "نقطه سر‌به‌سر (تن باتری ورودی)", "تن", "num")
    b.frow(S, "LEV.be_input", lambda t: "=IFERROR(%s/%s*%s,0)"
           % (R("LEV.be_ton", t), b.ref("ENG_OPS", "OPS.sold_tons", t),
              b.ref("ENG_OPS", "OPS.intake", t)))
    sh.add_row("LEV.mos", "حاشیه ایمنی", "درصد", "pct")
    b.frow(S, "LEV.mos", lambda t: "=IFERROR(1-%s/%s,0)"
           % (R("LEV.be_ton", t), b.ref("ENG_OPS", "OPS.sold_tons", t)))

    sh.gap()
    sh.section("ب) درجات اهرم")
    sh.add_row("LEV.dol", "درجه اهرم عملیاتی (DOL)", "ضریب", "num2")
    b.frow(S, "LEV.dol", lambda t: "=IFERROR(%s/%s,0)" % (R("LEV.cm", t), F("FCF.ebit", t)))
    sh.add_row("LEV.dfl", "درجه اهرم مالی (DFL)", "ضریب", "num2")
    b.frow(S, "LEV.dfl", lambda t: "=IFERROR(%s/%s,0)" % (F("FCF.ebit", t), F("FCF.ebt", t)))
    sh.add_row("LEV.dtl", "درجه اهرم ترکیبی (DTL)", "ضریب", "num2")
    b.frow(S, "LEV.dtl", lambda t: "=%s*%s" % (R("LEV.dol", t), R("LEV.dfl", t)))
    sh.add_row("LEV.sens_ni", "تغییر سود خالص به ازای ۱٪ تغییر فروش", "درصد", "pct")
    b.frow(S, "LEV.sens_ni", lambda t: "=" + R("LEV.dtl", t))

    sh.gap()
    sh.section("ج) بازدهی و ساختار سرمایه")
    sh.add_row("LEV.roe", "بازده حقوق صاحبان سهام (ROE)", "درصد", "pct")
    b.frow(S, "LEV.roe", lambda t: "=IFERROR(%s/%s,0)" % (F("FCF.ni", t), F("FCF.equity", t)))
    sh.add_row("LEV.roa", "بازده دارایی‌ها (ROA)", "درصد", "pct")
    b.frow(S, "LEV.roa", lambda t: "=IFERROR(%s/%s,0)" % (F("FCF.ni", t), F("FCF.assets", t)))
    sh.add_row("LEV.roic", "بازده سرمایه به‌کارگرفته‌شده (ROIC)", "درصد", "pct")
    b.frow(S, "LEV.roic", lambda t: "=IFERROR(%s/(%s+%s),0)"
           % (F("FCF.nopat", t), F("FCF.ltd", t), F("FCF.equity", t)))
    sh.add_row("LEV.de", "نسبت بدهی به حقوق صاحبان سهام", "ضریب", "num2")
    b.frow(S, "LEV.de", lambda t: "=IFERROR(%s/%s,0)" % (F("FCF.ltd", t), F("FCF.equity", t)))
    sh.add_row("LEV.debt_ebitda", "نسبت بدهی به EBITDA", "ضریب", "num2")
    b.frow(S, "LEV.debt_ebitda", lambda t: "=IFERROR(%s/%s,0)"
           % (F("FCF.ltd", t), F("FCF.ebitda", t)))
    sh.add_row("LEV.net_debt_ebitda", "بدهی خالص به EBITDA", "ضریب", "num2")
    b.frow(S, "LEV.net_debt_ebitda", lambda t: "=IFERROR((%s-%s)/%s,0)"
           % (F("FCF.ltd", t), F("FCF.cash", t), F("FCF.ebitda", t)))
    return sh


def build(b):
    eng_capex(b)
    eng_alloc(b)
    eng_wc(b)
    eng_dcf(b)
    eng_lev(b)
