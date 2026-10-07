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


def build(b):
    eng_capex(b)
    eng_alloc(b)
    eng_wc(b)
