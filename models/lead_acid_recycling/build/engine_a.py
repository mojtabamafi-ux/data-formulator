# -*- coding: utf-8 -*-
"""لایهٔ موتور — بخشِ الف: سناریو و عملیات.

  ENG_SCEN : انتخابِ سناریوی فعال، مسیرهای تعدیل‌شده، شاخص‌های قیمتی و مدلِ سایهٔ
             سه سناریو (برای مقایسهٔ هم‌زمان پایه / خوش‌بینانه / بدبینانه).
  ENG_OPS  : ظرفیت مؤثر و محدودیتِ گلوگاه، جریان مواد و تراز جرم، محصولات قابل
             فروش، مصارف انرژی/آب و قیمت‌گذاری و درآمد.

هر سلول این دو شیت فرمول است و تنها به لایهٔ ورودی یا به همین لایه ارجاع می‌دهد.
"""

from openpyxl.utils import get_column_letter

import common
import spec

# ستون‌های خاص
COL_ACTIVE = 7          # ستونِ G: مقدارِ فعالِ ضرایب سناریو
LAST_T = spec.NY - 1    # آخرین دوره (t=12)


def L(t):
    """حرفِ ستونِ دورهٔ t (برای ارجاع‌های نسبیِ ستون)"""
    return get_column_letter(spec.C0 + t)


# ===================================================================== ENG_SCEN
def eng_scen(b):
    sh = b.sheet("ENG_SCEN", "موتور سناریو و حساسیت",
                 "اعمال ضرایب سناریوی فعال روی مفروضات کلیدی، ساخت مسیرهای تعدیل‌شده "
                 "و شاخص‌های قیمتی.", "engine")

    def R(key, t=0):
        return b.ref("ENG_SCEN", key, t)

    # ------------------------------------------- الف) سناریوی فعال و ضرایب
    sh.section("الف) سناریوی فعال و ضرایب تعدیل")
    r_active = sh.r
    sh.add_row(None, "سناریوی فعال (۱=پایه / ۲=خوش‌بینانه / ۳=بدبینانه)", "شاخص", "text")
    b.mark("ENG_SCEN", "active", r_active)
    b.fcell("ENG_SCEN", r_active, spec.C0, "=" + b.ref("IN_SCEN", "SCEN.active_in", 0))

    COEF = [("VOL", "ضریب حجم تولید و فروش — مقدار فعال", "ضریب", "num2"),
            ("PRICE", "ضریب قیمت فروش محصولات — مقدار فعال", "ضریب", "num2"),
            ("RM", "ضریب قیمت خرید باتری ضایعاتی — مقدار فعال", "ضریب", "num2"),
            ("VAR", "ضریب سایر هزینه‌های متغیر — مقدار فعال", "ضریب", "num2"),
            ("FIX", "ضریب هزینه‌های ثابت — مقدار فعال", "ضریب", "num2"),
            ("CAPEX", "ضریب سرمایه‌گذاری ثابت — مقدار فعال", "ضریب", "num2"),
            ("FX", "ضریب نرخ ارز — مقدار فعال", "ضریب", "num2"),
            ("WACC", "تعدیل نرخ تنزیل — مقدار فعال", "واحد درصد", "pct"),
            ("KD", "تعدیل نرخ سود تسهیلات — مقدار فعال", "واحد درصد", "pct")]
    for key, lab, unit, fmt in COEF:
        sh.add_row(key, lab, unit, fmt)
        for i in range(spec.NSCEN):            # D=پایه، E=خوش‌بینانه، F=بدبینانه
            b.fcell("ENG_SCEN", sh.r - 1, spec.C0 + i, "=" + b.ref("IN_SCEN", key, i))
        # G: ضریبِ فعال بر پایهٔ سناریوی انتخابی
        b.fcell("ENG_SCEN", sh.r - 1, COL_ACTIVE,
                "=INDEX(%s,%s)" % (b.rng("IN_SCEN", key, 0, spec.NSCEN - 1),
                                   b.tagref("ENG_SCEN", "active", 0)))

    def coef(key):
        """ضریبِ فعالِ سناریو (ستونِ G)"""
        return b.cellref("ENG_SCEN", b.row_of("ENG_SCEN", key), COL_ACTIVE)

    # ------------------------------------------- ب) مسیرهای تعدیل‌شدهٔ کلان
    sh.gap()
    sh.section("ب) مسیرهای تعدیل‌شدهٔ اقتصاد کلان")
    PATHS = [("SCEN.infl", "شاخص تورم عمومی", "MAC.infl", "pct"),
             ("SCEN.wage", "شاخص رشد دستمزد", "MAC.wage", "pct"),
             ("SCEN.energy", "شاخص رشد قیمت انرژی", "MAC.energy", "pct"),
             ("SCEN.ppi", "شاخص قیمت تولیدکننده", "MAC.ppi", "pct"),
             ("SCEN.rf", "نرخ سود بدون ریسک", "MAC.rf", "pct"),
             ("SCEN.gdp", "رشد GDP بخش هدف", "MAC.gdp", "pct"),
             ("SCEN.lme", "قیمت جهانی سرب LME", "MAC.lme", "num")]
    for key, lab, src, fmt in PATHS:
        sh.add_row(key, lab, "─", fmt)
        b.frow("ENG_SCEN", key, lambda t, src=src: "=" + b.ref("IN_MACRO", src, t))

    sh.add_row("SCEN.fx", "نرخ ارز (تعدیل‌شده)", "─", "num")
    b.frow("ENG_SCEN", "SCEN.fx",
           lambda t: "=" + b.ref("IN_MACRO", "MAC.fx", t) + "*" + coef("FX"))

    # ------------------------------------------- ج) شاخص‌های قیمتی تجمعی
    sh.gap()
    sh.section("ج) شاخص‌های قیمتی تجمعی (سال پایه = ۱)")
    IDX = [("idx_infl", "شاخص تورم عمومی", "SCEN.infl", "num2"),
           ("idx_wage", "شاخص دستمزد", "SCEN.wage", "num2"),
           ("idx_energy", "شاخص قیمت انرژی", "SCEN.energy", "num2"),
           ("idx_ppi", "شاخص قیمت تولیدکننده", "SCEN.ppi", "num2")]
    for tag, lab, src, fmt in IDX:
        row = sh.r
        sh.add_row(None, lab, "ضریب", fmt)
        b.mark("ENG_SCEN", tag, row)
        for t in range(spec.NY):
            f = "=1" if t == 0 else "=%s%d*(1+%s)" % (L(t - 1), row, R(src, t))
            b.fcell("ENG_SCEN", row, spec.C0 + t, f)

    row = sh.r
    sh.add_row(None, "شاخص نرخ ارز", None, "num2")
    b.mark("ENG_SCEN", "idx_fx", row)
    b.frow_at("ENG_SCEN", row,
              lambda t: "=" + R("SCEN.fx", t) + "/" + R("SCEN.fx", 0))

    row = sh.r
    sh.add_row(None, "شاخص ترکیبی LME×ارز", None, "num2")
    b.mark("ENG_SCEN", "idx_lme_fx", row)
    b.frow_at("ENG_SCEN", row, lambda t: "=(%s*%s)/(%s*%s)"
              % (R("SCEN.lme", t), R("SCEN.fx", t), R("SCEN.lme", 0), R("SCEN.fx", 0)))

    # ------------------------------------------- د) مدل سایهٔ سه سناریو
    sh.gap()
    sh.section("د) مدل سایهٔ سه سناریو (پایه / خوش‌بینانه / بدبینانه)")
    sh.note("توضیح: مدل تفصیلی برای «سناریوی فعال» اجرا می‌شود؛ این مدل سایه با "
            "استفاده از نسبت‌های ساختاریِ استخراج‌شده از همان مدل تفصیلی (فرمول‌محور، "
            "بدون عدد ثابت) شاخص‌های هر سه سناریو را هم‌زمان محاسبه می‌کند.")
    names = ["پایه", "خوش\u200cبینانه", "بدبینانه"]
    for i in range(spec.NSCEN):
        sh.subhead("سناریوی %d: %s" % (i + 1, names[i]))
        _shadow_block(b, sh, i)
        sh.gap()
    return sh


def _shadow_block(b, sh, i):
    """یک بلوکِ مدل سایه برای سناریوی i (پایه/خوش‌بینانه/بدبینانه)"""
    S = "ENG_SCEN"
    P = "SB%d." % (i + 1)

    def R(key, t=0):
        return b.ref(S, P + key, t)

    def SC(key):
        return b.ref("IN_SCEN", key, i)

    def FCF(key, t=0):
        return b.ref("ENG_FCF", key, t)

    rows = {}

    def add(key, label, unit, fmt, tmpl):
        row = sh.r
        sh.add_row(P + key, label, unit, fmt)
        b.mark(S, P + key, row)
        b.frow_at(S, row, tmpl)
        rows[key] = row
        return row

    pre = "[%d] " % (i + 1)
    add("vol", pre + "حجم فروش محصولات (تن)", "تن", "num",
        lambda t: "=" + b.ref("ENG_OPS", "OPS.sold_tons", t) + "*" + SC("VOL"))
    add("price", pre + "درآمد خالص هر تن", "میلیون ریال/تن", "num1",
        lambda t: "=" + b.ref("ENG_OPS", "OPS.rev_per_ton", t) + "*" + SC("PRICE"))
    add("rev", pre + "درآمد خالص", "میلیون ریال", "num",
        lambda t: "=" + R("vol", t) + "*" + R("price", t))
    add("rm", pre + "هزینه مواد اولیه", "میلیون ریال", "num",
        lambda t: "=" + b.ref("ENG_COST", "EC.rm", t) + "*" + SC("VOL") + "*" + SC("RM"))
    add("var", pre + "سایر هزینه‌های متغیر", "میلیون ریال", "num",
        lambda t: "=" + b.ref("ENG_COST", "EC.var_ex_rm", t) + "*" + SC("VOL") + "*" + SC("VAR"))
    add("fix", pre + "هزینه‌های ثابت نقدی", "میلیون ریال", "num",
        lambda t: "=" + b.ref("ENG_COST", "EC.fixed_cash", t) + "*" + SC("FIX"))
    add("dep", pre + "استهلاک", "میلیون ریال", "num",
        lambda t: "=" + b.ref("ENG_CAPEX", "CAPEX.dep_total", t) + "*" + SC("CAPEX"))
    add("ebitda", pre + "EBITDA", "میلیون ریال", "num",
        lambda t: "=" + R("rev", t) + "-" + R("rm", t) + "-" + R("var", t) + "-" + R("fix", t))
    add("ebit", pre + "EBIT", "میلیون ریال", "num",
        lambda t: "=" + R("ebitda", t) + "-" + R("dep", t))
    add("tax", pre + "مالیات", "میلیون ریال", "num",
        lambda t: "=MAX(0,%s)*%s" % (R("ebit", t), b.ref("IN_MACRO", "MAC.tax", 0)))
    add("nopat", pre + "NOPAT", "میلیون ریال", "num",
        lambda t: "=" + R("ebit", t) + "-" + R("tax", t))
    add("capex", pre + "CAPEX و جایگزینی", "میلیون ریال", "num",
        lambda t: "=(%s+%s)*%s" % (b.ref("ENG_CAPEX", "CAPEX.cap_total", t),
                                   b.ref("ENG_CAPEX", "CAPEX.repl_total", t), SC("CAPEX")))
    add("dwc", pre + "تغییرات سرمایه در گردش", "میلیون ریال", "num",
        lambda t: "=" + b.ref("ENG_WC", "WC.dwc", t) + "*" + SC("VOL") + "*" + SC("PRICE"))
    add("fcf", pre + "جریان نقد آزاد (FCFF)", "میلیون ریال", "num",
        lambda t: "=" + R("nopat", t) + "+" + R("dep", t) + "-" + R("capex", t) + "-" + R("dwc", t))
    add("fcf_rep", pre + "FCFF تکرارپذیر", "میلیون ریال", "num",
        lambda t: "=" + R("nopat", t) + "+" + R("dep", t) + "-" + R("dwc", t))

    row = sh.r
    sh.add_row(P + "cum", pre + "FCFF تجمعی", "میلیون ریال", "num")
    b.mark(S, P + "cum", row)
    rows["cum"] = row
    b.frow_at(S, row, lambda t: "=" + R("fcf", 0) if t == 0
              else "=" + R("cum", t - 1) + "+" + R("fcf", t))

    add("wacc", pre + "نرخ تنزیل (WACC)", "درصد", "pct",
        lambda t: "=" + b.ref("ENG_DCF", "DCF.wacc", t) + "+" + SC("WACC"))

    row = sh.r
    sh.add_row(P + "df", pre + "ضریب تنزیل", "ضریب", "num2")
    b.mark(S, P + "df", row)
    rows["df"] = row
    b.frow_at(S, row, lambda t: "=1" if t == 0
              else "=" + R("df", t - 1) + "/(1+" + R("wacc", t) + ")")

    add("pv", pre + "ارزش فعلی FCFF", "میلیون ریال", "num",
        lambda t: "=" + R("fcf", t) + "*" + R("df", t))

    row = sh.r
    sh.add_row(P + "pvc", pre + "ارزش فعلی تجمعی", "میلیون ریال", "num")
    b.mark(S, P + "pvc", row)
    rows["pvc"] = row
    b.frow_at(S, row, lambda t: "=" + R("pv", 0) if t == 0
              else "=" + R("pvc", t - 1) + "+" + R("pv", t))

    # ارزش پایانی: گوردون (روش ۱) یا ضریب خروج (روش ۲)
    add("tv", pre + "ارزش پایانی", "میلیون ریال", "num",
        lambda t: "=IF(%s<%s,0,IF(%s=1,MAX(0,%s)*(1+%s)/MAX(0.005,%s-%s),%s*%s))" % (
            b.ref("IN_MACRO", "MAC.t", t), b.ref("IN_MACRO", "MAC.horizon", 0),
            b.ref("IN_MACRO", "MAC.tv_method", 0), R("fcf_rep", LAST_T),
            b.ref("IN_MACRO", "MAC.g_term", 0), R("wacc", LAST_T),
            b.ref("IN_MACRO", "MAC.g_term", 0), R("ebitda", LAST_T),
            b.ref("IN_MACRO", "MAC.exit_mult", 0)))
    add("pvtv", pre + "ارزش فعلی ارزش پایانی", "میلیون ریال", "num",
        lambda t: "=" + R("tv", t) + "*" + R("df", t))
    add("fcf_tv", pre + "جریان قابل بازدهی (FCF + TV)", "میلیون ریال", "num",
        lambda t: "=" + R("fcf", t) + "+" + R("tv", t))

    # CFADS و خدمت دین بر همان مبنایِ مدلِ اصلی (ENG_FCF): مالیات روی EBT و
    # خدمت دین = اصل + سود بلندمدت + کارمزد + سود کوتاه‌مدت. ضریبِ سناریو فقط
    # روی نرخ سود اعمال می‌شود.
    kd_scen = "(%s+%s)" % (b.ref("IN_FIN", "FIN.kd", 0), SC("KD"))
    kd_act = "(%s+%s)" % (b.ref("IN_FIN", "FIN.kd", 0),
                          b.cellref(S, b.row_of(S, "KD"), COL_ACTIVE))
    add("cfads", pre + "CFADS", "میلیون ریال", "num",
        lambda t: "=%s-MAX(0,%s-(%s*%s/%s+%s+%s-%s))*%s-%s" % (
            R("ebitda", t), R("ebit", t), FCF("FCF.int_lt", t), kd_scen, kd_act,
            FCF("FCF.fees", t), FCF("FCF.int_st", t), FCF("FCF.int_inc", t),
            b.ref("IN_MACRO", "MAC.tax", 0), R("dwc", t)))
    add("debt_service", pre + "خدمت دین", "میلیون ریال", "num",
        lambda t: "=%s+%s*%s/%s+%s+%s" % (
            FCF("FCF.principal", t), FCF("FCF.int_lt", t), kd_scen, kd_act,
            FCF("FCF.fees", t), FCF("FCF.int_st", t)))
    add("dscr", pre + "DSCR", "ضریب", "num2",
        lambda t: '=IFERROR(%s/%s,"─")' % (R("cfads", t), R("debt_service", t)))

    # ---- جمع‌بندیِ سناریو (فقط در ستونِ نخست)
    rng_pv = b.rng(S, P + "pv", 0, LAST_T)
    rng_pvtv = b.rng(S, P + "pvtv", 0, LAST_T)
    rng_fcf_tv = b.rng(S, P + "fcf_tv", 0, LAST_T)
    rng_cum = b.rng(S, P + "cum", 0, LAST_T)
    rng_dscr = b.rng(S, P + "dscr", 0, LAST_T)
    rng_ds = b.rng(S, P + "debt_service", 0, LAST_T)
    rng_t = b.rng("IN_MACRO", "MAC.t", 0, LAST_T)
    last_neg = "COUNTIF(%s,\"<0\")" % rng_cum
    row = sh.r
    sh.add_row(None, pre + "ارزش فعلی خالص (NPV)", "میلیون ریال", "num")
    b.mark(S, "SB%d.npv" % (i + 1), row)
    b.fcell(S, row, spec.C0, "=SUM(%s)+SUM(%s)" % (rng_pv, rng_pvtv))
    row = sh.r
    sh.add_row(None, pre + "IRR پروژه", "درصد", "pct")
    b.mark(S, "SB%d.irr" % (i + 1), row)
    b.fcell(S, row, spec.C0, "=IFERROR(IRR(%s),NA())" % rng_fcf_tv)
    row = sh.r
    sh.add_row(None, pre + "دوره بازپرداشت (سال)", "سال", "num1")
    b.mark(S, "SB%d.pay" % (i + 1), row)
    b.fcell(S, row, spec.C0,
            '=IF(%s=0,0,IF(%s>=13,NA(),INDEX(%s,%s)+(0-INDEX(%s,%s))'
            '/(INDEX(%s,%s+1)-INDEX(%s,%s))))'
            % (last_neg, last_neg, rng_t, last_neg, rng_cum, last_neg,
               rng_cum, last_neg, rng_cum, last_neg))
    row = sh.r
    sh.add_row(None, pre + "شاخص سودآوری", "ضریب", "num2")
    b.mark(S, "SB%d.pi" % (i + 1), row)
    b.fcell(S, row, spec.C0,
            '=IFERROR((SUMIF(%s,">0")+SUM(%s))/ABS(SUMIF(%s,"<0")),NA())'
            % (rng_pv, rng_pvtv, rng_pv))
    row = sh.r
    sh.add_row(None, pre + "حداقل DSCR", "ضریب", "num2")
    b.mark(S, "SB%d.min_dscr" % (i + 1), row)
    b.fcell(S, row, spec.C0,
            "=IFERROR(IF(COUNT(%s)=0,NA(),MIN(%s)),NA())" % (rng_dscr, rng_dscr))
    row = sh.r
    sh.add_row(None, pre + "میانگین DSCR دوره بازپرداخت", "ضریب", "num2")
    b.mark(S, "SB%d.avg_dscr" % (i + 1), row)
    b.fcell(S, row, spec.C0,
            '=IFERROR(AVERAGEIF(%s,">0",%s),NA())' % (rng_ds, rng_dscr))


# ====================================================================== ENG_OPS
def eng_ops(b):
    sh = b.sheet("ENG_OPS", "موتور محاسبات عملیاتی",
                 "تبدیل ورودی‌های فنی ایستگاه‌ها به ظرفیت، جریان مواد، محصول قابل فروش "
                 "و درآمد.", "engine")
    S = "ENG_OPS"

    def R(key, t=0):
        return b.ref(S, key, t)

    def T(key, j=0):
        """پارامترِ ایستگاهی یا سطحِ کارخانه از IN_TECH (ستونِ j برای ایستگاه‌ها)"""
        return b.cellref("IN_TECH", b.row_of("IN_TECH", key), spec.C0 + j)

    def SCRAP(j):
        return T("TECH.scrap", j)

    # ------------------------------------------- الف) ظرفیت و بارگذاری
    sh.section("الف) ظرفیت و بارگذاری")
    sh.add_row("OPS.ramp", "ضریب دستیابی به ظرفیت (Ramp)", "درصد", "pct")
    b.frow(S, "OPS.ramp", lambda t: "=" + b.ref("IN_TECH", "TECH.ramp", t)
           + "*" + b.ref("IN_MACRO", "MAC.ops_flg", t))

    avail_terms = ",".join(T("TECH.avail", j) for j in range(spec.NS))
    sh.add_row("OPS.avail", "نرخ دسترسی کل خط (کمینهٔ ایستگاه‌ها)", "درصد", "pct")
    b.frow(S, "OPS.avail", lambda t: "=MIN(%s)" % avail_terms)

    sh.add_row("OPS.util", "ضریب بهره‌برداری مؤثر", "درصد", "pct")
    b.frow(S, "OPS.util", lambda t: "=" + R("OPS.ramp", t) + "*" + R("OPS.avail", t)
           + "*" + b.cellref("ENG_SCEN", b.row_of("ENG_SCEN", "VOL"), COL_ACTIVE))

    # سهمِ خوراکِ هر ایستگاه به‌ازای هر تنِ ورودی (ثابت و مستقل از ورودی → بدون دور)
    S0, S1, S2 = SCRAP(0), SCRAP(1), SCRAP(2)
    S3, S4, S5 = SCRAP(3), SCRAP(4), SCRAP(5)
    b2 = "(1-%s)*(1-%s)" % (S0, S1)                       # پس از دو ایستگاه نخست
    ps = "((%s)*%s)" % (b2, T("TECH.sh_paste"))           # خمیرِ تفکیک‌شده
    ds = "(%s*%s*(1-%s))" % (ps, T("TECH.desulf"), S2)    # خمیرِ سولفات‌زدایی‌شده
    mt = "((%s)*%s)" % (b2, T("TECH.sh_metal"))           # فلزِ تفکیک‌شده
    fn = "(%s*%s+%s*%s)" % (ds, T("TECH.pb_paste"), mt, T("TECH.pb_metal"))
    crude = "(%s*%s*(1-%s))" % (fn, T("TECH.smelt"), S3)
    refined = "(%s*%s*(1-%s))" % (crude, T("TECH.refine"), S4)
    pp_ = "((%s)*%s)" % (b2, T("TECH.sh_pp"))
    SPLITS = ["=1",
              "=(1-%s)" % S0,
              "=" + ps,
              "=(%s+%s)" % (ds, mt),
              "=" + crude,
              "=(%s+%s*%s+%s)" % (refined, pp_, T("TECH.pp_rec"), T("TECH.na2so4"))]
    for j in range(spec.NS):
        sh.add_row("OPS.split.%d" % j,
                   "سهم خوراک ایستگاه %d (به‌ازای هر تن ورودی)" % (j + 1), "تن/تن", "num2")
        b.frow(S, "OPS.split.%d" % j, lambda t, j=j: SPLITS[j])

    sh.add_row("OPS.intake_cap", "بیشینهٔ ورودی قابل پردازش (محدودیت گلوگاه)", "تن", "num")
    b.frow(S, "OPS.intake_cap", lambda t: "=MIN(%s)" % ",".join(
        "IFERROR(%s/%s,1E+15)" % (T("TECH.cap_yr", j), R("OPS.split.%d" % j, t))
        for j in range(spec.NS)))

    sh.add_row("OPS.intake", "ورودی واقعی باتری ضایعاتی", "تن", "num")
    b.frow(S, "OPS.intake", lambda t: "=MIN(%s*%s,%s)"
           % (T("TECH.cap_in"), R("OPS.util", t), R("OPS.intake_cap", t)))

    sh.add_row("OPS.hours_nom", "ساعات اسمی کارکرد در سال", "ساعت", "num")
    b.frow(S, "OPS.hours_nom", lambda t: "=%s*%s*%s"
           % (T("TECH.days"), T("TECH.shifts"), T("TECH.hshift")))
    sh.add_row("OPS.hours", "ساعات مؤثر کارکرد", "ساعت", "num")
    b.frow(S, "OPS.hours", lambda t: "=MAX(0,(%s-%s-%s))*%s"
           % (R("OPS.hours_nom", t), T("TECH.dt_plan"), T("TECH.dt_unplan"),
              R("OPS.util", t)))
    sh.add_row("OPS.caputil", "نرخ استفاده از ظرفیت اسمی", "درصد", "pct")
    b.frow(S, "OPS.caputil", lambda t: "=IFERROR(%s/%s,0)"
           % (R("OPS.intake", t), T("TECH.cap_in")))

    # ------------------------------------------- ب) جریان مواد و تراز جرم
    sh.gap()
    sh.section("ب) جریان مواد و تراز جرم")
    sh.add_row("OPS.s1out", "خروجی ایستگاه ۱ (پس از تخلیه)", "تن", "num")
    b.frow(S, "OPS.s1out", lambda t: "=%s*(1-%s)" % (R("OPS.intake", t), SCRAP(0)))
    sh.add_row("OPS.s2out", "خروجی ایستگاه ۲ (پس از خردایش)", "تن", "num")
    b.frow(S, "OPS.s2out", lambda t: "=%s*(1-%s)" % (R("OPS.s1out", t), SCRAP(1)))
    for key, lab, src in [("OPS.paste", "خمیر سرب (جریان تفکیک‌شده)", "TECH.sh_paste"),
                          ("OPS.metal", "قطعات فلزی سربی (جریان تفکیک‌شده)", "TECH.sh_metal"),
                          ("OPS.pp", "پلی‌پروپیلن (جریان تفکیک‌شده)", "TECH.sh_pp"),
                          ("OPS.res", "جداکننده و سایر پسماند (جریان تفکیک‌شده)", "TECH.sh_res")]:
        sh.add_row(key, lab, "تن", "num")
        b.frow(S, key, lambda t, src=src: "=%s*%s" % (R("OPS.s2out", t), T(src)))
    sh.add_row("OPS.elec", "الکترولیت تخلیه‌شده", "تن", "num")
    b.frow(S, "OPS.elec", lambda t: "=%s*%s" % (R("OPS.intake", t), T("TECH.sh_elec")))
    sh.add_row("OPS.paste_ds", "خمیر سولفات‌زدایی‌شده", "تن", "num")
    b.frow(S, "OPS.paste_ds", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.paste", t), T("TECH.desulf"), SCRAP(2)))
    sh.add_row("OPS.furnace_pb", "سرب ورودی به کوره", "تن", "num")
    b.frow(S, "OPS.furnace_pb", lambda t: "=%s*%s+%s*%s"
           % (R("OPS.paste_ds", t), T("TECH.pb_paste"), R("OPS.metal", t), T("TECH.pb_metal")))
    sh.add_row("OPS.crude", "سرب خام خروجی کوره", "تن", "num")
    b.frow(S, "OPS.crude", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.furnace_pb", t), T("TECH.smelt"), SCRAP(3)))
    sh.add_row("OPS.refined", "سرب تصفیه‌شده", "تن", "num")
    b.frow(S, "OPS.refined", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.crude", t), T("TECH.refine"), SCRAP(4)))

    # ------------------------------------------- ج) محصولات قابل فروش
    sh.gap()
    sh.section("ج) محصولات قابل فروش")
    sh.add_row("OPS.soft", "شمش سرب خالص (تولید)", "تن", "num")
    b.frow(S, "OPS.soft", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.refined", t), T("TECH.sp_soft"), SCRAP(5)))
    sh.add_row("OPS.alloy", "شمش آلیاژ سرب (تولید)", "تن", "num")
    b.frow(S, "OPS.alloy", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.refined", t), T("TECH.sp_alloy"), SCRAP(5)))
    sh.add_row("OPS.ppg", "گرانول پلی‌پروپیلن (تولید)", "تن", "num")
    b.frow(S, "OPS.ppg", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.pp", t), T("TECH.pp_rec"), SCRAP(5)))
    sh.add_row("OPS.na2so4", "سولفات سدیم (تولید)", "تن", "num")
    b.frow(S, "OPS.na2so4", lambda t: "=%s*%s*(1-%s)"
           % (R("OPS.intake", t), T("TECH.na2so4"), SCRAP(5)))
    sh.add_row("OPS.prod_total", "جمع محصولات تولیدی", "تن", "num")
    b.frow(S, "OPS.prod_total", lambda t: "=%s" % "+".join(
        R(k, t) for k in ("OPS.soft", "OPS.alloy", "OPS.ppg", "OPS.na2so4")))
    sh.add_row("OPS.waste", "ضایعات و پسماند (کنترل تراز)", "تن", "num")
    b.frow(S, "OPS.waste", lambda t: "=%s-%s" % (R("OPS.intake", t), R("OPS.prod_total", t)))

    # ------------------------------------------- د) ورودی و بارگذاری ایستگاه‌ها
    sh.gap()
    sh.section("د) ورودی و بارگذاری ایستگاه‌ها")
    FEED = [("OPS.feed.0", lambda t: "=" + R("OPS.intake", t)),
            ("OPS.feed.1", lambda t: "=" + R("OPS.s1out", t)),
            ("OPS.feed.2", lambda t: "=" + R("OPS.paste", t)),
            ("OPS.feed.3", lambda t: "=%s+%s" % (R("OPS.paste_ds", t), R("OPS.metal", t))),
            ("OPS.feed.4", lambda t: "=" + R("OPS.crude", t)),
            ("OPS.feed.5", lambda t: "=%s+%s/(1-%s)+%s/(1-%s)"
             % (R("OPS.refined", t), R("OPS.ppg", t), SCRAP(5),
                R("OPS.na2so4", t), SCRAP(5)))]
    for key, tmpl in FEED:
        sh.add_row(key, "ورودی ایستگاه %d" % (int(key[-1]) + 1), "تن", "num")
        b.frow(S, key, tmpl)

    feed_rows = {}
    for j in range(spec.NS):
        key = "S%d" % (j + 1)
        row = sh.r
        sh.add_row(key, "بارگذاری ایستگاه %d" % (j + 1), "درصد", "pct")
        feed_rows[j] = row
        b.frow(S, key, lambda t, j=j: "=IFERROR(%s/%s,0)"
               % (R("OPS.feed.%d" % j, t), T("TECH.cap_yr", j)))

    r0, r1 = feed_rows[0], feed_rows[spec.NS - 1]
    sh.add_row("OPS.bottleneck", "بیشترین بارگذاری (گلوگاه)", "درصد", "pct")
    b.frow(S, "OPS.bottleneck", lambda t: "=MAX(%s)"
           % ",".join(b.cellref(S, feed_rows[j], spec.C0 + t) for j in range(spec.NS)))
    sh.add_row("OPS.bottleneck_st", "شماره ایستگاه گلوگاه", "شماره", "num")
    b.frow(S, "OPS.bottleneck_st", lambda t:
           "=IFERROR(INDEX(%s,MATCH(%s,INDEX(%s,%d,0),0),1),0)"
           % (b.rng_at(S, r0, 2, r1, 2), R("OPS.bottleneck", t),
              b.rng_at(S, r0, spec.C0, r1, spec.C0 + LAST_T), t + 1))

    # ------------------------------------------- ه‍) مصارف انرژی و آب
    sh.gap()
    sh.section("ه‍) مصارف انرژی و آب")
    for key, lab, src in [("OPS.kwh", "برق مصرفی کل", "TECH.kwh"),
                          ("OPS.water", "آب مصرفی کل", "TECH.water")]:
        sh.add_row(key, lab, "kWh" if key.endswith("kwh") else "m3", "num")
        b.frow(S, key, lambda t, src=src: "=%s" % "+".join(
            "(%s*%s)" % (R("OPS.feed.%d" % j, t), T(src, j)) for j in range(spec.NS)))

    # ------------------------------------------- و) قیمت‌گذاری و درآمد
    sh.gap()
    sh.section("و) قیمت‌گذاری و درآمد")

    def SP(key, j):
        return b.cellref("IN_SALES", b.row_of("IN_SALES", key), spec.C0 + j)

    active_price = b.cellref("ENG_SCEN", b.row_of("ENG_SCEN", "PRICE"), COL_ACTIVE)
    price_rows = {}
    for j in range(spec.NP):
        row = sh.r
        key = "OPS.price" if j == 0 else None
        sh.add_row(key, "قیمت فروش محصول %d" % (j + 1), "میلیون ریال/تن", "num1")
        if key:
            b.mark(S, "OPS.price", row)
        price_rows[j] = row
        b.frow_at(S, row, lambda t, j=j: "=IF(%s=1,%s*%s/1000000*%s,%s*POWER(1+%s,%s)"
                                         "*(1+%s*(%s/%s-1)))*%s"
                  % (SP("SAL.lme", j), b.ref("ENG_SCEN", "SCEN.lme", t),
                     b.ref("ENG_SCEN", "SCEN.fx", t), SP("SAL.lmef", j),
                     SP("SAL.price", j), SP("SAL.esc", j),
                     b.ref("IN_MACRO", "MAC.t", t), SP("SAL.fxsh", j),
                     b.ref("ENG_SCEN", "SCEN.fx", t), b.ref("ENG_SCEN", "SCEN.fx", 0),
                     active_price))

    PROD = ["OPS.soft", "OPS.alloy", "OPS.ppg", "OPS.na2so4"]
    vol_rows, grev_rows = {}, {}
    for j in range(spec.NP):
        row = sh.r
        key = "OPS.vol" if j == 0 else None
        sh.add_row(key, "مقدار فروش محصول %d" % (j + 1), "تن", "num")
        if key:
            b.mark(S, "OPS.vol", row)
        vol_rows[j] = row
        b.frow_at(S, row, lambda t, j=j: "=%s*%s" % (b.ref(S, PROD[j], t), SP("SAL.real", j)))

    for j in range(spec.NP):
        row = sh.r
        key = "OPS.grev" if j == 0 else None
        sh.add_row(key, "درآمد ناخالص محصول %d" % (j + 1), "میلیون ریال", "num")
        if key:
            b.mark(S, "OPS.grev", row)
        grev_rows[j] = row
        b.frow_at(S, row, lambda t, j=j: "=%s*%s"
                  % (b.cellref(S, vol_rows[j], spec.C0 + t),
                     b.cellref(S, price_rows[j], spec.C0 + t)))

    def grev(j, t):
        return b.cellref(S, grev_rows[j], spec.C0 + t)

    def vol(j, t):
        return b.cellref(S, vol_rows[j], spec.C0 + t)

    sh.add_row("OPS.grev_t", "جمع درآمد ناخالص", "میلیون ریال", "num")
    b.frow(S, "OPS.grev_t", lambda t: "=%s" % "+".join(grev(j, t) for j in range(spec.NP)))
    sh.add_row("OPS.disc", "تخفیفات و برگشت از فروش", "میلیون ریال", "num")
    b.frow(S, "OPS.disc", lambda t: "=%s" % "+".join(
        "(%s*%s)" % (grev(j, t), SP("SAL.disc", j)) for j in range(spec.NP)))
    sh.add_row("OPS.rev", "درآمد خالص", "میلیون ریال", "num")
    b.frow(S, "OPS.rev", lambda t: "=%s-%s" % (R("OPS.grev_t", t), R("OPS.disc", t)))
    sh.add_row("OPS.sold_tons", "جمع تن فروخته‌شده", "تن", "num")
    b.frow(S, "OPS.sold_tons", lambda t: "=%s" % "+".join(vol(j, t) for j in range(spec.NP)))
    sh.add_row("OPS.rev_per_ton", "درآمد خالص هر تن محصول", "میلیون ریال/تن", "num1")
    b.frow(S, "OPS.rev_per_ton", lambda t: "=IFERROR(%s/%s,0)"
           % (R("OPS.rev", t), R("OPS.sold_tons", t)))
    sh.add_row("OPS.fxsh_rev", "سهم ارزی درآمد (وزنی)", "درصد", "pct")
    b.frow(S, "OPS.fxsh_rev", lambda t: "=IFERROR((%s)/(%s),0)"
           % ("+".join("(%s*%s)" % (grev(j, t), SP("SAL.fxsh", j)) for j in range(spec.NP)),
              R("OPS.grev_t", t)))
    return sh


def build(b):
    eng_scen(b)
    eng_ops(b)
