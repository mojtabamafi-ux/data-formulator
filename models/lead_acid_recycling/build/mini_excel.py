# -*- coding: utf-8 -*-
"""ارزیابِ کمینه برای فرمول‌های اکسل.

فقط توابعِ به‌کاررفته در این مدل (به‌همراه چند تابعِ پرکاربردِ دیگر) پیاده‌سازی شده‌اند:
IF, IFERROR, AND, OR, NOT, SUM, MIN, MAX, AVERAGE, COUNT, SUMIF, COUNTIF, AVERAGEIF,
INDEX, MATCH, CHOOSE, IRR, MIRR, NPV, PMT, IPMT, PPMT, ABS, POWER, MOD, ROUND, SQRT,
VALUE, MID, FIND, LEN, LEFT, RIGHT, NA, HYPERLINK, SUMPRODUCT, ISERROR, ISNUMBER.

نکتهٔ مهم: نتیجهٔ INDEX/MATCH یک «ارجاع» است؛ هر جا مقدارِ اسکالر لازم باشد
(عملگرها و آرگومان‌های عددی توابع) این ارجاع به مقدارِ سلول تبدیل می‌شود.
"""

import math
import re

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# --------------------------------------------------------------------- خطاها
class XLError(Exception):
    def __init__(self, code, detail=""):
        super().__init__(code if not detail else "%s | %s" % (code, detail))
        self.code = code


# --------------------------------------------------------------------- ارجاع
class Ref(object):
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2=None, c2=None):
        self.sheet = sheet
        self.r1, self.c1 = int(r1), int(c1)
        self.r2, self.c2 = int(r1 if r2 is None else r2), int(c1 if c2 is None else c2)

    @property
    def single(self):
        return self.r1 == self.r2 and self.c1 == self.c2

    def __repr__(self):
        return "%s!%s%d%s%s" % (
            self.sheet, get_column_letter(self.c1), self.r1,
            "" if self.single else ":",
            "" if self.single else "%s%d" % (get_column_letter(self.c2), self.r2))


# ------------------------------------------------------------------ توکن‌ساز
TOKEN_RE = re.compile(r"""
      (?P<ws>\s+)
    | (?P<str>"(?:[^"]|"")*")
    | (?P<ref>(?:(?:'(?P<sheet1>[^']+)'|(?P<sheet2>[A-Za-z_][A-Za-z0-9_.]*))!)?
              (?P<col>\$?[A-Za-z]{1,3}\$?)(?P<row>\$?\d+))
    | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
    | (?P<func>[A-Za-z][A-Za-z0-9_.]*(?=\s*\())
    | (?P<true>\bTRUE\b)|(?P<false>\bFALSE\b)
    | (?P<lp>\()|(?P<rp>\))|(?P<comma>,)|(?P<colon>:)|(?P<pct>%)
    | (?P<cmp><>|>=|<=|=|<|>)
    | (?P<op>[+\-*/^&])
""", re.VERBOSE)


def tokenize(s):
    toks, pos, n = [], 0, len(s)
    while pos < n:
        m = TOKEN_RE.match(s, pos)
        if not m:
            raise XLError("#VALUE!", "نویسهٔ نشناخته در «%s»" % s[pos:pos + 20])
        pos = m.end()
        kind = m.lastgroup
        if kind == "ws":
            continue
        toks.append((kind, m))
    return toks


# ----------------------------------------------------------------- تبدیل‌ها
def _num(v):
    """تبدیل به عدد؛ در صورت ناممکن بودن خطای #VALUE!"""
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", ""))
        except ValueError:
            raise XLError("#VALUE!")
    if v is None:
        return 0.0
    raise XLError("#VALUE!")


def _flatten(seq):
    out = []
    for v in seq:
        if isinstance(v, list):
            out.extend(_flatten(v))
        else:
            out.append(v)
    return out


def _isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _criteria(crit):
    """ساختِ تابعِ شرط برای COUNTIF/SUMIF/AVERAGEIF"""
    crit = str(crit).strip()
    for op in (">=", "<=", "<>", ">", "<", "="):
        if crit.startswith(op):
            try:
                val = float(crit[len(op):])
            except ValueError:
                return lambda v: isinstance(v, str) and str(v) == crit[len(op):]
            return {"<>": lambda v: _isnum(v) and v != val,
                    ">=": lambda v: _isnum(v) and v >= val,
                    "<=": lambda v: _isnum(v) and v <= val,
                    ">": lambda v: _isnum(v) and v > val,
                    "<": lambda v: _isnum(v) and v < val,
                    "=": lambda v: _isnum(v) and v == val}[op]
    try:
        val = float(crit)
    except ValueError:
        return lambda v: isinstance(v, str) and str(v) == crit
    return lambda v: _isnum(v) and v == val


# ------------------------------------------------------------- توابعِ مالی
def _irr(vals, guess=0.1):
    vals = [float(v) for v in vals if _isnum(v)]
    if len(vals) < 2 or not (min(vals) < 0 < max(vals)):
        raise XLError("#NUM!")
    lo, hi = -0.999999999, 10.0

    def npv_at(r):
        return sum(v / (1.0 + r) ** i for i, v in enumerate(vals))

    if npv_at(lo) * npv_at(hi) > 0:
        raise XLError("#NUM!")
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if npv_at(mid) > 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < 1e-12:
            break
    return (lo + hi) / 2.0


def _mirr(vals, fin, reinv):
    vals = [float(v) for v in vals if _isnum(v)]
    if not vals:
        raise XLError("#DIV/0!")
    n = len(vals) - 1
    neg = sum(v / (1.0 + fin) ** i for i, v in enumerate(vals) if v < 0)
    pos = sum(v * (1.0 + reinv) ** (n - i) for i, v in enumerate(vals) if v > 0)
    if neg == 0 or pos == 0:
        raise XLError("#DIV/0!")
    return (pos / -neg) ** (1.0 / n) - 1.0


def _pmt(rate, nper, pv, fv=0.0, typ=0):
    if nper == 0:
        raise XLError("#DIV/0!")
    if rate == 0:
        return -(pv + fv) / nper
    f = (1.0 + rate) ** nper
    return -(pv * f + fv) / (((f - 1.0) / rate) * (1.0 + rate * typ)) / (1.0 + rate * typ) \
        if typ else -(pv * f + fv) / ((f - 1.0) / rate)


def _ipmt(rate, per, nper, pv, fv=0.0, typ=0):
    if per < 1 or per > nper:
        raise XLError("#NUM!")
    pmt = _pmt(rate, nper, pv, fv, typ)
    if rate == 0:
        return 0.0
    if per == 1:
        bal = pv if typ == 0 else pv - pmt
    else:
        f = (1.0 + rate) ** (per - 2)
        bal = pv * f + pmt * (f - 1.0) / rate * (1.0 + rate * typ)
        if typ:
            bal -= pmt
    return -bal * rate


def _ppmt(rate, per, nper, pv, fv=0.0, typ=0):
    return _pmt(rate, nper, pv, fv, typ) - _ipmt(rate, per, nper, pv, fv, typ)


# ------------------------------------------------------------------ ارزیاب
class Evaluator(object):
    def __init__(self, source):
        if isinstance(source, str):
            self.wb = load_workbook(source, data_only=False)
        else:
            self.wb = source
        self.cells = {}
        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        self.cells[(ws.title, c.row, c.column)] = c.value
        self.cache = {}
        self.busy = set()
        self.path = []
        self.max_depth = 0

    # ------------------------------------------------------- سلول‌ها
    def set_cell(self, sheet, row, col, value):
        """تغییرِ مقدارِ یک سلول (برای آزمون‌های آبشاری)"""
        self.cells[(sheet, row, col)] = value
        self.cache.clear()

    def get(self, sheet, row, col):
        key = (sheet, row, col)
        if key in self.cache:
            return self.cache[key]
        if key in self.busy:
            raise XLError("دور ارجاعی در %s!%s%d" % (sheet, get_column_letter(col), row))
        v = self.cells.get(key)
        if isinstance(v, str) and v.startswith("="):
            self.busy.add(key)
            self.path.append("%s!%s%d" % (sheet, get_column_letter(col), row))
            if len(self.path) > self.max_depth:
                self.max_depth = len(self.path)
            try:
                res = self.eval_formula(sheet, v[1:])
                if isinstance(res, Ref):
                    res = self.get(res.sheet, res.r1, res.c1) if res.single \
                        else XLError("#VALUE!")
                if isinstance(res, XLError):
                    raise res
            except XLError as e:
                raise XLError(e.code, "مسیر: " + " ← ".join(self.path[-6:]))
            finally:
                self.busy.discard(key)
                self.path.pop()
            self.cache[key] = res
            return res
        self.cache[key] = v
        return v

    def eval_formula(self, sheet, text):
        self.sheet = sheet
        self.toks = tokenize(text)
        self.pos = 0
        node = self.parse_expr()
        if self.pos != len(self.toks):
            raise XLError("#VALUE!", "عبارتِ اضافی در انتهای فرمول")
        return self.value(node)

    # ------------------------------------------------------- تجزیه‌گر
    def peek(self):
        if self.pos < len(self.toks):
            return self.toks[self.pos][0], self.toks[self.pos][1]
        return None, None

    def take(self, kind):
        k, m = self.peek()
        if k != kind:
            raise XLError("#VALUE!", "انتظار %s در «%s»" % (kind, m.group(0) if m else "پایان"))
        self.pos += 1
        return k, m

    def parse_expr(self):
        return self.parse_cmp()

    def parse_cmp(self):
        left = self.parse_concat()
        k, m = self.peek()
        if k == "cmp":
            self.pos += 1
            op = m.group(0)
            return ("cmp", op, left, self.parse_concat())
        return left

    def parse_concat(self):
        node = self.parse_add()
        while self.peek()[0] == "op" and self.peek()[1].group(0) == "&":
            self.pos += 1
            node = ("concat", node, self.parse_add())
        return node

    def parse_add(self):
        node = self.parse_mul()
        while self.peek()[0] == "op" and self.peek()[1].group(0) in "+-":
            op = self.peek()[1].group(0)
            self.pos += 1
            node = ("arith", op, node, self.parse_mul())
        return node

    def parse_mul(self):
        node = self.parse_power()
        while self.peek()[0] == "op" and self.peek()[1].group(0) in "*/":
            op = self.peek()[1].group(0)
            self.pos += 1
            node = ("arith", op, node, self.parse_power())
        return node

    def parse_power(self):
        node = self.parse_unary()
        if self.peek()[0] == "op" and self.peek()[1].group(0) == "^":
            self.pos += 1
            node = ("arith", "^", node, self.parse_power())
        return node

    def parse_unary(self):
        k, m = self.peek()
        if k == "op" and m.group(0) in "+-":
            self.pos += 1
            node = self.parse_unary()
            return node if m.group(0) == "+" else ("neg", node)
        return self.parse_primary()

    def parse_primary(self):
        node = self.parse_atom()
        while self.peek()[0] == "pct":
            self.pos += 1
            node = ("pct", node)
        if self.peek()[0] == "colon":
            self.pos += 1
            node2 = self.parse_atom()
            if node[0] == "ref" and node2[0] == "ref":
                a, b = node[1], node2[1]
                sheet = a.sheet
                if sheet == self.sheet and b.sheet != self.sheet:
                    sheet = b.sheet
                return ("ref", Ref(sheet, min(a.r1, b.r1), min(a.c1, b.c1),
                                   max(a.r2, b.r2), max(a.c2, b.c2)))
            return ("range", node, node2)
        return node

    def parse_atom(self):
        k, m = self.peek()
        if k == "num":
            self.pos += 1
            return ("num", float(m.group(0)))
        if k == "str":
            self.pos += 1
            return ("str", m.group(0)[1:-1].replace('""', '"'))
        if k == "true":
            self.pos += 1
            return ("num", 1.0)
        if k == "false":
            self.pos += 1
            return ("num", 0.0)
        if k == "lp":
            self.pos += 1
            v = self.parse_expr()
            self.take("rp")
            return v
        if k == "func":
            name = m.group(0).upper()
            self.pos += 1
            self.take("lp")
            args = []
            if self.peek()[0] != "rp":
                args.append(self.parse_expr())
                while self.peek()[0] == "comma":
                    self.pos += 1
                    args.append(self.parse_expr())
            self.take("rp")
            return ("call", name, args)
        if k == "ref":
            self.pos += 1
            return ("ref", self.make_ref(m))
        raise XLError("#VALUE!", "عبارت نامعتبر در «%s»" % (m.group(0) if m else "؟"))

    def make_ref(self, m):
        sheet = m.group("sheet1") or m.group("sheet2") or self.sheet
        col_txt = m.group("col").replace("$", "")
        col = 0
        for ch in col_txt:
            col = col * 26 + (ord(ch.upper()) - 64)
        row = int(m.group("row").replace("$", ""))
        return Ref(sheet, row, col)

    # --------------------------------------------------------- ارزیابی
    def deref(self, v):
        if isinstance(v, Ref):
            if not v.single:
                return XLError("#VALUE!")
            v = self.get(v.sheet, v.r1, v.c1)
        if isinstance(v, XLError):
            raise v
        return v

    def v1(self, node):
        return self.deref(self.value(node))

    def value(self, node):
        kind = node[0]
        if kind == "num" or kind == "str":
            return node[1]
        if kind == "ref":
            ref = node[1]
            return ref if not ref.single else self.get(ref.sheet, ref.r1, ref.c1)
        if kind == "range":
            a, b = self.value(node[1]), self.value(node[2])
            if not isinstance(a, Ref) or not isinstance(b, Ref):
                raise XLError("#REF!")
            return Ref(a.sheet or b.sheet, min(a.r1, b.r1), min(a.c1, b.c1),
                       max(a.r2, b.r2), max(a.c2, b.c2))
        if kind == "neg":
            return -_num(self.v1(node[1]))
        if kind == "pct":
            return _num(self.v1(node[1])) / 100.0
        if kind == "arith":
            op = node[1]
            a = _num(self.v1(node[2]))
            b = _num(self.v1(node[3]))
            if op == "+":
                return a + b
            if op == "-":
                return a - b
            if op == "*":
                return a * b
            if op == "/":
                if b == 0:
                    raise XLError("#DIV/0!")
                return a / b
            if op == "^":
                try:
                    return a ** b
                except (OverflowError, ValueError):
                    raise XLError("#NUM!")
        if kind == "concat":
            return "%s%s" % (self.text(node[1]), self.text(node[2]))
        if kind == "cmp":
            op = node[1]
            a, b = self.v1(node[2]), self.v1(node[3])
            if isinstance(a, str) or isinstance(b, str):
                a, b = str(a), str(b)
            return {">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b,
                    "=": a == b, "<>": a != b}[op]
        if kind == "call":
            return self.call(node[1], node[2])
        raise XLError("#VALUE!", "گرهٔ ناشناخته %s" % kind)

    def text(self, node):
        v = self.value(node)
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        return str(v)

    def arg(self, node):
        """مقدارِ آرگومان؛ بازه به صورت آرایهٔ دوبعدی برگردانده می‌شود."""
        v = self.value(node)
        if isinstance(v, Ref):
            return self.ref_values(v)
        return v

    def arr(self, node):
        v = self.arg(node)
        return _flatten(v) if isinstance(v, list) else [v]

    def ref_values(self, ref):
        out = []
        for r in range(ref.r1, ref.r2 + 1):
            row = []
            for c in range(ref.c1, ref.c2 + 1):
                v = self.get(ref.sheet, r, c)
                row.append(v)
            out.append(row)
        return out

    # ---------------------------------------------------------- توابع
    def call(self, name, args):
        if name == "IF":
            cond = self.value(args[0])
            cond = bool(cond) if not isinstance(cond, XLError) else False
            if cond:
                return self.value(args[1])
            return self.value(args[2]) if len(args) > 2 else False
        if name == "IFERROR":
            try:
                return self.value(args[0])
            except XLError:
                return self.value(args[1]) if len(args) > 1 else ""
        if name in ("ISERROR", "ISERR"):
            try:
                self.value(args[0])
                return False
            except XLError:
                return True
        if name == "ISNUMBER":
            try:
                self.value(args[0])
                return _isnum(self.v1(args[0]))
            except XLError:
                return False
        if name in ("AND", "OR"):
            vals = self.arr(args[0]) if len(args) == 1 else [self.value(a) for a in args]
            vals = [bool(v) for v in vals if not isinstance(v, str)]
            return all(vals) if name == "AND" else any(vals)
        if name == "NOT":
            return not bool(self.value(args[0]))
        if name == "SUM":
            tot = 0.0
            for a in args:
                for v in self.arr(a):
                    if _isnum(v):
                        tot += v
            return tot
        if name == "SUMPRODUCT":
            cols = [self.arr(a) for a in args]
            n = min(len(c) for c in cols) if cols else 0
            return sum(math.prod([_num(c[i]) for c in cols]) for i in range(n))
        if name in ("MIN", "MAX"):
            vals = [v for a in args for v in self.arr(a) if _isnum(v)]
            if not vals:
                return 0.0
            return min(vals) if name == "MIN" else max(vals)
        if name == "AVERAGE":
            vals = [v for a in args for v in self.arr(a) if _isnum(v)]
            if not vals:
                raise XLError("#DIV/0!")
            return sum(vals) / len(vals)
        if name == "COUNT":
            return sum(1 for a in args for v in self.arr(a) if _isnum(v))
        if name == "COUNTA":
            return sum(1 for a in args for v in self.arr(a)
                       if v is not None and v != "")
        if name == "ABS":
            return abs(_num(self.v1(args[0])))
        if name == "SQRT":
            v = _num(self.v1(args[0]))
            if v < 0:
                raise XLError("#NUM!")
            return math.sqrt(v)
        if name == "ROUND":
            return float(round(_num(self.v1(args[0])), int(_num(self.v1(args[1])))))
        if name == "POWER":
            return _num(self.v1(args[0])) ** _num(self.v1(args[1]))
        if name == "MOD":
            a, b = _num(self.v1(args[0])), _num(self.v1(args[1]))
            if b == 0:
                raise XLError("#DIV/0!")
            return a - b * math.floor(a / b)
        if name == "COUNTIF":
            return sum(1 for v in self.arr(args[0])
                       if _criteria(self.v1(args[1]))(v))
        if name == "SUMIF":
            rng = self.arr(args[0])
            fn = _criteria(self.v1(args[1]))
            other = self.arr(args[2]) if len(args) > 2 else rng
            return sum(v for v, flag in zip(other, rng) if flag and _isnum(v))
        if name == "AVERAGEIF":
            rng = self.arr(args[0])
            fn = _criteria(self.v1(args[1]))
            other = self.arr(args[2]) if len(args) > 2 else rng
            sel = [v for v, flag in zip(other, rng) if flag and _isnum(v)]
            if not sel:
                raise XLError("#DIV/0!")
            return sum(sel) / len(sel)
        if name == "INDEX":
            arg0 = self.value(args[0])
            if not isinstance(arg0, Ref):
                raise XLError("#REF!")
            row_n = int(_num(self.v1(args[1]))) if len(args) > 1 else 1
            col_n = int(_num(self.v1(args[2]))) if len(args) > 2 else None
            if col_n is None:
                if arg0.r1 == arg0.r2:
                    col_n, row_n = row_n, 1
                else:
                    row_n, col_n = row_n, 1
            r1 = arg0.r1 + (row_n - 1 if row_n else 0)
            r2 = (arg0.r1 + row_n - 1) if row_n else arg0.r2
            c1 = arg0.c1 + (col_n - 1 if col_n else 0)
            c2 = (arg0.c1 + col_n - 1) if col_n else arg0.c2
            return Ref(arg0.sheet, r1, c1, r2, c2)
        if name == "MATCH":
            target = self.v1(args[0])
            arg0 = self.value(args[1])
            vals = self.arr(args[1]) if isinstance(arg0, Ref) else [arg0]
            for i, v in enumerate(vals):
                if isinstance(v, str) or isinstance(target, str):
                    if str(v) == str(target):
                        return float(i + 1)
                elif _isnum(v) and _isnum(target) and abs(v - target) < 1e-9:
                    return float(i + 1)
            raise XLError("#N/A")
        if name == "CHOOSE":
            idx = int(_num(self.v1(args[0])))
            if idx < 1 or idx > len(args) - 1:
                raise XLError("#VALUE!")
            return self.value(args[idx])
        if name == "IRR":
            return _irr(self.arr(args[0]))
        if name == "MIRR":
            return _mirr(self.arr(args[0]), _num(self.v1(args[1])),
                         _num(self.v1(args[2])))
        if name == "NPV":
            rate = _num(self.v1(args[0]))
            vals = [v for a in args[1:] for v in self.arr(a)]
            return sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(vals)
                       if _isnum(v))
        if name == "PMT":
            return _pmt(_num(self.v1(args[0])), int(_num(self.v1(args[1]))),
                        _num(self.v1(args[2])))
        if name == "IPMT":
            return _ipmt(_num(self.v1(args[0])), int(_num(self.v1(args[1]))),
                         int(_num(self.v1(args[2]))), _num(self.v1(args[3])))
        if name == "PPMT":
            r = _num(self.v1(args[0]))
            per = int(_num(self.v1(args[1])))
            n = int(_num(self.v1(args[2])))
            pv = _num(self.v1(args[3]))
            return _ppmt(r, per, n, pv)
        if name == "VALUE":
            try:
                return float(str(self.v1(args[0])).strip())
            except ValueError:
                raise XLError("#VALUE!")
        if name == "MID":
            s = str(self.v1(args[0]))
            start = int(_num(self.v1(args[1])))
            ln = int(_num(self.v1(args[2])))
            return s[start - 1:start - 1 + ln]
        if name in ("LEFT", "RIGHT"):
            s = str(self.v1(args[0]))
            n = int(_num(self.v1(args[1]))) if len(args) > 1 else 1
            return s[:n] if name == "LEFT" else s[-n:]
        if name == "LEN":
            return float(len(str(self.v1(args[0]))))
        if name == "FIND":
            needle = str(self.v1(args[0]))
            hay = str(self.v1(args[1]))
            start = int(_num(self.v1(args[2]))) if len(args) > 2 else 1
            idx = hay.find(needle, start - 1)
            if idx < 0:
                raise XLError("#VALUE!")
            return float(idx + 1)
        if name == "NA":
            raise XLError("#N/A")
        if name == "HYPERLINK":
            return str(self.v1(args[1])) if len(args) > 1 else str(self.v1(args[0]))
        raise XLError("#NAME?", "تابع پیاده‌سازی‌نشده: %s" % name)
