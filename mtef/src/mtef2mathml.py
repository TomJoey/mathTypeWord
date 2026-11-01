"""Convert a parsed MTEF tree (see mtef.py) to MathML (xml.etree elements)."""
import xml.etree.ElementTree as ET
from mtef import Char, Tmpl, Pile, Matrix, Line, Skip, Embell

MATHML_NS = "http://www.w3.org/1998/Math/MathML"

# typeface index (typeface byte - 128)
TF_TEXT, TF_FUNCTION, TF_VARIABLE, TF_LCGREEK, TF_UCGREEK, TF_SYMBOL, TF_VECTOR, TF_NUMBER = 1, 2, 3, 4, 5, 6, 7, 8
TF_USER1, TF_USER2, TF_MTEXTRA, TF_TEXT_FE = 9, 10, 11, 12

# ---- template selector tables -------------------------------------------------
# (name, class) per selector; v3 = "Table A", v5 = "Table B" (see rtf2latex2e docs)
V3 = {0: "ANGLE", 1: "PAREN", 2: "BRACE", 3: "BRACK", 4: "BAR", 5: "DBAR", 6: "FLOOR", 7: "CEILING",
      8: "LBLB", 9: "RBRB", 10: "RBLB", 11: "LBRP", 12: "LPRB", 13: "ROOT", 14: "FRACT", 15: "SCRIPT",
      16: "UBAR", 17: "OBAR", 18: "LARROW", 19: "RARROW", 20: "BARROW", 21: "SINT", 22: "DINT", 23: "TINT",
      24: "SSINT", 25: "DSINT", 26: "TSINT", 27: "UHBRACE", 28: "LHBRACE", 29: "SUM", 30: "ISUM", 31: "PROD",
      32: "IPROD", 33: "COPROD", 34: "ICOPROD", 35: "UNION", 36: "IUNION", 37: "INTER", 38: "IINTER",
      39: "LIM", 40: "LDIV", 41: "SLFRACT", 42: "INTOP", 43: "SUMOP", 44: "LSCRIPT", 45: "DIRAC",
      46: "UARROW", 47: "OARROW", 48: "OARC"}
V5 = {0: "ANGLE", 1: "PAREN", 2: "BRACE", 3: "BRACK", 4: "BAR", 5: "DBAR", 6: "FLOOR", 7: "CEILING",
      8: "OBRACK", 9: "INTERVAL", 10: "ROOT", 11: "FRACT", 12: "UBAR", 13: "OBAR", 14: "ARROW", 15: "INTEG",
      16: "SUM", 17: "PROD", 18: "COPROD", 19: "UNION", 20: "INTER", 21: "INTOP", 22: "SUMOP", 23: "LIM",
      24: "HBRACE", 25: "HBRACK", 26: "LDIV", 27: "SUB", 28: "SUP", 29: "SUBSUP", 30: "DIRAC", 31: "VEC",
      32: "TILDE", 33: "HAT", 34: "ARC", 35: "JSTATUS", 36: "STRIKE", 37: "BOX"}

FENCES = {"ANGLE": ("⟨", "⟩"), "PAREN": ("(", ")"), "BRACE": ("{", "}"), "BRACK": ("[", "]"),
          "BAR": ("|", "|"), "DBAR": ("‖", "‖"), "FLOOR": ("⌊", "⌋"),
          "CEILING": ("⌈", "⌉"), "OBRACK": ("[", "["), "INTERVAL": ("(", ")"),
          "LBLB": ("[", "["), "RBRB": ("]", "]"), "RBLB": ("]", "["), "LBRP": ("[", ")"), "LPRB": ("(", "]")}

EMBELL_ACCENT = {2: "˙", 3: "¨", 4: "⃛", 8: "~", 9: "^", 11: "→", 12: "←",
                 13: "↔", 14: "⇀", 15: "↽", 16: "¯", 17: "¯", 19: "⌢", 20: "⌣"}
EMBELL_PRIME = {5: "′", 6: "″", 7: "‵", 18: "‴"}

SPACE_RANGE = (0xEB00, 0xEBFF)  # MathType private-use spacing characters
SKIP_CODES = {0xEC07, 0xEC08, 0xEC09, 0xEC0A}  # stretchy fence pieces (MT Extra) -> chosen by template


class Warn(list):
    pass


def E(tag, text=None, **attrs):
    e = ET.Element(tag, {k.replace("_", "-"): v for k, v in attrs.items()})
    if text is not None:
        e.text = text
    return e


def mrow(children):
    children = [c for c in children if c is not None]
    if len(children) == 1:
        return children[0]
    r = E("mrow")
    r.extend(children)
    return r


class Converter:
    def __init__(self, version):
        self.version = version
        self.table = V3 if version < 5 else V5
        self.warnings = []

    # ---- characters -----------------------------------------------------------
    def tf_index(self, ch):
        return ch.typeface - 128 if ch.typeface >= 128 else ch.typeface

    def char_token(self, ch):
        """returns (cls, text) or None to skip."""
        code = ch.code
        if code is None:
            return None
        if SPACE_RANGE[0] <= code <= SPACE_RANGE[1] or code in SKIP_CODES:
            return None
        text = chr(code)
        if text == "-":
            text = "−"
        tf = self.tf_index(ch)
        if tf == TF_NUMBER:
            return ("mn", text)
        if tf == TF_TEXT or tf == TF_TEXT_FE:
            if text.isdigit():
                return ("mn", text)
            if text.isalpha() or text.isspace() or ord(text) > 0x2E80:
                return ("mtext", text)
            return ("mo", text)
        if tf == TF_FUNCTION:
            return ("fn", text) if text.isalpha() else ("mo", text)
        if tf in (TF_VARIABLE, TF_LCGREEK, TF_UCGREEK):
            return ("mi", text)
        if tf == TF_VECTOR:
            return ("mi", text)
        if text.isdigit():
            return ("mn", text)
        if text.isalpha() and tf not in (TF_SYMBOL, TF_MTEXTRA):
            return ("mi", text)
        if text == "∞":
            return ("mi", text)
        return ("mo", text)

    def char_element(self, ch):
        tok = self.char_token(ch)
        if tok is None:
            return None
        cls, text = tok
        el = E("mi" if cls == "fn" else cls, text)
        return self.apply_embells(el, ch)

    def apply_embells(self, el, ch):
        return self._embell_wrap(el, ch)

    def _embell_wrap(self, el, ch):
        for em in ch.embells:
            if not isinstance(em, Embell):
                continue
            if em.etype in EMBELL_PRIME:
                w = E("msup")
                w.append(el)
                w.append(E("mo", EMBELL_PRIME[em.etype]))
                el = w
            elif em.etype in EMBELL_ACCENT:
                w = E("mover", accent="true")
                w.append(el)
                w.append(E("mo", EMBELL_ACCENT[em.etype]))
                el = w
            elif em.etype == 10:  # "not": negated symbol
                if el.tag == "mo" and el.text:
                    el.text = el.text + "̸"
                else:
                    self.warnings.append("embellishment 'not' on non-operator")
            else:
                self.warnings.append(f"embellishment {em.etype}")
        return el

    # ---- sequences -------------------------------------------------------------
    def seq(self, children):
        """list of nodes (the contents of a LINE) -> list of MathML elements."""
        items = []  # ("tok", cls, text, ch) | ("el", element)
        for n in children:
            if isinstance(n, Skip):
                continue
            if isinstance(n, Char):
                tok = self.char_token(n)
                if tok is None:
                    continue
                if n.embells:
                    items.append(("el", self.char_element(n)))
                else:
                    items.append(("tok", tok[0], tok[1]))
            elif isinstance(n, Tmpl):
                name = self.table.get(n.sel, f"?{n.sel}")
                if name in ("SCRIPT", "SUB", "SUP", "SUBSUP", "LSCRIPT"):
                    items.append(("script", name, n))
                else:
                    el = self.template(n, name)
                    if el is not None:
                        items.append(("el", el))
            elif isinstance(n, Pile):
                items.append(("el", self.pile(n)))
            elif isinstance(n, Matrix):
                items.append(("el", self.matrix(n)))
            elif isinstance(n, Line):
                items.append(("el", self.line(n)))
        return self.assemble(items)

    def assemble(self, items):
        out = []
        i = 0
        # pass 1: merge runs of tokens
        merged = []
        while i < len(items):
            it = items[i]
            if it[0] == "tok":
                cls, text = it[1], it[2]
                j = i + 1
                if cls in ("mn", "mtext", "fn"):
                    while j < len(items) and items[j][0] == "tok" and items[j][1] == cls:
                        text += items[j][2]
                        j += 1
                    # decimal point between digits: "." tokens are mo
                    if cls == "mn":
                        while (j + 1 < len(items) and items[j][0] == "tok" and items[j][1] == "mo"
                               and items[j][2] in ".,．" and items[j + 1][0] == "tok"
                               and items[j + 1][1] == "mn" and items[j][2] == "." ):
                            text += "."
                            j += 1
                            while j < len(items) and items[j][0] == "tok" and items[j][1] == "mn":
                                text += items[j][2]
                                j += 1
                    merged.append(("el", E("mi" if cls == "fn" else cls, text)))
                else:
                    merged.append(("el", E(cls, text)))
                i = j
            else:
                merged.append(it)
                i += 1
        # pass 2: attach scripts to previous element
        for it in merged:
            if it[0] == "el":
                out.append(it[1])
            else:
                name, tmpl = it[1], it[2]
                out.append(("script", name, tmpl))
        res = []
        for it in out:
            if isinstance(it, tuple):
                _, name, tmpl = it
                base = res.pop() if res else E("mrow")
                res.append(self.script(base, name, tmpl))
            else:
                res.append(it)
        return res

    def slots(self, children):
        """Sub-object slots of a template: LINE / PILE / MATRIX / CHAR nodes, size records dropped."""
        return [c for c in children if not isinstance(c, Skip)]

    def line(self, ln):
        if ln.null:
            return None
        return mrow(self.seq(ln.children)) if ln.children else None

    def slot(self, node):
        if node is None:
            return None
        if isinstance(node, Line):
            return self.line(node)
        if isinstance(node, Char):
            return self.char_element(node)
        els = self.seq([node])
        return mrow(els) if els else None

    def slot_or_empty(self, node):
        el = self.slot(node)
        return el if el is not None else E("mrow")

    def script(self, base, name, t):
        slots = self.slots(t.children)
        lines = [s for s in slots if isinstance(s, Line)]
        sub = sup = None
        if name == "SUB":
            sub = self.slot(lines[0]) if lines else None
            if len(lines) > 1 and not lines[1].null:
                sup = self.slot(lines[1])
        elif name == "SUP":
            if len(lines) == 1:
                sup = self.slot(lines[0])
            else:
                sub = self.slot(lines[0]) if lines else None
                sup = self.slot(lines[1]) if len(lines) > 1 else None
        elif name == "SUBSUP":
            sub = self.slot(lines[0]) if lines else None
            sup = self.slot(lines[1]) if len(lines) > 1 else None
        else:  # v3 SCRIPT / LSCRIPT: variation 0 super, 1 sub, 2 subsup
            sub = self.slot(lines[0]) if lines else None
            sup = self.slot(lines[1]) if len(lines) > 1 else None
            if t.var == 0 and sup is None and sub is not None and len(lines) == 1:
                sup, sub = sub, None
        if name == "LSCRIPT":
            self.warnings.append("left script")
        if sub is not None and sup is not None:
            el = E("msubsup")
            el.extend([base, sub, sup])
        elif sup is not None:
            el = E("msup")
            el.extend([base, sup])
        elif sub is not None:
            el = E("msub")
            el.extend([base, sub])
        else:
            return base
        return el

    # ---- templates -------------------------------------------------------------
    def fence_chars(self, name, t, slots):
        chars = [s for s in slots if isinstance(s, Char) and s.code not in SKIP_CODES
                 and not (SPACE_RANGE[0] <= s.code <= SPACE_RANGE[1])]
        left, right = FENCES.get(name, ("(", ")"))
        if self.version < 5:
            has_l, has_r = t.var in (0, 1), t.var in (0, 2)
        else:
            has_l, has_r = bool(t.var & 1), bool(t.var & 2)
        if name == "INTERVAL":
            has_l = has_r = True
        texts = [chr(c.code) for c in chars]
        if has_l and has_r and len(texts) >= 2:
            left, right = texts[0], texts[1]
        elif has_l and not has_r and texts:
            left = texts[0]
        elif has_r and not has_l and texts:
            right = texts[-1]
        return (left if has_l else None), (right if has_r else None)

    def template(self, t, name):
        slots = self.slots(t.children)
        lines = [s for s in slots if not isinstance(s, Char)]
        main = lines[0] if lines else None
        if name in FENCES:
            left, right = self.fence_chars(name, t, slots)
            row = E("mrow")
            if left:
                row.append(E("mo", left, fence="true"))
            inner = self.slot(main)
            if inner is not None:
                row.append(inner)
            if right:
                row.append(E("mo", right, fence="true"))
            return row
        if name in ("FRACT", "SLFRACT"):
            num = self.slot_or_empty(lines[0] if lines else None)
            den = self.slot_or_empty(lines[1] if len(lines) > 1 else None)
            f = E("mfrac")
            if name == "SLFRACT":
                f.set("bevelled", "true")
            elif self.version < 5 and t.var == 1 or self.version >= 5 and t.var & 1:
                pass  # small fraction: same semantics
            f.extend([num, den])
            return f
        if name == "ROOT":
            rad = self.slot_or_empty(lines[0] if lines else None)
            idx = self.slot(lines[1]) if len(lines) > 1 else None
            if idx is None:
                r = E("msqrt")
                r.append(rad)
                return r
            r = E("mroot")
            r.extend([rad, idx])
            return r
        if name in ("UBAR", "OBAR"):
            w = E("munder" if name == "UBAR" else "mover", accent="true")
            w.extend([self.slot_or_empty(main), E("mo", "¯" if name == "OBAR" else "_")])
            return w
        if name in ("VEC", "TILDE", "HAT", "ARC", "ARROW", "UARROW", "OARROW", "LARROW", "RARROW", "BARROW"):
            return self.accent_like(t, name, lines, slots)
        if name in ("SUM", "PROD", "COPROD", "UNION", "INTER", "ISUM", "IPROD", "ICOPROD", "IUNION", "IINTER",
                    "INTEG", "SINT", "DINT", "TINT", "SSINT", "DSINT", "TSINT", "INTOP", "SUMOP"):
            return self.bigop(t, name, lines, slots)
        if name == "LIM":
            lim = E("mi", "lim")
            under = self.slot(lines[1]) if len(lines) > 1 else None
            if under is None:
                return lim
            w = E("munder")
            w.extend([lim, under])
            return w
        if name == "STRIKE":
            e = E("menclose", notation="updiagonalstrike")
            e.append(self.slot_or_empty(main))
            return e
        if name == "BOX":
            e = E("menclose", notation="box")
            e.append(self.slot_or_empty(main))
            return e
        if name == "LDIV":
            e = E("mrow")
            e.extend(self.seq(sum([self._flat(s) for s in lines], [])))
            self.warnings.append("long-division template approximated")
            return e
        self.warnings.append(f"unsupported template {name} (v{self.version} sel {t.sel})")
        row = E("mrow")
        for s in lines:
            el = self.slot(s)
            if el is not None:
                row.append(el)
        return row

    def _flat(self, s):
        return s.children if isinstance(s, Line) else [s]

    def accent_like(self, t, name, lines, slots):
        main = self.slot_or_empty(lines[0] if lines else None)
        chars = [s for s in slots if isinstance(s, Char)]
        mark = {"VEC": "→", "TILDE": "~", "HAT": "^", "ARC": "⌒"}.get(name)
        if mark is None:
            mark = chr(chars[0].code) if chars else "→"
        under = name in ("BARROW",) or (name == "ARROW" and False)
        w = E("mover", accent="true")
        w.extend([main, E("mo", mark)])
        return w

    def bigop(self, t, name, lines, slots):
        chars = [s for s in slots if isinstance(s, Char)]
        glyph = {"SUM": "∑", "PROD": "∏", "COPROD": "∐", "UNION": "⋃", "INTER": "⋂",
                 "ISUM": "∑", "IPROD": "∏", "ICOPROD": "∐", "IUNION": "⋃",
                 "IINTER": "⋂", "INTEG": "∫", "SINT": "∫", "DINT": "∬", "TINT": "∭",
                 "SSINT": "∫", "DSINT": "∬", "TSINT": "∭"}.get(name)
        if glyph is None and chars:
            glyph = chr(chars[0].code)
        glyph = glyph or "∑"
        # slot order: integrand/main, upper, lower
        main = self.slot(lines[0]) if lines else None
        upper = self.slot(lines[1]) if len(lines) > 1 else None
        lower = self.slot(lines[2]) if len(lines) > 2 else None
        is_int = glyph in "∫∬∭"
        op = E("mo", glyph)
        if upper is not None or lower is not None:
            tag_all = "msubsup" if is_int else "munderover"
            tag_low, tag_up = ("msub", "msup") if is_int else ("munder", "mover")
            if upper is not None and lower is not None:
                w = E(tag_all)
                w.extend([op, lower, upper])
            elif lower is not None:
                w = E(tag_low)
                w.extend([op, lower])
            else:
                w = E(tag_up)
                w.extend([op, upper])
            op = w
        row = E("mrow")
        row.append(op)
        if main is not None:
            row.append(main)
        return row

    def pile(self, p):
        tbl = E("mtable")
        for ln in p.children:
            if isinstance(ln, Line):
                row = E("mtr")
                cell = E("mtd")
                el = self.line(ln)
                if el is not None:
                    cell.append(el)
                row.append(cell)
                tbl.append(row)
        return tbl

    def matrix(self, m):
        tbl = E("mtable")
        cells = [c for c in m.children if isinstance(c, Line)]
        for r in range(m.rows):
            row = E("mtr")
            for c in range(m.cols):
                idx = r * m.cols + c
                cell = E("mtd")
                if idx < len(cells):
                    el = self.line(cells[idx])
                    if el is not None:
                        cell.append(el)
                row.append(cell)
            tbl.append(row)
        return tbl


def to_mathml(parser, nodes):
    cv = Converter(parser.version)
    # top level: usually a single LINE or PILE
    els = cv.seq(nodes)
    math = ET.Element("math", {"xmlns": MATHML_NS})
    math.append(mrow(els) if els else E("mrow"))
    return math, cv.warnings


def linear(el):
    """compact human readable rendering used for checking."""
    t = el.tag
    kids = [linear(c) for c in el]
    if t in ("mi", "mn", "mo", "mtext"):
        return el.text or ""
    if t == "mfrac":
        return f"({kids[0]})/({kids[1]})"
    if t == "msup":
        return f"{kids[0]}^({kids[1]})"
    if t == "msub":
        return f"{kids[0]}_({kids[1]})"
    if t == "msubsup":
        return f"{kids[0]}_({kids[1]})^({kids[2]})"
    if t == "msqrt":
        return f"sqrt({kids[0]})"
    if t == "mroot":
        return f"root[{kids[1]}]({kids[0]})"
    if t == "mover":
        return f"over({kids[0]},{kids[1]})"
    if t == "munder":
        return f"under({kids[0]},{kids[1]})"
    if t == "munderover":
        return f"{kids[0]}_({kids[1]})^({kids[2]})"
    if t == "mtr":
        return "[" + " | ".join(kids) + "]"
    if t == "mtd":
        return "".join(kids)
    if t == "mtable":
        return "{" + "; ".join(kids) + "}"
    return "".join(kids)
