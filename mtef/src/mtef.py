"""MTEF (MathType Equation Format) v3 / v5 parser.

Reads the 'Equation Native' stream of an Equation.3 / Equation.DSMT4 OLE object
and produces a small tree of nodes.  Pure Python, no third-party dependencies.

Node kinds (all are plain objects with a .kind attribute):
  Line(children, null)      Char(typeface, code, embells)   Tmpl(sel, var, opts, children)
  Pile(halign, valign, children)   Matrix(rows, cols, children)   Size(...)   (sizes are kept but ignored)
"""
import struct

# record types
END, LINE, CHAR, TMPL, PILE, MATRIX, EMBELL, RULER = 0, 1, 2, 3, 4, 5, 6, 7
FONT_STYLE_DEF_OR_FONT, SIZE, FULL, SUB, SUB2, SYM, SUBSYM = 8, 9, 10, 11, 12, 13, 14
COLOR, COLOR_DEF, FONT_DEF, EQN_PREFS, ENCODING_DEF = 15, 16, 17, 18, 19


class MtefError(Exception):
    pass


class Node:
    kind = "node"

    def __repr__(self):
        return f"<{self.__class__.__name__} {self.__dict__}>"


class Line(Node):
    kind = "line"

    def __init__(self, children, null=False):
        self.children, self.null = children, null


class Char(Node):
    kind = "char"

    def __init__(self, typeface, code, embells, opts=0, code8=None):
        self.typeface, self.code, self.embells, self.opts, self.code8 = typeface, code, embells, opts, code8


class Tmpl(Node):
    kind = "tmpl"

    def __init__(self, sel, var, opts, children):
        self.sel, self.var, self.opts, self.children = sel, var, opts, children


class Pile(Node):
    kind = "pile"

    def __init__(self, halign, valign, children):
        self.halign, self.valign, self.children = halign, valign, children


class Matrix(Node):
    kind = "matrix"

    def __init__(self, rows, cols, children, valign=0, hjust=0, vjust=0):
        self.rows, self.cols, self.children = rows, cols, children
        self.valign, self.hjust, self.vjust = valign, hjust, vjust


class Embell(Node):
    kind = "embell"

    def __init__(self, etype):
        self.etype = etype


class Skip(Node):
    """A record we parsed but do not need (size, font def ...)."""
    kind = "skip"

    def __init__(self, what):
        self.what = what


class Reader:
    def __init__(self, data, pos=0):
        self.d, self.p = data, pos

    def eof(self):
        return self.p >= len(self.d)

    def u8(self):
        if self.p >= len(self.d):
            raise MtefError("unexpected end of data")
        v = self.d[self.p]
        self.p += 1
        return v

    def u16(self):
        lo = self.u8()
        return lo | (self.u8() << 8)

    def i16(self):
        v = self.u16()
        return v - 65536 if v >= 32768 else v

    def cstr(self):
        s = self.p
        while self.u8() != 0:
            pass
        return self.d[s:self.p - 1]


OLE_HEADER_FORMAT = "<HIHIIII"  # cbHdr, version, cf, cbObject, reserved...


def strip_ole_header(stream):
    """The 'Equation Native' stream starts with a 28 byte EQNOLEFILEHDR."""
    if len(stream) < 28:
        raise MtefError("stream too short")
    cb = struct.unpack_from("<H", stream, 0)[0]
    if cb != 28:
        cb = 28
    return stream[cb:]


class Parser:
    def __init__(self, mtef):
        self.r = Reader(mtef)
        self.warnings = []
        r = self.r
        self.version = r.u8()
        if self.version not in (2, 3, 5):
            raise MtefError(f"unsupported MTEF version {self.version}")
        self.platform = r.u8()
        self.product = r.u8()
        self.prod_version = r.u8()
        self.prod_subversion = r.u8()
        self.app_key = b""
        if self.version == 5:
            self.app_key = r.cstr()
            self.eq_options = r.u8()

    # ---- helpers
    def nudge(self):
        r = self.r
        dx, dy = r.u8(), r.u8()
        if dx == 128 and dy == 128:
            r.i16()
            r.i16()

    def dimension_array(self):
        """Array of nibble-packed dimensions (EQN_PREFS)."""
        r = self.r
        n = r.u8()
        out = []
        nibs = []
        have = False

        def nib():
            nonlocal have
            b = r.u8()
            return b >> 4, b & 15

        for _ in range(n):
            # a dimension: units nibble then value nibbles terminated by 0xF
            val = []
            while True:
                if not nibs:
                    hi, lo = nib()
                    nibs.extend((hi, lo))
                x = nibs.pop(0)
                if x == 15:
                    break
                val.append(x)
            out.append(val)
        # the whole array is byte aligned
        nibs.clear()
        return out

    def ruler(self):
        r = self.r
        n = r.u8()
        for _ in range(n):
            r.u8()
            r.i16()

    # ---- list of objects, terminated by END
    def objlist(self):
        out = []
        r = self.r
        while True:
            if r.eof():
                self.warnings.append("EOF inside object list")
                return out
            node = self.record()
            if node is END:
                return out
            if node is not None:
                out.append(node)

    def record(self):
        r = self.r
        tag = r.u8()
        if self.version == 5:
            typ, opt = tag, None
        else:
            typ, opt = tag & 0x0F, tag >> 4
        if typ == END:
            return END
        if typ == LINE:
            return self.line(opt)
        if typ == CHAR:
            return self.char(opt)
        if typ == TMPL:
            return self.tmpl(opt)
        if typ == PILE:
            return self.pile(opt)
        if typ == MATRIX:
            return self.matrix(opt)
        if typ == EMBELL:
            if opt is None:
                opt = r.u8()
            if opt & 8:
                self.nudge()
            return Embell(r.u8())
        if typ == RULER:
            self.ruler()
            return Skip("ruler")
        if typ == SIZE:
            a = r.u8()
            if a == 101:
                r.u16()
            elif a == 100:
                r.u8()
                r.u16()
            else:
                r.u8()
            return Skip("size")
        if typ in (FULL, SUB, SUB2, SYM, SUBSYM):
            return Skip("typesize")
        if self.version == 5:
            if typ == COLOR:
                r.u8()
                return Skip("color")
            if typ == COLOR_DEF:
                o = r.u8()
                n = 4 if o & 1 else 3
                for _ in range(n):
                    r.u16()
                if o & 4:
                    r.cstr()
                return Skip("colordef")
            if typ == FONT_DEF:
                r.u8()
                r.cstr()
                return Skip("fontdef")
            if typ == ENCODING_DEF:
                r.cstr()
                return Skip("encdef")
            if typ == EQN_PREFS:
                r.u8()  # options
                self.dimension_array()  # sizes
                self.dimension_array()  # spaces
                n = r.u8()  # styles
                for _ in range(n):
                    if r.u8() != 0:
                        r.u8()
                return Skip("prefs")
            if typ == 8:  # FONT_STYLE_DEF
                r.u8()
                r.u8()
                return Skip("fontstyle")
            raise MtefError(f"unknown v5 record {typ} at {r.p - 1}")
        if typ == 8:  # v3 FONT
            r.u8()
            r.u8()
            r.cstr()
            return Skip("font")
        raise MtefError(f"unknown record {typ} at {r.p - 1}")

    def line(self, opt):
        r = self.r
        if opt is None:
            opt = r.u8()
        if opt & 8:
            self.nudge()
        if opt & 4:
            r.i16()
        if opt & 2:
            if r.u8() & 0x0F != RULER and self.version != 5:
                raise MtefError("ruler expected")
            self.ruler()
        if opt & 1:
            return Line([], null=True)
        return Line(self.objlist())

    def char(self, opt):
        r = self.r
        if opt is None:
            opt = r.u8()
        if opt & 8:
            self.nudge()
        typeface = r.u8()
        code8 = None
        if self.version == 5:
            code = None
            if not opt & 0x20:
                code = r.u16()
            if opt & 4:
                code8 = r.u8()
            if opt & 0x10:
                code8 = r.u16()
            if code is None:
                code = code8
            has_embell = bool(opt & 1)
        else:
            code = r.u16() if self.version == 3 else r.u8()
            has_embell = bool(opt & 2)
        embells = self.objlist() if has_embell else []
        return Char(typeface, code, embells, opt, code8)

    def tmpl(self, opt):
        r = self.r
        if opt is None:
            opt = r.u8()
        if opt & 8:
            self.nudge()
        sel = r.u8()
        var = r.u8()
        if self.version == 5 and var & 0x80:
            var = (var & 0x7F) | (r.u8() << 8)
        topt = r.u8()
        return Tmpl(sel, var, topt, self.objlist())

    def pile(self, opt):
        r = self.r
        if opt is None:
            opt = r.u8()
        if opt & 8:
            self.nudge()
        h, v = r.u8(), r.u8()
        if opt & 2:
            if self.version != 5:
                r.u8()
            self.ruler()
        return Pile(h, v, self.objlist())

    def matrix(self, opt):
        r = self.r
        if opt is None:
            opt = r.u8()
        if opt & 8:
            self.nudge()
        valign, hj, vj, rows, cols = r.u8(), r.u8(), r.u8(), r.u8(), r.u8()
        for n in (rows + 1, cols + 1):
            for _ in range((n * 2 + 7) // 8):
                r.u8()
        return Matrix(rows, cols, self.objlist(), valign, hj, vj)

    def parse(self):
        nodes = []
        r = self.r
        while not r.eof():
            node = self.record()
            if node is END:
                continue
            if node is not None:
                nodes.append(node)
        return nodes


def parse_equation_native(stream):
    """Parse a full 'Equation Native' stream. Returns (parser, nodes)."""
    p = Parser(strip_ole_header(stream))
    return p, p.parse()
