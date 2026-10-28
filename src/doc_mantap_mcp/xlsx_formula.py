"""A small Excel-formula evaluator for create_excel.

It validates structured references (``Sales[Revenue]``, ``[@Qty]``) against the tables in
the spec, and computes cached results for common formulas so previews that do not
recalculate (Quick Look, mail clients, some LibreOffice settings) still show real numbers.
Anything it does not understand is left to Excel, which recalculates on open.
"""

from __future__ import annotations

import fnmatch
import math
import re
from typing import Any, Callable


class FormulaError(ValueError):
    """A problem the author must fix (unknown table/column, unbalanced brackets...)."""


class Unsupported(Exception):
    """Valid Excel we do not evaluate; Excel will compute it."""


class ExcelError(Exception):
    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


_NUM_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_CELL_RE = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?")
_NAME_RE = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.]*")
_OPS2 = ("<>", "<=", ">=")
_OPS1 = set("+-*/^&=<>(),:;%")


def _bracket(s: str, pos: int) -> tuple[str, int]:
    depth, i = 0, pos
    while i < len(s):
        ch = s[i]
        if ch == "'":
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return s[pos + 1:i], i + 1
        i += 1
    raise FormulaError("Unbalanced [ ] brackets in formula.")


def _unescape(name: str) -> str:
    return re.sub(r"'(.)", r"\1", name).strip()


def parse_struct(body: str) -> tuple[bool, list[str], list[str]]:
    """Structured reference body -> (this_row, columns, specials)."""
    body = body.strip()
    if body.startswith("@"):
        rest = body[1:].strip()
        if rest.startswith("[") and rest.endswith("]"):
            rest = rest[1:-1]
        return True, ([_unescape(rest)] if rest else []), ["#This Row"]
    if "[" not in body:
        return (False, [], [body]) if body.startswith("#") else (False, [_unescape(body)], [])
    items = re.findall(r"\[((?:[^\[\]']|'.)*)\]", body)
    specials = [i.strip() for i in items if i.strip().startswith("#")]
    cols = [_unescape(i) for i in items if not i.strip().startswith("#")]
    return any(s.lower() == "#this row" for s in specials), cols, specials


def tokenize(formula: str) -> list[tuple]:
    s = formula[1:] if formula.startswith("=") else formula
    out, pos = [], 0
    while pos < len(s):
        ch = s[pos]
        if ch.isspace():
            pos += 1
        elif ch == '"':
            end = pos + 1
            while True:
                end = s.find('"', end)
                if end < 0:
                    raise FormulaError("Unclosed string literal in formula.")
                if end + 1 < len(s) and s[end + 1] == '"':
                    end += 2
                    continue
                break
            out.append(("str", s[pos + 1:end].replace('""', '"')))
            pos = end + 1
        elif ch == "[":
            body, pos = _bracket(s, pos)
            out.append(("struct", None, body))
        elif ch == "'":  # 'Sheet name'!A1
            end = s.find("'!", pos + 1)
            if end < 0:
                raise FormulaError("Unclosed sheet name quote in formula.")
            pos = end + 2
            m = _CELL_RE.match(s, pos) or _NAME_RE.match(s, pos)
            pos = m.end() if m else pos
            out.append(("cell", s[:pos]))
        elif s.startswith(_OPS2, pos):
            out.append(("op", s[pos:pos + 2]))
            pos += 2
        elif ch in _OPS1:
            out.append(("op", ";" if ch == ";" else ch))
            pos += 1
        elif (m := _CELL_RE.match(s, pos)) and not (m.end() < len(s) and
                                                    (s[m.end()] == "(" or s[m.end()].isalnum())):
            out.append(("cell", m.group()))
            pos = m.end()
        elif m := _NUM_RE.match(s, pos):
            out.append(("num", float(m.group())))
            pos = m.end()
        elif m := _NAME_RE.match(s, pos):
            name, pos = m.group(), m.end()
            if pos < len(s) and s[pos] == "[":
                body, pos = _bracket(s, pos)
                out.append(("struct", name, body))
            elif pos < len(s) and s[pos] == "!":
                m2 = _CELL_RE.match(s, pos + 1) or _NAME_RE.match(s, pos + 1)
                pos = m2.end() if m2 else pos + 1
                out.append(("cell", name))
            else:
                out.append(("name", name))
        else:
            raise FormulaError(f"Unexpected character '{ch}' in formula.")
    return out


# --------------------------------------------------------------------------- #
# Parsing into an AST
# --------------------------------------------------------------------------- #

class _Parser:
    def __init__(self, tokens):
        self.t, self.i = tokens, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def is_op(self, *ops):
        k, v = self.peek()[:2]
        return k == "op" and v in ops

    def parse(self):
        node = self.compare()
        if self.i != len(self.t):
            raise Unsupported("trailing tokens")
        return node

    def compare(self):
        node = self.concat()
        while self.is_op("=", "<>", "<", ">", "<=", ">="):
            node = ("bin", self.take()[1], node, self.concat())
        return node

    def concat(self):
        node = self.add()
        while self.is_op("&"):
            self.take()
            node = ("bin", "&", node, self.add())
        return node

    def add(self):
        node = self.mul()
        while self.is_op("+", "-"):
            node = ("bin", self.take()[1], node, self.mul())
        return node

    def mul(self):
        node = self.power()
        while self.is_op("*", "/"):
            node = ("bin", self.take()[1], node, self.power())
        return node

    def power(self):
        node = self.unary()
        while self.is_op("^"):
            self.take()
            node = ("bin", "^", node, self.unary())
        return node

    def unary(self):
        if self.is_op("-"):
            self.take()
            return ("neg", self.unary())
        if self.is_op("+"):
            self.take()
            return self.unary()
        node = self.range_()
        while self.is_op("%"):
            self.take()
            node = ("pct", node)
        return node

    def range_(self):
        node = self.primary()
        while self.is_op(":"):
            self.take()
            node = ("range", node, self.primary())
        return node

    def primary(self):
        tok = self.take()
        kind = tok[0]
        if kind == "num":
            return ("val", tok[1])
        if kind == "str":
            return ("val", tok[1])
        if kind == "struct":
            return ("struct", tok[1], tok[2])
        if kind == "cell":
            return ("cell", tok[1])
        if kind == "op" and tok[1] == "(":
            node = self.compare()
            if not self.is_op(")"):
                raise FormulaError("Missing ')' in formula.")
            self.take()
            return node
        if kind == "name":
            name = tok[1]
            if self.is_op("("):
                self.take()
                args = []
                if not self.is_op(")"):
                    while True:
                        args.append(self.compare())
                        if self.is_op(",", ";"):
                            self.take()
                            continue
                        break
                if not self.is_op(")"):
                    raise FormulaError(f"Missing ')' after {name}( in formula.")
                self.take()
                return ("call", name.upper().removeprefix("_XLFN."), args)
            if name.upper() in ("TRUE", "FALSE"):
                return ("val", name.upper() == "TRUE")
            return ("name", name)
        if kind is None:
            raise FormulaError("Formula ends unexpectedly.")
        raise FormulaError(f"Unexpected '{tok[1]}' in formula.")


def compile_formula(formula: str):
    return _Parser(tokenize(formula)).parse()


def struct_refs(node, out=None) -> list[tuple[str | None, str]]:
    """All (table, body) structured references in an AST."""
    out = [] if out is None else out
    if isinstance(node, tuple):
        if node[0] == "struct":
            out.append((node[1], node[2]))
        for part in node[1:]:
            if isinstance(part, tuple):
                struct_refs(part, out)
            elif isinstance(part, list):
                for p in part:
                    struct_refs(p, out)
    return out


# --------------------------------------------------------------------------- #
# Evaluation
# --------------------------------------------------------------------------- #

def _num(v) -> float:
    if isinstance(v, ExcelError):
        raise v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, list):
        raise Unsupported("range used as a number")
    try:
        return float(str(v))
    except ValueError:
        raise ExcelError("#VALUE!") from None


def _text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _flat(args):
    for a in args:
        if isinstance(a, list):
            yield from a
        else:
            yield a


def _numbers(args):
    for a in args:
        if isinstance(a, list):
            for v in a:
                if isinstance(v, ExcelError):
                    raise v
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    yield float(v)
        else:
            yield _num(a)


def _criteria(c) -> Callable[[Any], bool]:
    if isinstance(c, (int, float)) and not isinstance(c, bool):
        return lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and v == c
    m = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", _text(c), re.S)
    op, rhs = m.group(1) or "=", m.group(2)
    try:
        target = float(rhs)
        numeric = True
    except ValueError:
        numeric = False
    cmp = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b, "<": lambda a, b: a < b,
           ">": lambda a, b: a > b, "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}[op]

    def test(v):
        if numeric:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return cmp(float(v), target)
            return op == "<>"
        text = _text(v).lower()
        if op in ("=", "<>") and any(ch in rhs for ch in "*?"):
            hit = fnmatch.fnmatchcase(text, rhs.lower())
            return hit if op == "=" else not hit
        return cmp(text, rhs.lower())
    return test


def _round(x, n=0):
    x, n = _num(x), int(_num(n))
    factor = 10 ** n
    return math.copysign(math.floor(abs(x) * factor + 0.5) / factor, x)


def _avg(values):
    values = list(values)
    if not values:
        raise ExcelError("#DIV/0!")
    return sum(values) / len(values)


def _mod(a, b):
    a, b = _num(a), _num(b)
    if b == 0:
        raise ExcelError("#DIV/0!")
    return a - b * math.floor(a / b)


def _pick(values, criteria, target=None):
    """Values of `target` (default `values`) where `values` matches `criteria`."""
    if not isinstance(values, list):
        raise Unsupported("criteria range is not a column")
    target = values if target is None else target
    test = _criteria(criteria)
    return [target[i] for i, v in enumerate(values) if test(v)
            and isinstance(target[i], (int, float)) and not isinstance(target[i], bool)]


def _ifs(func: str, args):
    if func == "SUMIFS":
        target, pairs = args[0], args[1:]
    else:
        target, pairs = None, args
    if len(pairs) % 2 or not pairs:
        raise Unsupported("bad *IFS arguments")
    ranges = [pairs[i] for i in range(0, len(pairs), 2)]
    tests = [_criteria(pairs[i]) for i in range(1, len(pairs), 2)]
    if not all(isinstance(r, list) for r in ranges):
        raise Unsupported("criteria range is not a column")
    size = len(ranges[0])
    hits = [i for i in range(size) if all(t(r[i]) for r, t in zip(ranges, tests))]
    if func == "COUNTIFS":
        return float(len(hits))
    vals = [target[i] for i in hits if isinstance(target[i], (int, float))
            and not isinstance(target[i], bool)]
    return float(sum(vals)) if func == "SUMIFS" else _avg(vals)


_FUNCS: dict[str, Callable] = {
    "SUM": lambda *a: sum(_numbers(a)),
    "AVERAGE": lambda *a: _avg(_numbers(a)),
    "MIN": lambda *a: min(list(_numbers(a)) or [0.0]),
    "MAX": lambda *a: max(list(_numbers(a)) or [0.0]),
    "COUNT": lambda *a: float(sum(1 for v in _flat(a) if isinstance(v, (int, float))
                                  and not isinstance(v, bool))),
    "COUNTA": lambda *a: float(sum(1 for v in _flat(a) if v not in (None, ""))),
    "MEDIAN": lambda *a: (lambda v: (v[len(v) // 2] if len(v) % 2 else
                                     (v[len(v) // 2 - 1] + v[len(v) // 2]) / 2) if v
                          else 0.0)(sorted(_numbers(a))),
    "ROUND": _round,
    "ABS": lambda x: abs(_num(x)),
    "INT": lambda x: float(math.floor(_num(x))),
    "MOD": _mod,
    "SQRT": lambda x: math.sqrt(_num(x)),
    "POWER": lambda a, b: _num(a) ** _num(b),
    "AND": lambda *a: all(bool(_num(v)) for v in _flat(a)),
    "OR": lambda *a: any(bool(_num(v)) for v in _flat(a)),
    "NOT": lambda v: not bool(_num(v)),
    "SUMIF": lambda r, c, s=None: float(sum(_pick(r, c, s))),
    "COUNTIF": lambda r, c: float(sum(1 for v in r if _criteria(c)(v))),
    "AVERAGEIF": lambda r, c, s=None: _avg(_pick(r, c, s)),
    "CONCAT": lambda *a: "".join(_text(v) for v in _flat(a)),
    "CONCATENATE": lambda *a: "".join(_text(v) for v in a),
    "LEN": lambda v: float(len(_text(v))),
    "UPPER": lambda v: _text(v).upper(),
    "LOWER": lambda v: _text(v).lower(),
}


class Evaluator:
    """Evaluate compiled formulas against the workbook's tables and defined names."""

    def __init__(self, tables: dict, names: dict[str, Any]):
        # tables: lower-case name -> object with .columns (list[str]) and
        #         .value(row, col_idx, evaluator), .column(col_idx, evaluator)
        self.tables = tables
        self.names = {k.lower(): v for k, v in names.items()}

    def check(self, node, table_key: str | None) -> None:
        """Raise FormulaError for references to unknown tables or columns."""
        for tname, body in struct_refs(node):
            key = (tname or "").lower() or table_key
            if key is None:
                raise FormulaError(f"'[{body}]' needs a table name outside a table, "
                                   f"e.g. Sales[{body.lstrip('@')}].")
            if key not in self.tables:
                known = ", ".join(t.name for t in self.tables.values()) or "none"
                raise FormulaError(f"Unknown table '{tname}'. Tables: {known}.")
            table = self.tables[key]
            _, cols, _ = parse_struct(body)
            for col in cols:
                if table.index(col) is None:
                    raise FormulaError(f"Table '{table.name}' has no column '{col}'. "
                                       f"Columns: {', '.join(table.columns)}.")

    def value(self, node, table_key: str | None = None, row: int | None = None):
        """Cached result, or None when it cannot be computed here."""
        try:
            result = self._eval(node, table_key, row)
        except (Unsupported, ExcelError, RecursionError, ValueError, TypeError,
                OverflowError, IndexError, ZeroDivisionError):
            return None
        if isinstance(result, (list, ExcelError)):
            return None
        if isinstance(result, float) and (math.isnan(result) or math.isinf(result)):
            return None
        return result

    def _eval(self, node, tk, row):
        kind = node[0]
        if kind == "val":
            return node[1]
        if kind == "neg":
            return -_num(self._eval(node[1], tk, row))
        if kind == "pct":
            return _num(self._eval(node[1], tk, row)) / 100
        if kind in ("cell", "range"):
            raise Unsupported("A1 references")
        if kind == "name":
            key = node[1].lower()
            if key in self.names:
                return self.names[key]
            raise Unsupported(f"name {node[1]}")
        if kind == "struct":
            key = (node[1] or "").lower() or tk
            table = self.tables.get(key)
            if table is None:
                raise Unsupported("table")
            this_row, cols, specials = parse_struct(node[2])
            if len(cols) != 1 or any(s.lower() not in ("#this row", "#data") for s in specials):
                raise Unsupported("complex structured reference")
            idx = table.index(cols[0])
            if this_row:
                if key != tk or row is None:
                    raise Unsupported("this-row reference outside its table")
                return table.value(row, idx, self)
            return table.column(idx, self)
        if kind == "bin":
            op = node[1]
            a, b = self._eval(node[2], tk, row), self._eval(node[3], tk, row)
            if op == "&":
                return _text(a) + _text(b)
            if op in ("=", "<>", "<", ">", "<=", ">="):
                if isinstance(a, str) or isinstance(b, str):
                    a, b = _text(a).lower(), _text(b).lower()
                else:
                    a, b = _num(a), _num(b)
                return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b,
                        ">=": a >= b}[op]
            x, y = _num(a), _num(b)
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            if op == "/":
                if y == 0:
                    raise ExcelError("#DIV/0!")
                return x / y
            if op == "^":
                return x ** y
        if kind == "call":
            name, args = node[1], node[2]
            if name == "IF":
                cond = _num(self._eval(args[0], tk, row)) if args else 0
                if cond:
                    return self._eval(args[1], tk, row) if len(args) > 1 else True
                return self._eval(args[2], tk, row) if len(args) > 2 else False
            if name == "IFERROR":
                try:
                    v = self._eval(args[0], tk, row)
                    if isinstance(v, ExcelError):
                        raise v
                    return v
                except (ExcelError, ZeroDivisionError):
                    return self._eval(args[1], tk, row)
            if name in ("SUMIFS", "COUNTIFS", "AVERAGEIFS"):
                return _ifs(name, [self._eval(a, tk, row) for a in args])
            func = _FUNCS.get(name)
            if func is None:
                raise Unsupported(name)
            return func(*[self._eval(a, tk, row) for a in args])
        raise Unsupported(kind)
