"""Plovdiv full regeneration oracle (generator v1 acceptance test, blueprint 06
section 9a method extended to every block). Regenerates every supported formula
cell on the Aero, Non-aero, Opex and Operations Summary sheets from Plovdiv's
own cached Inputs values (openpyxl data_only), in dependency order via a
fixed-point pass, and diffs against Plovdiv's cached results. Tolerance zero
on block formulas (reported as max absolute difference; pass threshold 1e-9
for float representation). Unsupported formulas are listed, never silently
skipped.

DEVELOPMENT AND ACCEPTANCE TOOL ONLY: this is not part of the production build_model
chain. Its formula evaluator is a scoped interpreter that supports the Avia block
grammars, not a general Excel engine, so it is for validating the generator against a
known-good model, not for running client work.

The source workbook is resolved from the data root by config.find_reference, not
hardcoded. It previously carried an absolute path from a working session, which meant
the zero-difference claim could not be reproduced on any other machine.
Author: Avia Solutions."""
import sys, re, openpyxl

import json, os, tempfile
import config
SRC = config.find_reference("Plovdiv_FinancialModel_vDraft.xlsm")
CALC_SHEETS = sys.argv[1].split(",")
STATE = os.path.join(tempfile.gettempdir(), "oracle_state.json")
GIVEN_SHEETS = {"Inputs", "Control", "Capex", "Summary table", "Returns", "Transfer Sheet >>"}

wbf = openpyxl.load_workbook(SRC)
wbv = openpyxl.load_workbook(SRC, data_only=True)

SUPPORTED_FN = {"SUM", "IFERROR", "IF", "MAX", "MIN", "ROUND", "AVERAGE", "ABS"}
FN_RE = re.compile(r"([A-Z][A-Z0-9\.]*)\(")
REF_RE = re.compile(r"(?:(?:'([^']+)'|([A-Za-z][\w\. ]*?))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?")

def col_to_i(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n

def norm(ref):
    return ref.replace("$", "")

class Oracle:
    def __init__(self):
        self.regen = {}      # (sheet, 'A1') -> regenerated value
        self.unsupported = []

    def cached(self, sheet, ref):
        m = re.match(r"([A-Z]+)(\d+)", norm(ref))
        return wbv[sheet].cell(row=int(m.group(2)), column=col_to_i(m.group(1))).value

    def value(self, sheet, ref):
        ref = norm(ref)
        key = (sheet, ref)
        if key in self.regen:
            return self.regen[key]
        if sheet in GIVEN_SHEETS:
            return self.cached(sheet, ref)
        cell = wbf[sheet][ref]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            # calc-sheet formula not yet regenerated: rows 1-6 headers are given
            if cell.row <= 6:
                return self.cached(sheet, ref)
            raise KeyError(key)
        return cell.value

    def expand_range(self, sheet, a, b):
        ma, mb = re.match(r"([A-Z]+)(\d+)", norm(a)), re.match(r"([A-Z]+)(\d+)", norm(b))
        c1, r1 = col_to_i(ma.group(1)), int(ma.group(2))
        c2, r2 = col_to_i(mb.group(1)), int(mb.group(2))
        out = []
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                col = ""
                cc = c
                while cc:
                    cc, rem = divmod(cc - 1, 26)
                    col = chr(65 + rem) + col
                out.append(f"{col}{r}")
        return out

    def evaluate(self, sheet, formula):
        """Transform the Excel formula to python and evaluate. Raises on unsupported."""
        f = formula[1:]
        for fn in set(FN_RE.findall(f)):
            if fn not in SUPPORTED_FN:
                raise ValueError(f"unsupported function {fn}")
        f = f.replace("<>", "!=").replace("^", "**")
        # ranges then single refs
        def range_sub(m):
            sh = m.group(1) or m.group(2) or sheet
            if m.group(4):
                cells = self.expand_range(sh, m.group(3), m.group(4))
                vals = [self.value(sh.strip(), c) for c in cells]
                return "[" + ",".join(repr(0 if v is None else v) for v in vals) + "]"
            v = self.value(sh.strip(), m.group(3))
            return repr(0 if v is None else v)
        f = REF_RE.sub(range_sub, f)
        # excel functions to python
        f = re.sub(r"\bSUM\(", "_sum(", f)
        f = re.sub(r"\bAVERAGE\(", "_avg(", f)
        f = re.sub(r"\bMAX\(", "_max(", f)
        f = re.sub(r"\bMIN\(", "_min(", f)
        f = re.sub(r"\bROUND\(", "round(", f)
        f = re.sub(r"\bABS\(", "abs(", f)
        f = self.wrap_lazy(f, "IFERROR", "_ifer")
        f = self.wrap_lazy(f, "IF", "_if")
        f = f.replace("=", "==").replace("<==", "<=").replace(">==", ">=").replace("!==", "!=")
        env = {"_sum": lambda *a: sum(x for arg in a for x in (arg if isinstance(arg, list) else [arg])
                                      if isinstance(x, (int, float))),
               "_avg": lambda *a: (lambda xs: sum(xs)/len(xs))([x for arg in a for x in
                                   (arg if isinstance(arg, list) else [arg]) if isinstance(x, (int, float))]),
               "_max": lambda *a: max(x for arg in a for x in (arg if isinstance(arg, list) else [arg])
                                      if isinstance(x, (int, float))),
               "_min": lambda *a: min(x for arg in a for x in (arg if isinstance(arg, list) else [arg])
                                      if isinstance(x, (int, float))),
               "_ifer": lambda fn, alt: self._ifer(fn, alt),
               "_if": lambda c, a=True, b=False: (a() if callable(a) else a) if c else (b() if callable(b) else b),
               "round": round, "abs": abs, "TRUE": True, "FALSE": False}
        return eval(f, {"__builtins__": {}}, env)

    @staticmethod
    def _ifer(fn, alt):
        try:
            v = fn() if callable(fn) else fn
            if isinstance(v, complex):
                return alt() if callable(alt) else alt
            return v
        except Exception:
            return alt() if callable(alt) else alt

    def wrap_lazy(self, f, name, target):
        """NAME(a,b[,c]) -> target(lambda:(a), lambda:(b), ...) with paren matching;
        IF must not catch IFERROR (handle IFERROR first, then IF with word boundary)."""
        out = ""
        i = 0
        while i < len(f):
            m = re.match(rf"\b{name}\(", f[i:])
            prev_ok = (i == 0) or not (f[i-1].isalnum() or f[i-1] == "_")
            if m and prev_ok:
                j = i + len(name) + 1
                depth = 1
                args, cur = [], ""
                while depth:
                    ch = f[j]
                    if ch == "(":
                        depth += 1; cur += ch
                    elif ch == ")":
                        depth -= 1
                        if depth: cur += ch
                    elif ch == "," and depth == 1:
                        args.append(cur); cur = ""
                    else:
                        cur += ch
                    j += 1
                args.append(cur)
                args = [self.wrap_lazy(a, name, target) for a in args]
                if target == "_ifer":
                    out += f"_ifer(lambda:({args[0]}),lambda:({args[1]}))"
                else:
                    la = ",".join(f"lambda:({a})" for a in args[1:])
                    out += f"_if({args[0]}," + la + ")"
                i = j
            else:
                out += f[i]
                i += 1
        return out

    def run(self):
        results = {}
        if os.path.exists(STATE):
            d = json.load(open(STATE))
            self.regen = {(s, c): v for s, c, v in d}
        for sn in CALC_SHEETS:
            ws = wbf[sn]
            todo = []
            for row in ws.iter_rows(min_row=7):
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        todo.append((c.coordinate, c.value))
            todo.sort(key=lambda t: (len(re.match(r"[A-Z]+", t[0]).group()), re.match(r"[A-Z]+", t[0]).group(), int(re.search(r"\d+", t[0]).group())))
            progress, passes = True, 0
            while todo and progress and passes < 60:
                progress = False; passes += 1
                remaining = []
                for coord, f in todo:
                    try:
                        self.regen[(sn, coord)] = self.evaluate(sn, f)
                        progress = True
                    except KeyError:
                        remaining.append((coord, f))
                    except ValueError as e:
                        self.unsupported.append((sn, coord, f[:60], str(e)))
                        progress = True
                    except Exception as e:
                        self.unsupported.append((sn, coord, f[:60], f"eval error {e}"))
                        progress = True
                todo = remaining
            for coord, f in todo:
                self.unsupported.append((sn, coord, f[:60], "unresolved dependency"))
            # diff
            maxd, n, worst = 0.0, 0, None
            for (s, coord), v in self.regen.items():
                if s != sn:
                    continue
                cv = self.cached(sn, coord)
                if isinstance(v, (int, float)) and isinstance(cv, (int, float)):
                    d = abs(v - cv)
                    n += 1
                    if d > maxd:
                        maxd, worst = d, (coord, v, cv)
            results[sn] = (n, maxd, worst)
        return results

o = Oracle()
res = o.run()
json.dump([[s, c, v] for (s, c), v in o.regen.items() if isinstance(v,(int,float,bool))], open(STATE,"w"))
print("PLOVDIV FULL REGENERATION ORACLE")
ok = True
for sn, (n, maxd, worst) in res.items():
    status = "PASS" if maxd < 1e-9 else "FAIL"
    if maxd >= 1e-9:
        ok = False
    print(f"{sn}: {n} cells regenerated | max abs diff {maxd:.3e} | {status}"
          + (f" | worst {worst}" if worst and maxd >= 1e-9 else ""))
print(f"unsupported/skipped: {len(o.unsupported)}")
from collections import Counter
print(Counter(u[3].split(' ')[0] + ' ' + (u[3].split(' ')[1] if len(u[3].split(' '))>1 else '') for u in o.unsupported).most_common(8))
for u in o.unsupported[:8]:
    print("  ", u)
print("ORACLE:", "PASS" if ok else "FAIL")
