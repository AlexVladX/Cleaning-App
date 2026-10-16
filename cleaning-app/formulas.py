"""
Indicatori proprii: formule simple scrise de utilizator ("€ / oră = incasari_fara_inter / ore_lucru").

Se accepta doar: numere, variabilele de mai jos, + - * / ( ) si functiile min, max, round, abs.
Nimic altceva nu se executa (expresia e verificata pe arborele de sintaxa, fara eval).
Aceeasi logica exista si in pagina (JS), pentru calculul pe loc.
"""
import ast
import re

VARS = {
    "interventii": "toate intervențiile",
    "menaje": "menaje complete (Standard / Ouverture)",
    "interm": "menaje Intermédiaire",
    "incasari": "încasări totale (€)",
    "incasari_fara_inter": "încasări fără Intermédiaire (€)",
    "incasari_inter": "încasări din Intermédiaire (€)",
    "ore": "ore estimate / menaj (1 persoană)",
    "ore_lucru": "ore de lucru = menaje × ore",
    "ore_inter": "ore facturate la Intermédiaire",
    "ore_facturate": "toate orele facturate la oră",
    "distanta": "km de la bază",
    "timp_drum": "minute de la bază",
    "km_drum": "km de drum în perioadă (traseele zilnice)",
    "motorina": "motorina estimată (€)",
    "profit": "profit estimat (€)",
    "pe_luna": "intervenții pe lună",
    "zile_ultima": "zile de la ultima intervenție",
    "case": "numărul de case (1 la o casă)",
}
FUNCS = {"min": min, "max": max, "round": round, "abs": abs}
MAX_FORMULAS = 20


class FormulaError(ValueError):
    pass


def _check(node):
    if isinstance(node, ast.Expression):
        return _check(node.body)
    if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
        _check(node.left)
        _check(node.right)
        return
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
        _check(node.operand)
        return
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return
    if isinstance(node, ast.Name):
        if node.id not in VARS:
            raise FormulaError(f"nu cunosc „{node.id}”")
        return
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in FUNCS and not node.keywords:
        for a in node.args:
            _check(a)
        return
    raise FormulaError("se acceptă doar numere, variabile, + - * / ( ) și min, max, round, abs")


def compile_expr(expr):
    expr = re.sub(r"(\d),(\d)", r"\1.\2", (expr or "").strip())   # 2,5 -> 2.5
    if not expr:
        raise FormulaError("expresia e goală")
    if len(expr) > 300:
        raise FormulaError("expresia e prea lungă")
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError:
        raise FormulaError("scriere greșită (verifică parantezele și semnele)")
    _check(tree)
    return compile(tree, "<formula>", "eval")


def evaluate(code, values):
    """Valori lipsa (None) sau impartire la 0 -> None (afisat ca „—”)."""
    if any(values.get(v) is None for v in code.co_names if v in VARS):
        return None
    try:
        r = eval(code, {"__builtins__": {}}, {**FUNCS, **values})  # noqa: S307 - arbore verificat in _check
    except (ZeroDivisionError, TypeError, ValueError, OverflowError):
        return None
    return round(r, 2) if isinstance(r, (int, float)) else None


def parse_lines(text):
    """'Nume = expresie' pe fiecare rand -> ([{nume, expr}], [erori])."""
    out, errors = [], []
    for i, line in enumerate(str(text or "").splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            errors.append(f"rândul {i}: lipsește „=” (ex. € / oră = incasari_fara_inter / ore_lucru)")
            continue
        nume, expr = line.split("=", 1)
        nume = nume.strip()[:40]
        if not nume:
            errors.append(f"rândul {i}: lipsește numele coloanei")
            continue
        try:
            compile_expr(expr)
        except FormulaError as e:
            errors.append(f"rândul {i} („{nume}”): {e}")
            continue
        out.append({"nume": nume, "expr": expr.strip()})
    if len(out) > MAX_FORMULAS:
        errors.append(f"cel mult {MAX_FORMULAS} formule")
        out = out[:MAX_FORMULAS]
    return out, errors


def house_values(p):
    """Variabilele unei case (aceleasi nume ca in pagina)."""
    t = p.get("tipuri") or {}
    return {
        "interventii": p.get("interventii") or 0,
        "menaje": p.get("menaje_complete") or 0,
        "interm": p.get("interm_n") or t.get("intermediaire", 0),
        "incasari": p.get("total") or 0,
        "incasari_fara_inter": p.get("venit_complet"),
        "incasari_inter": p.get("interm_venit") or 0,
        "ore": p.get("ore_est"),
        "ore_lucru": p.get("ore_lucru"),
        "ore_inter": p.get("interm_ore") or 0,
        "ore_facturate": p.get("ore") or 0,
        "distanta": p.get("distanta_km"),
        "timp_drum": p.get("timp_min"),
        "km_drum": p.get("km_drum"),
        "motorina": p.get("motorina_est"),
        "profit": p.get("profit_est"),
        "pe_luna": p.get("frecventa_luna"),
        "zile_ultima": p.get("zile_de_la_ultima"),
        "case": 1,
    }


DEFAULT_FORMULAS = [
    {"nume": "€ / oră", "expr": "incasari_fara_inter / ore_lucru"},
    {"nume": "€ / menaj", "expr": "incasari_fara_inter / menaje"},
    {"nume": "Marjă %", "expr": "profit / incasari * 100"},
]
