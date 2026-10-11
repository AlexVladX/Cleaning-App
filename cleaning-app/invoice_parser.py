"""
Parser determinist pentru facturile SimNettoyage.

Citeste direct pozitiile cuvintelor din PDF (pdfplumber), fara AI:
  - antet: client ("Facture pour"), numar factura, data emiterii, scadenta
  - tabelul "Détail des prestations": fiecare rand = o casa, cu toate datele de curatenie
  - totaluri: Subtotal, Ajustements, Total

Randurile tabelului sunt delimitate de dreptunghiurile de fundal (zebra) din PDF,
asa ca datele scrise pe mai multe linii (ex: 5 date intr-o singura celula) sau
numele rupte pe doua linii ("Appartement A" / "Bala") raman in randul corect.

Calculul interventiilor se face in cod (nu de AI), cu aceleasi reguli ca inainte:
  - FORFAIT (pret unitar >= FORFAIT_MIN_PRICE): fiecare data = o interventie
  - ORA (pret unitar <  FORFAIT_MIN_PRICE): orele se impart egal pe date
"""
import os
import re
import unicodedata
from functools import lru_cache
from datetime import datetime, timedelta

import pdfplumber

FORFAIT_MIN_PRICE = float(os.environ.get("FORFAIT_MIN_PRICE", "40"))

# anul poate fi scris cu 4 cifre (05/27/2026) sau cu 2 (08/20/26)
DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/(\d{4}|\d{2})$")
INVISIBLE = dict.fromkeys(map(ord, "​‌‍﻿  "), None)


class ParseError(Exception):
    """Factura nu are structura asteptata (alt sablon, scanata etc.)."""


# --------------------------------------------------------------------------- utilitare

def _clean(s):
    return (s or "").translate(INVISIBLE).strip()


def parse_money(s):
    """'€1,200.00' / '€30' / '1 200,50 €' / '-€15.00' -> float"""
    s = _clean(s).replace("€", "").replace("EUR", "").replace(" ", "")
    neg = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()")
    if not s:
        raise ValueError("suma goala")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):          # 1.200,50
            s = s.replace(".", "").replace(",", ".")
        else:                                     # 1,200.50
            s = s.replace(",", "")
    elif "," in s:
        head, tail = s.rsplit(",", 1)
        s = head.replace(",", "") + ("." + tail if len(tail) <= 2 else tail)
    v = float(s)
    return -v if neg else v


def parse_number(s):
    return float(_clean(s).replace(",", "."))


def is_money(s):
    s = _clean(s)
    if "€" not in s:
        return False
    try:
        parse_money(s)
        return True
    except ValueError:
        return False


def is_number(s):
    try:
        parse_number(s)
        return True
    except ValueError:
        return False


def parse_ddmmyyyy(s):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(_clean(s), fmt).date()
        except ValueError:
            pass
    return None


def parse_service_date(s, invoice_date=None):
    """
    Datele din 'Détail des prestations' sunt scrise MM/DD/YYYY (ex: 05/27/2026),
    iar antetul foloseste DD/MM/YYYY. Alegem interpretarea valida; daca ambele
    sunt valide, o preferam pe cea care cade inainte de data facturii.
    """
    s = _clean(s)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):          # ISO (ex. din rezerva AI)
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError:
            return None
    try:
        a, b, y = (int(x) for x in s.split("/"))
    except ValueError:
        return None
    if y < 100:
        y += 2000  # '26' -> 2026
    candidates = []
    for month, day in ((a, b), (b, a)):           # MM/DD intai (conventia sablonului)
        try:
            candidates.append(datetime(y, month, day).date())
        except ValueError:
            pass
    if not candidates:
        return None
    if invoice_date and len(candidates) == 2:
        for c in candidates:
            if invoice_date - timedelta(days=120) <= c <= invoice_date + timedelta(days=1):
                return c
    return candidates[0]


def _lines(words, tol=3.0):
    """Grupeaza cuvintele pe linii (dupa 'top')."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def _join(words):
    return " ".join(_clean(w["text"]) for l in _lines(words) for w in l).strip()


def _find_seq(words, *tokens):
    """Primul cuvant care incepe secventa `tokens` pe aceeasi linie."""
    toks = [t.lower() for t in tokens]
    for line in _lines(words):
        texts = [_clean(w["text"]).lower() for w in line]
        for i in range(len(texts) - len(toks) + 1):
            if texts[i:i + len(toks)] == toks:
                return line[i]
    return None


# --------------------------------------------------------------------------- antet

def _parse_header(pages_words, full_text):
    info = {"client": None, "client_adresa": None, "numar": None,
            "data_emitere": None, "scadenta": None}

    m = re.search(r"Soumis le\s+(\d{1,2}/\d{1,2}/(?:\d{4}|\d{2}))\b", full_text)
    if m:
        info["data_emitere"] = parse_ddmmyyyy(m.group(1))

    def value_below(words, label, x_right, y_limit):
        """Cuvintele de pe primul rand de sub o eticheta, in coloana ei."""
        cand = [w for w in words if label["bottom"] < w["top"] < min(label["bottom"] + 22, y_limit)
                and label["x0"] - 6 <= w["x0"] < x_right]
        if not cand:
            return []
        top = min(w["top"] for w in cand)
        return sorted((w for w in cand if abs(w["top"] - top) <= 3), key=lambda w: w["x0"])

    for words in pages_words:
        fp = _find_seq(words, "Facture", "pour")
        if not fp:
            continue
        fnum = _find_seq(words, "Facture", "#")
        fdue = _find_seq(words, "Date", "d'échéance")
        desc = _find_seq(words, "Description")
        y_end = desc["top"] - 2 if desc else fp["top"] + 110

        # coloana clientului se termina unde incepe urmatoarea eticheta de pe randul
        # "Facture pour" (ex. "Payable à" sau "Facture #"), oricare ar fi aspectul antetului
        same_line = [w for w in words if abs(w["top"] - fp["top"]) <= 3 and w["x0"] > fp["x1"] + 20]
        x_client = min((w["x0"] for w in same_line), default=(fnum["x0"] if fnum else fp["x0"] + 165)) - 2

        block = [w for w in words if fp["bottom"] + 1 < w["top"] < y_end]
        lines = _lines([w for w in block if w["x0"] < x_client])
        if lines:
            info["client"] = " ".join(_clean(w["text"]) for w in lines[0])
            info["client_adresa"] = ", ".join(
                " ".join(_clean(w["text"]) for w in l) for l in lines[1:]
            ) or None

        if fnum:
            x_right = fdue["x0"] - 2 if fdue and abs(fdue["top"] - fnum["top"]) <= 3 else fnum["x0"] + 100
            num = value_below(words, fnum, x_right, y_end)
            txt = " ".join(_clean(w["text"]) for w in num)
            m = re.search(r"[A-Z]{2,6}-\d{2}-\d{1,5}", txt)
            if m or (txt and re.search(r"\d", txt)):
                info["numar"] = m.group(0) if m else txt
        if fdue:
            due = value_below(words, fdue, fdue["x0"] + 150, y_end)
            d = next((w for w in due if DATE_RE.match(_clean(w["text"]))), None)
            if d:
                info["scadenta"] = parse_ddmmyyyy(d["text"])
        break

    if not info["numar"]:
        m = re.search(r"\b([A-Z]{2,6}-\d{2}-\d{1,5})\b", full_text)
        if m:
            info["numar"] = m.group(1)
    return info


# --------------------------------------------------------------------------- totaluri

def _parse_totals(words, y_from=0):
    """Subtotal / Ajustements / total final (suma fara eticheta de sub Ajustements)."""
    res = {"subtotal": None, "ajustari": 0.0, "total": None}
    ws = [w for w in words if w["top"] >= y_from]
    st = _find_seq(ws, "Subtotal")
    if not st:
        return res
    lines = _lines(ws)

    def money_on_line(anchor):
        for l in lines:
            if any(w is anchor for w in l):
                vals = [w for w in l if w["x0"] > anchor["x0"] and is_money(w["text"])]
                return parse_money(vals[0]["text"]) if vals else None
        return None

    res["subtotal"] = money_on_line(st)
    aj = _find_seq([w for w in ws if w["top"] >= st["top"] - 1], "Ajustements")
    if aj:
        v = money_on_line(aj)
        res["ajustari"] = v or 0.0
        below = [w for w in ws if w["top"] > aj["bottom"] and is_money(w["text"])]
        if below:
            res["total"] = parse_money(sorted(below, key=lambda w: w["top"])[0]["text"])
    if res["total"] is None and res["subtotal"] is not None:
        res["total"] = round(res["subtotal"] + res["ajustari"], 2)
    return res


# --------------------------------------------------------------------------- tabel detaliu

def _cells_from_anchors(table_words, total_right_edge, x_numeric, max_gap=14.0):
    """
    Rezerva cand PDF-ul nu are fundal pe randuri.
    Ancora unui rand = suma din coloana 'Prix total'. Textul dintr-o celula
    (nume pe 2 linii, mai multe date) e centrat vertical pe ancora, deci grupam
    liniile apropiate pe fiecare coloana in blocuri si legam blocul de ancora
    cea mai apropiata de centrul lui.
    """
    def mid(w):
        return (w["top"] + w["bottom"]) / 2

    anchors = sorted((w for w in table_words
                      if is_money(w["text"]) and abs(w["x1"] - total_right_edge) < 8),
                     key=mid)
    if not anchors:
        return []
    cells = [[a] for a in anchors]

    def nearest(y):
        return min(range(len(anchors)), key=lambda i: abs(mid(anchors[i]) - y))

    rest = [w for w in table_words if not any(w is a for a in anchors)]
    numeric = [w for w in rest if w["x0"] >= x_numeric]
    text = [w for w in rest if w["x0"] < x_numeric]
    for w in numeric:
        cells[nearest(mid(w))].append(w)

    # coloane de text: grupam dupa x0 apropiat
    columns = []
    for w in sorted(text, key=lambda w: w["x0"]):
        if columns and w["x0"] - columns[-1][0] < 30:
            columns[-1][1].append(w)
        else:
            columns.append((w["x0"], [w]))
    for _, col in columns:
        blocks = []
        for w in sorted(col, key=lambda w: w["top"]):
            if blocks and w["top"] - blocks[-1][-1]["top"] <= max_gap:
                blocks[-1].append(w)
            else:
                blocks.append([w])
        for b in blocks:
            center = (b[0]["top"] + b[-1]["bottom"]) / 2
            cells[nearest(center)].extend(b)
    return cells


def _detail_rows(page):
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    title = _find_seq(words, "Détail", "des", "prestations")
    if not title:
        return None, words, 0

    below = [w for w in words if w["top"] > title["bottom"]]
    h_nom = _find_seq(below, "Nom")
    h_type = _find_seq(below, "Type")
    h_qty = next((w for w in below if _clean(w["text"]).lower().startswith("qty")), None)
    h_date = _find_seq(below, "Date", "de")
    prix = [w for w in below if _clean(w["text"]) == "Prix"]
    h_unit = _find_seq(below, "unitaire")
    h_total = _find_seq(below, "total")
    if not (h_nom and h_qty and h_total and len(prix) >= 2):
        raise ParseError("Antetul tabelului 'Détail des prestations' nu a fost gasit")

    header_words = [w for w in below if w["top"] <= h_nom["bottom"] + 10
                    and w["top"] >= (h_date["top"] if h_date else h_nom["top"]) - 1]
    y_start = max(w["bottom"] for w in header_words)
    sub = _find_seq([w for w in below if w["top"] > y_start], "Subtotal")
    y_end = sub["top"] if sub else page.height

    # coloane: text aliniat la stanga, sume aliniate la dreapta
    x_type = (h_type["x0"] - 20) if h_type else h_nom["x0"] + 85
    x_qty = h_qty["x0"] - 3
    right_edges = {
        "qty": h_qty["x1"],
        "pu": (h_unit or prix[0])["x1"],
        "total": h_total["x1"],
    }

    table_words = [w for w in below if y_start < w["top"] < y_end]

    # randurile = dreptunghiurile de fundal din zona tabelului
    rects = []
    for r in page.rects:
        if r["width"] > 250 and 8 < r["height"] < 200 \
                and r["top"] >= y_start - 4 and r["bottom"] <= y_end + 2 \
                and r["x0"] <= h_nom["x0"]:
            rects.append((r["top"], r["bottom"]))
    rects.sort()
    bands = []
    for top, bottom in rects:
        if bands and top < bands[-1][1] - 2:      # suprapunere -> acelasi rand
            bands[-1] = (bands[-1][0], max(bands[-1][1], bottom))
        else:
            bands.append((top, bottom))

    def in_band(w):
        mid = (w["top"] + w["bottom"]) / 2
        return any(top - 1 <= mid <= bottom + 1 for top, bottom in bands)

    # Fundalurile acopera TOATE randurile doar la unele facturi; la altele doar randurile
    # gri au dreptunghi (cele albe nu). Daca vreo suma din coloana 'Prix total' ramane
    # in afara fundalurilor, folosim metoda pe ancore, care nu depinde de fundal.
    totals_col = [w for w in table_words
                  if is_money(w["text"]) and abs(w["x1"] - right_edges["total"]) < 8]
    if bands and all(in_band(w) for w in totals_col):
        cells = [[w for w in table_words if top - 1 <= (w["top"] + w["bottom"]) / 2 <= bottom + 1]
                 for top, bottom in bands]
    else:
        cells = _cells_from_anchors(table_words, right_edges["total"], x_qty - 40)

    rows = []
    for cell in cells:
        if not cell:
            continue
        nom, typ, dates, nums = [], [], [], {"qty": [], "pu": [], "total": []}
        for w in cell:
            t = _clean(w["text"])
            if DATE_RE.match(t):
                dates.append(w)
            elif w["x0"] >= x_qty - 40 and (is_money(t) or is_number(t)):
                col = min(right_edges, key=lambda k: abs(right_edges[k] - w["x1"]))
                nums[col].append(w)
            elif w["x0"] < x_type:
                nom.append(w)
            elif w["x0"] < x_qty:
                typ.append(w)  # include si cuvintele lungi ("Intermédiaire") care trec in coloana datei
        if not nom and not dates:
            continue                               # rand gol (doar €0.00)

        def first(col, conv):
            return conv(nums[col][0]["text"]) if nums[col] else None

        rows.append({
            "nom": re.sub(r"\s+", " ", _join(nom)),
            "type": _join(typ) or None,
            "dates_raw": [_clean(w["text"]) for w in sorted(dates, key=lambda w: w["top"])],
            "qty": first("qty", parse_money if nums["qty"] and "€" in nums["qty"][0]["text"] else parse_number),
            "pu": first("pu", parse_money),
            "total": first("total", parse_money),
        })
    return rows, words, y_start


# --------------------------------------------------------------------------- reguli de calcul

DEFAULT_RULES = {
    "prag_forfait": FORFAIT_MIN_PRICE,      # sub acest pret unitar -> la ora (rezerva)
    "tipuri_la_ora": ["intermediaire"],     # tipuri de menaj facturate mereu la ora
    "clienti_la_ora": [],                   # clienti facturati mereu la ora
}

TIP_LABEL = {"standard": "Standard", "intermediaire": "Intermédiaire", "ouverture": "Ouverture"}


@lru_cache(maxsize=20000)
def norm_text(s):
    """'Intermédiaire ' -> 'intermediaire'; fara accente, litere mici, spatii simple."""
    s = unicodedata.normalize("NFKD", _clean(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c)).casefold()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def canonical_type(t):
    n = norm_text(t)
    if not n or "standard" in n:
        return "standard"
    if "interm" in n:
        return "intermediaire"
    if "ouvert" in n or "opening" in n:
        return "ouverture"
    return n


def billing_mode(row, client, rules):
    """('heure'|'forfait', motiv). Ordinea: client la ora > tip la ora > pret sub prag."""
    rules = {**DEFAULT_RULES, **(rules or {})}
    c = norm_text(client)
    for name in rules.get("clienti_la_ora") or []:
        if norm_text(name) and norm_text(name) in c:
            return "heure", "client"
    tip = canonical_type(row.get("type"))
    if tip in {canonical_type(t) for t in rules.get("tipuri_la_ora") or []}:
        return "heure", "tip"
    pu = row.get("pu")
    if pu is not None and pu < float(rules.get("prag_forfait") or FORFAIT_MIN_PRICE):
        return "heure", "pret"
    return "forfait", None


# --------------------------------------------------------------------------- interventii

def compute_interventions(rows, header, rules=None):
    """Transforma randurile facturii in interventii individuale + avertismente."""
    interventii, warnings = [], []
    inv_date = header.get("data_emitere")
    if isinstance(inv_date, str):
        inv_date = parse_ddmmyyyy(inv_date)

    for r in rows:
        nom = r.get("nom") or "Necunoscut"
        qty, pu, total = r.get("qty"), r.get("pu"), r.get("total")
        if total is None and qty is not None and pu is not None:
            total = round(qty * pu, 2)
        if total is None:
            warnings.append(f"{nom}: lipseste suma totala pe rand")
            continue
        if pu is None and qty:
            pu = round(total / qty, 2)
        if qty is None and pu:
            qty = round(total / pu, 2)
        if qty is not None and pu is not None and abs(qty * pu - total) > 0.01:
            warnings.append(f"{nom}: {qty} x {pu} = {qty * pu:.2f} ≠ total {total:.2f}")

        dates = []
        for d in r.get("dates_raw") or []:
            pd = parse_service_date(d, inv_date)
            if pd:
                dates.append(pd)
            else:
                warnings.append(f"{nom}: data invalida '{d}'")

        mode, motiv = billing_mode({**r, "pu": pu}, header.get("client"), rules)
        tip = canonical_type(r.get("type"))
        n = len(dates)
        if n == 0:
            n = int(round(qty)) if (mode == "forfait" and qty and float(qty).is_integer()) else 1
            dates = [None] * n
            warnings.append(f"{nom}: nicio data de interventie pe factura")
        elif mode == "forfait" and qty is not None and abs(qty - n) > 0.01:
            warnings.append(f"{nom}: cantitate {qty:g} dar {n} date — verificati")

        per = round(total / n, 2)
        for i, d in enumerate(sorted(dates, key=lambda x: (x is None, x))):
            suma = per if i < n - 1 else round(total - per * (n - 1), 2)
            ore = round(qty / n, 2) if (mode == "heure" and qty) else None
            interventii.append({
                "proprietate": nom,
                "client": header.get("client"),
                "data": d.strftime("%d/%m/%Y") if d else "",
                "data_iso": d.isoformat() if d else "",
                "tip_serviciu": r.get("type") or "",
                "tip_menaj": tip,
                "tip_facturare": mode,
                "motiv_ora": motiv,
                "qty": ore if mode == "heure" else 1,
                "ore": ore,
                "pret_unitar": pu,
                "suma": suma,
                "moneda": "EUR",
                "factura": header.get("numar"),
            })
    return interventii, warnings


# --------------------------------------------------------------------------- intrare

def layout_text(pdf_source):
    """Text cu aspectul paginii pastrat (folosit doar pentru rezerva AI)."""
    if hasattr(pdf_source, "seek"):
        pdf_source.seek(0)
    with pdfplumber.open(pdf_source) as pdf:
        return "\n\n".join((p.extract_text(layout=True) or "") for p in pdf.pages).strip()


def invoice_warnings(rows, totals):
    """Avertismente la nivel de factura (nu depind de regulile de calcul)."""
    warnings = []
    rows_sum = round(sum(r.get("total") or 0 for r in rows), 2)
    if totals.get("subtotal") is not None and abs(rows_sum - totals["subtotal"]) > 0.01:
        warnings.append(f"Suma randurilor {rows_sum:.2f} ≠ Subtotal {totals['subtotal']:.2f}")
    if totals.get("ajustari"):
        warnings.append(f"Ajustari de {totals['ajustari']:.2f} € — nu sunt repartizate pe proprietati")
    return warnings


def parse_invoice(pdf_source, rules=None):
    """
    pdf_source: cale sau file-like (BytesIO).
    Intoarce dict cu antet, randuri, totaluri, interventii, avertismente.
    Ridica ParseError daca factura nu are sablonul SimNettoyage.
    """
    if hasattr(pdf_source, "seek"):
        pdf_source.seek(0)
    with pdfplumber.open(pdf_source) as pdf:
        pages_words = [p.extract_words() for p in pdf.pages]
        full_text = "\n".join(_clean(p.extract_text() or "") for p in pdf.pages)
        if len(full_text.strip()) < 20:
            raise ParseError("PDF fara text (probabil scanat)")

        header = _parse_header(pages_words, full_text)

        rows, totals = None, None
        for page in pdf.pages:
            found, words, y0 = _detail_rows(page)
            if found is not None:
                rows = found
                totals = _parse_totals(words, y0)
                break
        if rows is None:
            raise ParseError("Tabelul 'Détail des prestations' lipseste")
        if not rows:
            raise ParseError("Tabelul 'Détail des prestations' nu contine randuri")
        if totals is None or totals.get("subtotal") is None:
            totals = _parse_totals(pages_words[0])

    if not header.get("client"):
        raise ParseError("Clientul ('Facture pour') nu a fost gasit")

    interventii, calc_warnings = compute_interventions(rows, header, rules)
    inv_warnings = invoice_warnings(rows, totals)

    return {
        "metoda": "parser",
        "antet": {
            **header,
            "data_emitere": header["data_emitere"].strftime("%d/%m/%Y") if header.get("data_emitere") else None,
            "scadenta": header["scadenta"].strftime("%d/%m/%Y") if header.get("scadenta") else None,
        },
        "randuri": rows,
        "totaluri": totals,
        "interventii": interventii,
        "avertismente": calc_warnings + inv_warnings,
        "avertismente_factura": inv_warnings,
    }


if __name__ == "__main__":
    import json
    import sys
    for path in sys.argv[1:]:
        print("=" * 70, path)
        try:
            print(json.dumps(parse_invoice(path), ensure_ascii=False, indent=2, default=str))
        except ParseError as e:
            print("ParseError:", e)
