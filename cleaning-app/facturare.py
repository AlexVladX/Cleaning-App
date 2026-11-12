"""
Generatorul de facturi SimNettoyage.

Totul vine din Sheet-ul de facturare al utilizatorului (nimic nu e scris in cod: repo-ul e public):
  - primele doua foi  = cele doua formate de factura (emitent, cont bancar, nota de plata)
  - "Clients"         = datele de facturare ale clientilor
  - "Case /m2"        = tarife pe casa (Standard / Ouverture / tarif orar Intermédiaire) + tarife speciale
  - "Case /h"         = case facturate la ora (ore alocate x tarif orar)

PDF-ul are acelasi aspect si aceeasi structura ca facturile de pana acum ("Facture pour", "Facture #",
"Détail des prestations", Subtotal / Ajustements), deci cititorul de facturi il intelege fara schimbari.
Datele menajelor sunt scrise ZZ/LL/AAAA; ca sa nu fie citite pe dos (LL/ZZ), PDF-ul poarta un semn
in proprietatile lui (Creator = MARKER), pe care cititorul il recunoaste.
"""
import io
import os
import re
import unicodedata
from datetime import date, datetime, timedelta

MARKER = "SimNettoyage-Facturare"          # in proprietatile PDF-ului: datele sunt ZZ/LL/AAAA
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "facturare")    # Roboto-*.ttf + logo.png (optionale: fara ele -> Helvetica, fara logo)

TYPES = ("Standard", "Ouverture", "Intermédiaire", "Autre")


# --------------------------------------------------------------------------- utilitare

def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().casefold()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def money(s):
    """'€1,200' / '25€/h' / '220' / '€1.3' -> float sau None."""
    s = str(s or "").replace(" ", "").replace(" ", "")
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def num(s):
    try:
        return float(str(s).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None


def cell(row, i):
    return str(row[i]).strip() if i < len(row) and row[i] is not None else ""


# --------------------------------------------------------------------------- citirea Sheet-ului

def read_book(svc, sheet_id):
    """Citeste foile necesare. svc = serviciul Google Sheets. Intoarce {titlu foaie: randuri}."""
    meta = svc.spreadsheets().get(spreadsheetId=sheet_id, fields="sheets.properties.title").execute(num_retries=3)
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    want = titles[:2] + [t for t in titles[2:] if norm(t) in ("clients", "clienti", "case m2", "case h")]
    ranges = ["'" + t.replace("'", "''") + "'!A1:N400" for t in want]
    res = svc.spreadsheets().values().batchGet(spreadsheetId=sheet_id, ranges=ranges).execute(num_retries=3)
    out = {"_titles": titles}
    for t, vr in zip(want, res.get("valueRanges", [])):
        out[t] = vr.get("values", [])
    return out


def _issuer(rows):
    """Blocul emitentului dintr-o foaie-format: nume, adresa, siret, telefon, cont bancar, nota, numarul facturii."""
    col = [cell(r, 1) for r in rows]
    lines = [c for c in col[:14] if c]
    bank = [c for c in lines if re.match(r"(?i)(titulaire|iban|rib|bic)\b", c)]
    head = [c for c in lines if c not in bank and norm(c) not in ("facture",) and not c.lower().startswith("soumis")]
    em = {"nume": head[0] if head else "", "linii": head[1:4], "banca": bank, "nota": "", "numar": "", "payable": False}
    for c in bank:
        m = re.match(r"(?i)titulaire[^:]*:\s*(.+)", c)
        if m:
            em["titular"] = m.group(1).strip()
        m = re.match(r"(?i)iban\s*:\s*(.+)", c)
        if m:
            em["iban"] = re.sub(r"\s+", "", m.group(1))
    for r in rows:
        for i, c in enumerate(r):
            c = str(c or "").strip()
            if c.lower().startswith("note:") and not em["nota"]:
                em["nota"] = c
            if norm(c) == "payable a":
                em["payable"] = True
    for ri, r in enumerate(rows):
        for i, c in enumerate(r):
            if str(c or "").strip() == "Facture #" and ri + 1 < len(rows):
                em["numar"] = cell(rows[ri + 1], i)
    return em


def parse_book(book):
    """Randurile brute -> {formate, clienti, grupuri (tarife), speciale}."""
    titles = book.get("_titles") or []
    by = {norm(t): book.get(t) or [] for t in book if t != "_titles"}
    formate = []
    for t in titles[:2]:
        em = _issuer(book.get(t) or [])
        if em["nume"]:
            em["id"] = "f%d" % (len(formate) + 1)
            formate.append(em)

    clienti, grup = [], ""
    for r in (by.get("clients") or by.get("clienti") or [])[1:]:
        if cell(r, 0):
            grup = cell(r, 0)
        if not cell(r, 1):
            continue
        clienti.append({"nume": cell(r, 1), "adresa": cell(r, 2), "siret": cell(r, 3),
                        "grup": grup, "casa": cell(r, 4), "email": cell(r, 5)})

    grupuri, speciale = {}, []
    rows = by.get("case m2") or []
    g = ""
    for ri, r in enumerate(rows[1:], start=2):
        if cell(r, 0):
            g = cell(r, 0)
        # tarife speciale (coloanele L-N): "KARINE | Repassage | 25€/h"
        if cell(r, 12) and money(cell(r, 13)) is not None and not money(cell(r, 12)):
            speciale.append({"client": cell(r, 11), "nume": cell(r, 12), "pret": money(cell(r, 13)),
                             "unitate": re.sub(r"^[\d.,€\s/]+", "", cell(r, 13)) or "u"})
        if not cell(r, 1) or not g:
            continue
        casa = {"nume": re.sub(r"\s+", " ", cell(r, 1)), "mod": "forfait",
                "standard": money(cell(r, 2)), "ouverture": money(cell(r, 3)) or None,
                "tarif_ora": money(cell(r, 4)), "variante": []}
        # pret alternativ pe acelasi rand, in coloanele de dupa m2: "€330 | cu intermediar"
        for k in range(9, 13):
            a, b = cell(r, k), cell(r, k + 1)
            if re.fullmatch(r"€?\s*\d[\d.,]*\s*€?", a) and b and not re.search(r"\d", b) and not b.startswith("http"):
                casa["variante"].append({"pret": money(a), "eticheta": b})
                break
        grupuri.setdefault(g, []).append(casa)
    g = ""
    for r in (by.get("case h") or [])[1:]:
        if cell(r, 0):
            g = cell(r, 0)
        if not cell(r, 1) or not g:
            continue
        grupuri.setdefault(g, []).append({
            "nume": re.sub(r"\s+", " ", cell(r, 1)), "mod": "ora", "ore_standard": num(cell(r, 2)),
            "ore_interm": num(cell(r, 3)) or None, "tarif_ora": money(cell(r, 4)), "variante": []})
    return {"formate": formate, "clienti": clienti,
            "grupuri": [{"nume": k, "case": v} for k, v in grupuri.items()], "speciale": speciale}


def guess_group(client, grupuri):
    """Lista de tarife care se potriveste cel mai bine cu un client (dupa nume, porecla sau casa lui)."""
    best, score = None, 0
    cn, cg, cc = norm(client.get("nume")), norm(client.get("grup")), norm(client.get("casa"))
    for g in grupuri:
        gn = norm(g["nume"])
        s = 0
        if gn and (gn in cn or cn in gn):
            s = 5
        elif gn and cg and (gn in cg or cg in gn or gn[:4] == cg[:4]):
            s = 3
        else:
            gw, cw = set(gn.split()), set(cn.split()) | set(cg.split())
            common = {w for w in gw & cw if len(w) > 2 and w not in ("villa", "villas", "les", "sarl", "sas")}
            if common:
                s = 2
        if cc and any(cc in norm(c["nume"]) or norm(c["nume"]) in cc for c in g["case"]):
            s = max(s, 4)
        if s > score:
            best, score = g["nume"], s
    return best


# --------------------------------------------------------------------------- numerotare

def next_number(fmt_number, known, when):
    """
    Urmatorul numar, dupa modelul formatului ('NET-26-021' / 'SIMNET-016') si numerele deja folosite.
    Modelul cu an (NET-AA-NNN) reia de la 001 in fiecare an.
    """
    m = re.match(r"^(.*?)(\d+)$", fmt_number or "")
    if not m:
        return ""
    prefix, width = m.group(1), len(m.group(2))
    ym = re.match(r"^(.*?-)(\d{2})-$", prefix)
    if ym:
        prefix = "%s%02d-" % (ym.group(1), when.year % 100)
    last = 0
    for k in list(known) + [fmt_number]:
        mm = re.match(r"^" + re.escape(prefix) + r"(\d+)$", (k or "").strip())
        if mm:
            last = max(last, int(mm.group(1)))
    return "%s%0*d" % (prefix, width, last + 1)


# --------------------------------------------------------------------------- randurile facturii

def build_rows(lines):
    """
    Liniile din formular -> randurile facturii. Valideaza si calculeaza totalurile.
    line = {nom, type, dates: ['2026-11-03', ...], qty, pu, desc?}
    """
    rows, errs = [], []
    for i, ln in enumerate(lines or [], start=1):
        nom = re.sub(r"\s+", " ", str(ln.get("nom") or "")).strip()
        if not nom:
            continue
        typ = str(ln.get("type") or "Standard").strip()
        dates = []
        for d in ln.get("dates") or []:
            try:
                dates.append(datetime.strptime(str(d)[:10], "%Y-%m-%d").date())
            except ValueError:
                errs.append(f"Rândul {i} ({nom}): dată greșită „{d}”")
        dates = sorted(set(dates))
        qty, pu = num(ln.get("qty")), num(ln.get("pu"))
        if pu is None:
            errs.append(f"Rândul {i} ({nom}): lipsește prețul")
            continue
        if qty is None:
            qty = float(len(dates))
        if qty <= 0:
            errs.append(f"Rândul {i} ({nom}): cantitatea trebuie să fie peste 0")
            continue
        if not dates:
            errs.append(f"Rândul {i} ({nom}): lipsește data")
        rows.append({"nom": nom, "type": "" if typ == "Autre" else typ, "dates": dates,
                     "qty": qty, "pu": pu, "total": round(qty * pu, 2),
                     "desc": (str(ln.get("desc") or "").strip() or None), "casa": not ln.get("liber")})
    return rows, errs


def adjustment(subtotal, mod, val):
    """Reducere / majorare: mod 'eur' (suma, cu semn) sau 'pct' (procent, cu semn)."""
    v = num(val) or 0.0
    if not v:
        return 0.0
    return round(subtotal * v / 100.0, 2) if mod == "pct" else round(v, 2)


# --------------------------------------------------------------------------- PDF

DARK = (0.1765, 0.3059, 0.3216)
PINK = (0.8784, 0.1059, 0.5176)
GRAY = (0.4, 0.4, 0.4)
INK = (0.2627, 0.2627, 0.2627)
ZEBRA = (0.9529, 0.9529, 0.9529)
RULE = (0.7176, 0.7176, 0.7176)
PAGE_W, PAGE_H = 595.92, 842.88
X0, X_QTY_R, X_PU_R, X_TOT_R = 87.4, 351.6, 431.9, 507.1
ROW_X0, ROW_X1 = 84.5, 510.0


def fmt_money(v, decimals=2):
    s = "{:,.{d}f}".format(abs(v), d=decimals)
    return ("-" if v < -0.004 else "") + "€" + s


def fmt_pu(v):
    return fmt_money(v, 0 if abs(v - round(v)) < 0.005 else 2)


def fmt_qty(v, short=False):
    if abs(v - round(v)) < 0.005:
        return "%d" % round(v) if short else "%.1f" % v
    return ("%.2f" % v).rstrip("0")


def fmt_date(d):
    return d.strftime("%d/%m/%Y")


def _fonts():
    """Roboto (ca in facturile existente) daca fisierele sunt in folderul 'facturare'; altfel Helvetica."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    names = {"R": "Helvetica", "B": "Helvetica-Bold", "I": "Helvetica-Oblique"}
    for key, fn in (("R", "Roboto-Regular"), ("B", "Roboto-Bold"), ("I", "Roboto-Italic")):
        path = os.path.join(ASSETS, fn + ".ttf")
        if os.path.exists(path):
            try:
                if fn not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(fn, path))
                names[key] = fn
            except Exception:  # noqa: BLE001
                pass
    return names


def _wrap(text, font, size, width):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    out = []
    for para in str(text or "").split("\n"):
        cur = ""
        for w in para.split():
            t = (cur + " " + w).strip()
            if cur and stringWidth(t, font, size) > width:
                out.append(cur)
                cur = w
            else:
                cur = t
        out.append(cur)
    return [l for l in out if l] or [""]


class _Page:
    """Desen cu coordonate 'de sus' (ca in PDF-urile existente): top = marginea de sus a textului."""

    def __init__(self, c, height, fonts):
        self.c, self.h, self.f = c, height, fonts

    def text(self, x, top, s, size=9.5, font="R", color=GRAY, align="l"):
        c = self.c
        c.setFont(self.f[font], size)
        c.setFillColorRGB(*color)
        y = self.h - top - size * 0.93
        if align == "r":
            c.drawRightString(x, y, s)
        elif align == "c":
            c.drawCentredString(x, y, s)
        else:
            c.drawString(x, y, s)

    def rect(self, x0, top, x1, bottom, color):
        self.c.setFillColorRGB(*color)
        self.c.rect(x0, self.h - bottom, x1 - x0, bottom - top, stroke=0, fill=1)

    def line(self, x0, x1, top, color=RULE, width=0.6):
        self.c.setStrokeColorRGB(*color)
        self.c.setLineWidth(width)
        self.c.line(x0, self.h - top, x1, self.h - top)


def _totals_block(p, y, subtotal, ajust, note_lines=None):
    """Linia de sub tabel + Subtotal / Ajustements / total. Intoarce y-ul de dupa."""
    p.line(ROW_X0, ROW_X1, y)
    p.text(X_PU_R, y + 9.5, "Subtotal", color=DARK, align="r")
    p.text(X_TOT_R, y + 9.5, fmt_money(subtotal), font="B", color=(0, 0, 0), align="r")
    p.text(X_PU_R, y + 28, "Ajustements", color=DARK, align="r")
    p.text(X_TOT_R, y + 28, fmt_money(ajust), font="B", color=INK, align="r")
    p.text(X_TOT_R, y + 47, fmt_money(subtotal + ajust), size=17, font="B", color=PINK, align="r")
    ny = y + 6
    for i, (txt, font) in enumerate(note_lines or []):
        p.text(X0, ny, txt, size=7.5, font=font, color=(0.55, 0.55, 0.55))
        ny += 10
    return max(y + 80, ny + 10)


def build_pdf(doc):
    """
    doc = {emitent, client{nume, adresa, siret}, numar, data(date), scadenta(date), rows, ajustari,
           eticheta_ajustare?, prefix_descriere?}
    Intoarce bytes. Pagina 1 = rezumat, pagina 2 = "Détail des prestations" (tot tabelul pe o pagina:
    cititorul citeste tabelul de pe o singura pagina, deci la facturi lungi pagina se lungeste).
    """
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase.pdfmetrics import stringWidth
    f = _fonts()
    em, cl, rows = doc["emitent"], doc["client"], doc["rows"]
    subtotal = round(sum(r["total"] for r in rows), 2)
    ajust = round(doc.get("ajustari") or 0.0, 2)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    c.setCreator(MARKER + " (dates JJ/MM/AAAA)")
    c.setTitle("Facture %s" % doc["numar"])
    c.setAuthor(em.get("nume") or "")

    # ---------------- pagina 1
    ROW_H = 18.6
    addr_lines = _wrap(cl.get("adresa"), f["R"], 9.5, 158)
    payable = bool(em.get("payable"))
    head_bottom = 335.9 + 16.4 + 11.5 * len(addr_lines) + (20 if cl.get("siret") else 6)
    if payable:
        head_bottom = max(head_bottom, 400)
    y_rule = max(407.2, head_bottom + 12)
    y_rows = y_rule + 43.3
    note = em.get("nota") or ""
    note_lines = []
    for k, para in enumerate(note.split("\n")):
        for ln in _wrap(para, f["R"], 7.5, 235):
            note_lines.append((ln, "I" if k > 0 else "R"))
    if ajust and doc.get("eticheta_ajustare"):
        note_lines.append((doc["eticheta_ajustare"], "B"))
    need = y_rows + ROW_H * len(rows) + max(90, 20 + 10 * len(note_lines)) + 40
    h1 = max(PAGE_H, need)
    c.setPageSize((PAGE_W, h1))
    p = _Page(c, h1, f)

    p.rect(49.4, 53.0, 545.1, 59.4, DARK)
    logo = os.path.join(ASSETS, "logo.png")
    if os.path.exists(logo):
        try:
            c.drawImage(logo, 385, h1 - 177, width=124, height=123, mask="auto")
        except Exception:  # noqa: BLE001
            pass
    p.text(X0, 76.5, em.get("nume") or "", size=19.1, font="B", color=DARK)
    y = 100.9
    for ln in em.get("linii") or []:
        for part in _wrap(ln, f["R"], 9.5, 250):
            p.text(X0, y, part)
            y += 11.5
        y += 2.1
    p.line(85, 435, 170.8)
    y = 173.3
    for ln in em.get("banca") or []:
        p.text(X0, y, ln, color=INK)
        y += 13.9
    p.text(X0, 248.3, "Facture", size=31.5, font="B", color=DARK)
    p.text(X0, 285.7, "Soumis le " + fmt_date(doc["data"]), size=11.5, font="B", color=PINK)

    XB, XC = 255.7, 361.7
    p.text(X0, 317.9, "Facture pour", size=11.5, font="B", color=INK)
    p.text(X0, 335.9, cl["nume"], font="B", color=DARK)
    y = 352.3
    for ln in addr_lines:
        p.text(X0, y, ln)
        y += 11.5
    if cl.get("siret"):
        p.text(X0, y + 4.2, cl["siret"])
    if payable:
        p.text(XB, 317.9, "Payable à", size=11.5, font="B", color=INK)
        p.text(XC, 317.9, "IBAN", size=11.5, font="B", color=INK)
        p.rect(XB - 2.5, 331.5, 545.1, 349.5, ZEBRA)
        p.text(XB, 335.9, em.get("titular") or "")
        p.text(XC, 335.9, em.get("iban") or "", color=(0, 0, 0))
        p.text(XB, 357.5, "Facture #", size=11.5, font="B", color=INK)
        p.text(XC, 357.5, "Date d'échéance", size=11.5, font="B", color=INK)
        p.text(XB, 380.3, doc["numar"])
        p.text(XC, 379.5, fmt_date(doc["scadenta"]), size=11.5, font="B", color=PINK)
    else:
        p.text(XB, 317.9, "Facture #", size=11.5, font="B", color=INK)
        p.text(XC, 317.9, "Date d'échéance", size=11.5, font="B", color=INK)
        p.text(XB, 335.9, doc["numar"])
        p.text(XC, 335.1, fmt_date(doc["scadenta"]), size=11.5, font="B", color=PINK)

    p.line(85, 510, y_rule)
    yh = y_rule + 24.6
    p.text(X0, yh, "Description", size=11.5, font="B", color=DARK)
    p.text(X_QTY_R, yh, "Qty/heures", size=10.5, font="B", color=DARK, align="r")
    p.text(X_PU_R, yh, "Prix unitaire", size=11.5, font="B", color=DARK, align="r")
    p.text(X_TOT_R, yh, "Prix total", size=11.5, font="B", color=DARK, align="r")
    y = y_rows
    prefix = doc.get("prefix_descriere", "Ménage ")
    for i, r in enumerate(rows):
        p.rect(ROW_X0, y, ROW_X1, y + ROW_H + 0.7, ZEBRA if i % 2 == 0 else (1, 1, 1))
        desc = r.get("desc") or ((prefix + r["nom"]) if r.get("casa") else r["nom"])
        while stringWidth(desc, f["R"], 9.5) > 205 and len(desc) > 8:
            desc = desc[:-2].rstrip() + "…"
        p.text(X0, y + 5.7, desc, color=(0, 0, 0))
        p.text(325.5, y + 5.7, fmt_qty(r["qty"]), align="c")
        p.text(X_PU_R, y + 5.7, fmt_money(r["pu"]), align="r")
        p.text(X_TOT_R, y + 5.7, fmt_money(r["total"]), align="r")
        y += ROW_H
    _totals_block(p, y + 0.7, subtotal, ajust, note_lines)
    c.showPage()

    # ---------------- pagina 2: Détail des prestations
    NAME_W, LINE = 84.0, 11.4
    prepared = []
    for r in rows:
        nlines = _wrap(r["nom"], f["R"], 9.5, NAME_W)
        dlines = [fmt_date(d) for d in r["dates"]] or [""]
        hgt = max(ROW_H, 7.2 + LINE * max(len(nlines), len(dlines)))
        prepared.append((r, nlines, dlines, hgt))
    Y_TAB = 270.0
    need = Y_TAB + sum(x[3] for x in prepared) + 150
    h2 = max(PAGE_H, need)
    c.setPageSize((PAGE_W, h2))
    p = _Page(c, h2, f)
    p.text(PAGE_W / 2, 168.8, "Détail des prestations", size=31.5, font="B", color=DARK, align="c")
    p.text(X0, 249.9, "Nom", size=11.5, font="B", color=DARK)
    p.text(192.0, 249.9, "Type", size=11.5, font="B", color=DARK)
    p.text(239.2, 243.5, "Date de", size=11.5, font="B", color=DARK)
    p.text(239.2, 257.1, "nettoyage", size=11.5, font="B", color=DARK)
    p.text(356.7, 250.6, "Qty/heures", size=10.5, font="B", color=DARK, align="r")
    p.text(X_PU_R, 249.9, "Prix unitaire", size=11.5, font="B", color=DARK, align="r")
    p.text(X_TOT_R, 249.9, "Prix total", size=11.5, font="B", color=DARK, align="r")
    y = Y_TAB
    for i, (r, nlines, dlines, hgt) in enumerate(prepared):
        p.rect(ROW_X0, y, ROW_X1, y + hgt, ZEBRA if i % 2 == 0 else (1, 1, 1))
        n = max(len(nlines), len(dlines))
        mid = y + hgt / 2

        def col_top(k, _n=n, _mid=mid):
            return _mid - (LINE * k) / 2 + 0.9

        t = col_top(len(nlines))
        for j, ln in enumerate(nlines):
            p.text(X0, t + j * LINE, ln, color=(0, 0, 0))
        if r["type"]:
            p.text(174.8, col_top(1), r["type"], color=(0, 0, 0))
        t = col_top(len(dlines))
        for j, ln in enumerate(dlines):
            p.text(239.2, t + j * LINE, ln, color=(0, 0, 0))
        t = col_top(1)
        p.text(356.7, t, fmt_qty(r["qty"]), align="r")
        p.text(X_PU_R, t, fmt_pu(r["pu"]), align="r")
        p.text(X_TOT_R, t, fmt_money(r["total"]), align="r")
        y += hgt
    _totals_block(p, y, subtotal, ajust)
    c.showPage()
    c.save()
    return buf.getvalue()


def due_date(d, days=10):
    return d + timedelta(days=days)


def file_name(numar):
    return "Facture %s.pdf" % re.sub(r"[^\w\-]+", "-", numar or "noua").strip("-")


def today():
    return date.today()
