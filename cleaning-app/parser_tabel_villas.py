"""
Al doilea format de factura acceptat: tabelul "Nom des villas | Date de menage | Quantite | Prix | Total"
(ex. facturile "Best Cleaning Service 2A"). Produce aceleasi randuri brute ca parserul SimNettoyage
({nom, type, dates_raw, qty, pu, total}), deci restul aplicatiei (interventii, reguli, indicatori) e acelasi.

Ce intelege:
  - un rand pe casa: datele menajelor ("02.08, 18.08,"), cantitatea, pretul pe menaj, totalul;
  - ore in plus pe aceeasi casa: "2h le 06.08" in coloana datelor si "3+4h" la cantitate
    -> menajele raman forfait, orele devin interventii Intermédiaire (tariful orar = (total - menaje x pret) / ore);
  - randuri doar cu ore ("Bellec | 3h le 06.08, 3h le 13.08 | 12h | 26€");
  - randul "Intermediaire" / "Interm X" cu lista "4h Alivu le 03.08, ..." -> ore pe fiecare casa din lista;
  - ajustari "-100,00€ Cirindinu le 30.07" -> nu sunt menaje: apar ca ajustari (avertisment), nu ca interventii.
"""
import re
from datetime import date

import pdfplumber

MONEY = re.compile(r"^-?\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})?€?$|^-?\d+(?:,\d{2})?€$")
DM = r"(\d{1,2})\.(\d{1,2})"
HOURS_ITEM = re.compile(r"(\d+(?:[.,]\d+)?)\s*h\b\s*(?:\(\s*\d+(?:[.,]\d+)?\s*h\s*\))?\s*(.*?)\s*(?:\ble\b\s*)?" + DM + r"\b")
ADJUST = re.compile(r"(-\s?\d[\d.\s]*(?:,\d{2})?)\s*€\s*(.*?)\s*(?:\ble\b\s*)?" + DM)


def is_format(text):
    t = (text or "").casefold()
    return "nom des villas" in t and ("date de menage" in t or "date de ménage" in t)


def money(s):
    s = (s or "").replace("€", "").replace(" ", "").replace(" ", "").strip()
    if not s:
        return None
    neg = s.startswith("-")
    s = s.lstrip("-")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _lines(words, tol=3.0):
    rows = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if rows and abs(rows[-1][0] - w["top"]) <= tol:
            rows[-1][1].append(w)
        else:
            rows.append([w["top"], [w]])
    return [sorted(ws, key=lambda w: w["x0"]) for _, ws in rows]


def _txt(ws):
    return " ".join(w["text"] for w in ws).strip()


def _year_for(month, inv):
    """Anul unei date fara an (ex. '15.08'), dupa data facturii: luni de dupa factura = anul trecut."""
    if not inv:
        return date.today().year
    return inv.year - 1 if month > inv.month + 1 else inv.year


def _d(dd, mm, inv):
    return f"{int(dd):02d}/{int(mm):02d}/{_year_for(int(mm), inv)}"


def parse_header(words_p0, text):
    head = {"client": None, "numar": None, "data_emitere": None, "scadenta": None}
    m = re.search(r"\bLe\s+(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        try:
            head["data_emitere"] = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    m = re.search(r"FACTURE\s+(N[°o]?\s*[\w/.-]+)", text, re.I)
    if m:
        head["numar"] = re.sub(r"\s+", "", m.group(1))
    # clientul: blocul din dreapta, intre data ("Le 31/08/2026") si "FACTURE N..."
    le = next((w for w in words_p0 if w["text"] == "Le"), None)
    fa = next((w for w in words_p0 if w["text"].upper() == "FACTURE"), None)
    if le and fa:
        x_min = min(le["x0"], fa["x0"]) - 5
        blk = [w for w in words_p0 if w["x0"] >= x_min and le["top"] + 4 < w["top"] < fa["top"] - 2]
        lines = [_txt(l) for l in _lines(blk)]
        lines = [l for l in lines if not re.fullmatch(r"(Mme|M|Mr|Mlle)\.?\s.*", l)]   # persoana de contact
        full = " ".join(lines)
        head["client"] = (full.split(",")[0] or full).strip() or None
        head["adresa_client"] = full
    return head


def _columns(line):
    """Pozitiile coloanelor din antetul tabelului."""
    get = {w["text"].casefold(): w for w in line}
    nom, date_w = get.get("nom"), get.get("date")
    qty = next((w for w in line if w["text"].casefold().startswith("quantit")), None)
    prix, total = get.get("prix"), get.get("total")
    if not (nom and date_w and qty and prix and total):
        return None
    return {"nom": nom["x0"] - 15, "date": date_w["x0"] - 6, "qty": qty["x0"] - 6,
            "prix": (qty["x1"] + prix["x0"]) / 2, "total": (prix["x1"] + total["x0"]) / 2,
            "end": total["x1"] + 60}


def table_rows(pages_words):
    """Randurile brute ale tabelului: [{nom, dates_text, qty_text, prix_text, total_text}]."""
    cols, rows, total_final = None, [], None
    for words in pages_words:
        for line in _lines(words):
            t = _txt(line).casefold()
            if "nom" in t and "villas" in t and "total" in t:
                cols = _columns(line) or cols
                continue
            if not cols:
                continue
            cells = {"nom": [], "date": [], "qty": [], "prix": [], "total": []}
            for w in line:
                x = w["x0"]
                if x < cols["nom"] or x > cols["end"]:
                    continue            # text decorativ pe margini (ex. numele clientului scris vertical)
                k = ("total" if x >= cols["total"] else "prix" if x >= cols["prix"] else
                     "qty" if x >= cols["qty"] else "date" if x >= cols["date"] else "nom")
                cells[k].append(w)
            c = {k: _txt(v) for k, v in cells.items()}
            if c["nom"].upper() == "TOTAL" or (c["nom"].upper().endswith("TOTAL") and not c["date"]):
                total_final = money(c["total"] or c["prix"])
                continue
            if not any(c.values()):
                continue
            starts = bool(c["nom"]) and bool(c["qty"] or c["prix"] or c["total"])
            if starts or not rows:
                rows.append({"nom": c["nom"], "dates_text": c["date"], "qty_text": c["qty"],
                             "prix_text": c["prix"], "total_text": c["total"]})
            else:                       # continuarea randului de mai sus (date pe mai multe randuri)
                r = rows[-1]
                if c["date"]:
                    r["dates_text"] = (r["dates_text"] + " " + c["date"]).strip()
                if c["nom"] and not c["date"]:
                    continue            # cuvant ratacit (ex. watermark) in coloana numelor
                if c["nom"]:
                    r["dates_text"] = (r["dates_text"] + " " + c["nom"]).strip()
    return rows, total_final


def _match_house(name, houses, prefix=""):
    """'MD' -> 'Villa M&D', 'H05' -> 'Dupuis H05', 'Figuier' -> 'Le Figuier' (dupa casele din aceeasi factura)."""
    def key(s):
        return re.sub(r"[^a-z0-9]", "", s.casefold())
    n = key(name)
    if not n:
        return name
    for h in houses:
        if key(h) == n:
            return h
    if prefix:
        for h in houses:
            if key(h) == key(prefix + name):
                return h
    cands = [h for h in houses if key(h).endswith(n) or key(h).startswith(n)]
    cands = cands or [h for h in houses if n in key(h)]
    if len(set(cands)) == 1:
        return cands[0]
    # o litera gresita ("Zouk" / "Zouck")
    import difflib
    close = difflib.get_close_matches(n, {key(h): h for h in houses}, n=2, cutoff=0.8)
    if len(close) == 1:
        return {key(h): h for h in houses}[close[0]]
    return name.strip()


def build_rows(raw, inv):
    """Randuri brute -> randurile aplicatiei (forfait + ore Intermédiaire) + ajustari + avertismente."""
    out, adjust, warns = [], [], []
    houses = [r["nom"] for r in raw if r["nom"] and not r["nom"].casefold().startswith("interm")]
    for r in raw:
        nom, txt = r["nom"].strip(), r["dates_text"]
        pu, total = money(r["prix_text"]), money(r["total_text"])
        qtxt = r["qty_text"].replace(" ", "")
        if not nom:
            continue
        # ajustari ("-100,00€ Cirindinu le 30.07") scoase din text
        for m in ADJUST.finditer(txt):
            adjust.append({"suma": money(m.group(1)), "casa": _match_house(m.group(2), houses) or nom,
                           "data": _d(m.group(3), m.group(4), inv)})
        txt = ADJUST.sub(" ", txt)
        discount = sum(float(x.replace(",", ".")) for x in re.findall(r"\(\s*-\s*(\d+(?:[.,]\d+)?)\s*e?\s*\)", txt, re.I))
        interm_list = nom.casefold().startswith("interm")
        prefix = re.sub(r"(?i)^interm\w*\.?\s*", "", nom) if interm_list else ""

        hours = []           # (ore, casa, data)
        for m in HOURS_ITEM.finditer(txt):
            h = float(m.group(1).replace(",", "."))
            who = m.group(2).strip()
            who = re.sub(r"(?i)\b(lits?|le)\b", " ", who).strip(" ,")
            casa = _match_house(who, houses, prefix + " " if prefix else "") if (interm_list and who) else nom
            hours.append((h, casa or nom, _d(m.group(3), m.group(4), inv)))
        rest = HOURS_ITEM.sub(" ", txt)
        # ore fara data ("31.08 +6h") -> pe singura data a randului
        loose = [float(x.replace(",", ".")) for x in re.findall(r"\+\s*(\d+(?:[.,]\d+)?)\s*h\b", rest, re.I)]
        rest = re.sub(r"\+\s*\d+(?:[.,]\d+)?\s*h\b", " ", rest, flags=re.I)
        dates = [_d(a, b, inv) for a, b in re.findall(DM + r"\b", rest)]
        if loose:
            hours += [(h, nom, dates[0] if len(dates) == 1 else None) for h in loose]

        # cantitatea: "3", "3+4h", "12h", "184H"
        n_q = h_q = None
        for part in qtxt.casefold().split("+"):
            v = re.fullmatch(r"(\d+(?:[.,]\d+)?)(h?)", part.strip())
            if v:
                num = float(v.group(1).replace(",", "."))
                if v.group(2):
                    h_q = (h_q or 0) + num
                else:
                    n_q = (n_q or 0) + num
        h_sum = round(sum(h for h, _, _ in hours), 2)
        only_hours = interm_list or (not dates and hours) or (n_q is None and h_q is not None)

        if only_hours:
            rate = pu if pu else (round(total / h_sum, 2) if total and h_sum else None)
            if h_q is not None and abs(h_q - h_sum) > 0.01:
                warns.append(f"{nom}: {h_q:g} ore la cantitate, dar {h_sum:g} ore în listă — verificați")
            for h, casa, d in hours:
                out.append({"nom": casa, "type": "Intermédiaire", "dates_raw": [d] if d else [],
                            "qty": h, "pu": rate, "total": round(h * rate, 2) if rate else None})
            if not hours and total:
                out.append({"nom": nom, "type": "Intermédiaire", "dates_raw": [], "qty": h_q, "pu": pu, "total": total})
            continue

        n = n_q if n_q is not None else len(dates)
        menaje_total = round((n or 0) * (pu or 0) - discount, 2) if pu is not None else None
        if hours or h_q:
            h_tot = h_sum or h_q or 0
            if h_q is not None and h_sum and abs(h_q - h_sum) > 0.01:
                warns.append(f"{nom}: {h_q:g} ore la cantitate, dar {h_sum:g} ore lângă date — verificați")
            rate = round((total - menaje_total) / h_tot, 2) if (total is not None and menaje_total is not None and h_tot) else None
            out.append({"nom": nom, "type": "Standard", "dates_raw": dates, "qty": n, "pu": pu,
                        "total": menaje_total})
            for h, casa, d in (hours or [(h_tot, nom, None)]):
                out.append({"nom": casa, "type": "Intermédiaire", "dates_raw": [d] if d else [],
                            "qty": h, "pu": rate, "total": round(h * rate, 2) if rate else None})
        else:
            out.append({"nom": nom, "type": "Standard", "dates_raw": dates, "qty": n, "pu": pu,
                        "total": total if total is not None else menaje_total})
    return out, adjust, warns


def parse(pdf_source, max_pages=6):
    if hasattr(pdf_source, "seek"):
        pdf_source.seek(0)
    with pdfplumber.open(pdf_source) as pdf:
        pages = pdf.pages[:max_pages]
        pages_words = [p.extract_words() for p in pages]
        text = "\n".join((p.extract_text() or "") for p in pages)
    head = parse_header(pages_words[0] if pages_words else [], text)
    raw, total_final = table_rows(pages_words)
    rows, adjust, warns = build_rows(raw, head.get("data_emitere"))
    # suma de plata: totalul de pe prima pagina ("Total DIVERS REPRISES"); TOTAL-ul tabelului e verificarea
    m = re.findall(r"Total\s+DIVERS\s+REPRISES\s+([\d.\s]+,\d{2})\s*€", text)
    total_doc = money(m[-1]) if m else None
    return head, rows, adjust, warns, total_final, total_doc
