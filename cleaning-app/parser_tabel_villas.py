"""
Al doilea format de factura acceptat (facturi facute in Word, ex. "Best Cleaning Service 2A"):
  - un rezumat "Ref | Désignation des services nettoyage | U | Q | Prix Unitaire | Prix Total" ("01.1 Ménage villas ...")
  - de obicei urmat de un tabel pe case: "Nom (des villas / villa / des apparts) | Date | Quantite | Prix | Total"
Produce aceleasi randuri brute ca parserul SimNettoyage ({nom, type, dates_raw, qty, pu, total}),
deci restul aplicatiei (interventii, reguli, indicatori) ramane acelasi.

Ce intelege:
  - un rand pe casa: datele menajelor ("02.08, 18.08,"), cantitatea, pretul pe menaj, totalul;
  - ore pe aceeasi casa: "2h le 06.08" in coloana datelor si "3+4h" la cantitate -> menajele raman la pret fix,
    orele devin interventii Intermédiaire (tariful orar = (total - menaje x pret) / ore);
  - randuri doar cu ore ("Bellec | 3h le 06.08, ... | 12h | 26€"), "06.08(3.5hX4p)" = 3,5 h x 4 persoane;
  - randul "Intermediaire" / "Interm X" cu lista "4h Alivu le 03.08, ..." / "1h Azur+6h Aria Mezana";
  - suplimente / reduceri pe o data: "04.07+30E", "22.08(-20e)"; ajustari "-100,00€ Cirindinu le 30.07";
  - randuri de total ("Total 1", "TOTAL 1+2") si randuri goale (",00€") sunt ignorate;
  - fara tabel pe case: randurile rezumatului ("01.1 Villa Muredda (14h le 10.08+6h le 13.08) Ens 20h 26€/h 520€");
  - mai multe facturi in acelasi PDF (ex. N05/06 + N06/06): totalul = suma lor.
"""
import difflib
import re
from datetime import date

import pdfplumber

DM = r"(\d{1,2})[./](\d{1,2})(?![\d/]*\d{4})(?!\d)"        # 15.08 / 09/08 (dar nu 31/08/2026)
HOURS_ITEM = re.compile(r"(\d+(?:[.,]\d+)?)\s*h\b\s*(?:\(\s*\d+(?:[.,]\d+)?\s*h\s*\))?\s*(.*?)\s*(?:\b(?:le|la)\b\s*)?"
                        + DM + r"([A-Z]\w*)?")
HOURS_NODATE = re.compile(r"(\d+(?:[.,]\d+)?)\s*h\b\s+([A-Za-zÀ-ÿ&'’][^+,()]*?)\s*(?=\+|,|$)")
PERSONS = re.compile(DM + r"\s*\(\s*(\d+(?:[.,]\d+)?)\s*h\s*[xX×*]\s*(\d+)\s*p?\s*\)")
ADJUST = re.compile(r"(-\s?\d[\d.\s]*(?:,\d{2})?)\s*€\s*(.*?)\s*(?:\ble\b\s*)?" + DM)
EXTRA = re.compile(r"(?:\(\s*)?([+-])\s*(\d+(?:[.,]\d+)?)\s*[eE€](?![a-zA-Z])\s*\)?")
GENERIC = r"(?i)\b(ménage|menage|nettoyage|services?|villas?|appartements?|apparts?|appt|et|d['’]|de|des|du|la|le|les|" \
          r"relocation|ouverture|fermeture|d['’]ouverture|complet|compl[eè]te|grand|petit)\b"


def is_format(text):
    t = (text or "").casefold()
    return "désignation des services" in t or "designation des services" in t or (
        "nom des villas" in t and ("date de menage" in t or "date de ménage" in t))


def money(s):
    s = (s or "").replace("€", "").replace(" ", "").replace(" ", "").strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in ("-", ",00", ".00"):
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


def _num(s):
    return float(s.replace(",", "."))


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
    try:
        d, m = int(dd), int(mm)
        if not (1 <= d <= 31 and 1 <= m <= 12):
            return None
        return f"{d:02d}/{m:02d}/{_year_for(m, inv)}"
    except ValueError:
        return None


# --------------------------------------------------------------------------- antet

def parse_header(words_p0, text):
    head = {"client": None, "numar": None, "data_emitere": None, "scadenta": None}
    m = re.search(r"\bLe\s+(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        try:
            head["data_emitere"] = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    nums = []
    for n in re.findall(r"FACTURE\s+(N[°o]?\s*[\w/.-]+)", text, re.I):
        n = re.sub(r"\s+", "", n)
        if n not in nums:
            nums.append(n)
    head["numar"] = "+".join(nums) or None
    head["facturi_in_pdf"] = len(nums)
    # clientul: blocul din dreapta, intre data ("Le 31/08/2026") si "FACTURE N..."
    le = next((w for w in words_p0 if w["text"] == "Le"), None)
    fa = next((w for w in words_p0 if w["text"].upper() == "FACTURE"), None)
    if le and fa:
        x_min = min(le["x0"], fa["x0"]) - 5
        top0 = min(le["top"], fa["top"]) - 60
        blk = [w for w in words_p0 if w["x0"] >= x_min and top0 < w["top"] < fa["top"] - 2
               and not (abs(w["top"] - le["top"]) < 3)]
        lines = [_txt(l) for l in _lines(blk)]
        lines = [l for l in lines if l and not re.fullmatch(r"(Mme|M|Mr|Mlle)\.?\s.*", l)]   # persoana de contact
        full = " ".join(lines)
        if "," in full:
            client = full.split(",")[0]
            # doar cod postal + oras pe prima linie (fara nume): numele lipseste
        else:
            client = lines[0] if lines else ""
        client = re.sub(r"\s+", " ", client).strip(" ,")
        cut = re.split(r"\s\d{2,}\b", client)[0].strip(" ,–-")       # "SAS KALESI 41 Hameau..." -> "SAS KALESI"
        if len(cut) >= 3:
            client = cut
        if re.match(r"^\d{5}\b", client) or not client:
            client = None
        head["client"] = client
        head["adresa_client"] = full
    return head


# --------------------------------------------------------------------------- tabelul pe case

def _is_header(line):
    ws = [w["text"].casefold().strip(":") for w in line]
    return (any(w.startswith("nom") for w in ws) and "total" in ws and "prix" in ws
            and any(w.startswith(("quantit", "q")) for w in ws))


def _columns(line):
    """Pozitiile coloanelor din antetul tabelului (Nom | Date | Quantite | Prix | Total, cu orice denumire)."""
    ws = [(w, w["text"].casefold()) for w in line]
    nom = next((w for w, t in ws if t.startswith("nom")), None)
    qty = next((w for w, t in ws if t.startswith(("quantit", "q")) and w is not nom), None)
    prix = next((w for w, t in ws if t == "prix"), None)
    total = next((w for w, t in ws if t == "total"), None)
    if not (nom and qty and prix and total):
        return None
    date_w = next((w for w, t in ws if (t.startswith("date") or t == "d") and nom["x1"] < w["x0"] < qty["x0"]), None)
    if date_w is None:          # "Nom des villas Date de menage" lipit / fara titlu: intre nume si cantitate
        cand = [w for w, t in ws if nom["x1"] < w["x0"] < qty["x0"] and t not in ("des", "de", "villas", "villa", "apparts")]
        date_w = cand[0] if cand else None
    date_x = (date_w["x0"] - 8) if date_w else (nom["x1"] + qty["x0"]) / 2
    return {"nom": nom["x0"] - 25, "date": date_x, "qty": qty["x0"] - 10,
            "prix": (qty["x1"] + prix["x0"]) / 2, "total": (prix["x1"] + total["x0"]) / 2,
            "end": total["x1"] + 60}


def table_rows(pages_words):
    """Randurile brute ale tabelului: [{nom, dates_text, qty_text, prix_text, total_text}]."""
    cols, rows, totals = None, [], []
    for words in pages_words:
        for line in _lines(words):
            if _is_header(line):
                cols = _columns(line) or cols
                continue
            if not cols:
                continue
            t = _txt(line)
            if re.search(r"(?i)désignation des services|divers reprises|^ref\b|^best cleaning", t):
                cols = None             # a inceput alta factura in acelasi PDF
                continue
            cells = {"nom": [], "date": [], "qty": [], "prix": [], "total": []}
            for w in line:
                x = w["x0"]
                if x < cols["nom"] or x > cols["end"]:
                    continue            # text decorativ pe margini (ex. numele clientului scris vertical)
                k = ("total" if x >= cols["total"] else "prix" if x >= cols["prix"] else
                     "qty" if x >= cols["qty"] else "date" if x >= cols["date"] else "nom")
                if k == "nom" and cells["nom"] and re.fullmatch(r"\d{1,2}[./]\d{1,2}[,.;]?\w*", w["text"]):
                    k = "date"          # data scrisa putin mai la stanga decat titlul coloanei
                if k == "nom" and not cells["nom"] and rows and re.fullmatch(r"\d{1,2}[./]\d{1,2}[,.;]?", w["text"]):
                    k = "date"          # rand de continuare care incepe cu o data
                cells[k].append(w)
            c = {k: _txt(v) for k, v in cells.items()}
            if not any(c.values()):
                continue
            words_nom = c["nom"].split()
            if words_nom and words_nom[-1].upper() == "TOTAL" or re.match(r"(?i)^\s*total\b", c["nom"]) \
                    or (not c["nom"] and not c["date"] and not c["qty"] and c["total"]):
                v = money(c["total"] or c["prix"])
                if v is not None:
                    totals.append((c["nom"], v))
                continue
            starts = bool(c["nom"]) and bool(c["qty"] or c["prix"] or c["total"] or c["date"]) \
                and not (rows and not (c["qty"] or c["prix"] or c["total"]) and not re.search(DM, c["date"]))
            if starts or not rows:
                rows.append({"nom": c["nom"], "dates_text": c["date"], "qty_text": c["qty"],
                             "prix_text": c["prix"], "total_text": c["total"]})
            else:                       # continuarea randului de mai sus (date / nume pe mai multe randuri)
                r = rows[-1]
                if c["date"]:
                    r["dates_text"] = (r["dates_text"] + " " + c["date"]).strip()
                if c["nom"]:
                    if re.fullmatch(r"[A-Z’'\-]{3,}", c["nom"]) and not c["date"]:
                        continue        # cuvant ratacit cu majuscule (watermark, ex. "HOME")
                    r["nom"] = (r["nom"] + " " + c["nom"]).strip()
                for k in ("qty_text", "prix_text", "total_text"):
                    if not r[k] and c[k.split("_")[0]]:
                        r[k] = c[k.split("_")[0]]
    return rows, totals


# --------------------------------------------------------------------------- randurile rezumatului ("01.1 ...")

def summary_rows(pages_words):
    """Randurile '01.1 Désignation  Ens  Q  PU  Total' (pentru facturile fara tabel pe case)."""
    out = []
    bad = re.compile(r"(?i)total|divers reprises|^ref\b|désignation|unitaire|^services nettoyage$|^\d{2}\.\d+$")
    for words in pages_words:
        for e in [w for w in words if w["text"].casefold() == "ens"]:
            band = [w for w in words if abs(w["top"] - e["top"]) <= 8 and w is not e]
            nums = sorted((w for w in band if w["x0"] > e["x1"]), key=lambda w: w["x0"])
            ref_x = min((w["x0"] for w in words if re.fullmatch(r"0\d\.\d{1,2}", w["text"])), default=-1)
            desc = sorted((w for w in band if w["x1"] <= e["x0"]
                           and not (re.fullmatch(r"0\d\.\d{1,2}", w["text"]) and abs(w["x0"] - ref_x) < 3)),
                          key=lambda w: w["x0"])
            # descrierea poate continua pe randul de dedesubt (inainte de urmatorul "Ens")
            below = [w for w in words if e["top"] + 6 < w["top"] <= e["top"] + 16 and w["x1"] <= e["x0"]]
            below_txt = _txt(sorted(below, key=lambda w: w["x0"]))
            desig = _txt(desc)
            if below_txt and not bad.search(below_txt) and not any(
                    x["text"].casefold() == "ens" and abs(x["top"] - below[0]["top"]) <= 8 for x in words):
                desig = (desig + " " + below_txt).strip()
            if not desig:
                above = [w for w in words if e["top"] - 18 <= w["top"] < e["top"] - 6 and w["x1"] <= e["x0"]]
                at = _txt(sorted(above, key=lambda w: w["x0"]))
                if at and not bad.search(at):
                    desig = at
            joined = re.sub(r"\s+,", ",", " ".join(w["text"] for w in nums))
            parts = joined.split()
            if len(parts) < 2:
                continue
            q, pu, tot = (parts[0], parts[1], parts[-1]) if len(parts) >= 3 else (parts[0], "", parts[-1])
            out.append({"desig": desig, "qty_text": q, "prix_text": pu, "total_text": tot})
    return out


def house_from_filename(name):
    """'Facture  Villa Arasu 1608.pdf' -> 'Villa Arasu' · 'Facture Mme Rocca Serra 1406.pdf' -> 'Rocca Serra'."""
    s = re.sub(r"(?i)\.pdf$", "", name or "")
    s = re.sub(r"(?i)\b(factures?|facture|mme|mr|m|avril|mai|juin|juillet|aout|août|septembre|octobre|novembre|"
               r"decembre|décembre|janvier|fevrier|février|mars)\b\.?", " ", s)
    s = re.sub(r"\d+", " ", s)
    return re.sub(r"\s+", " ", s).strip(" -_.") or None


def _house_from_desig(desig, client, hint=None):
    s = re.sub(r"\(.*?\)", " ", desig)
    s = re.sub(r"\b(?:le|la)\s+\d{1,2}[./]\d{1,2}\b", " ", s)
    s = re.sub(r"\d+(?:[.,]\d+)?\s*h\b", " ", s)
    s = re.sub(GENERIC, " ", s)
    s = re.sub(r"[;()]", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,.-")
    if len(s) < 3 or s.casefold() in ("services nettoyage",):
        return hint or client or desig
    return s


# --------------------------------------------------------------------------- interpretare

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
    keys = {key(h): h for h in houses}
    close = difflib.get_close_matches(n, keys, n=2, cutoff=0.8)     # o litera gresita ("Zouk" / "Zouck")
    if len(close) == 1:
        return keys[close[0]]
    return name.strip()


def _qty(qtxt):
    """'3' -> (3, None) · '3+4h' -> (3, 4) · '12h' / '184H' -> (None, 12) · '12,5h' -> (None, 12.5)."""
    n_q = h_q = None
    for part in re.sub(r"\s+", "", qtxt or "").casefold().split("+"):
        v = re.fullmatch(r"(\d+(?:[.,]\d+)?)(h?)", part)
        if v:
            num = _num(v.group(1))
            if v.group(2):
                h_q = (h_q or 0) + num
            else:
                n_q = (n_q or 0) + num
    return n_q, h_q


def _row_type(nom):
    n = nom.casefold()
    if "ouvert" in n:
        return "Ouverture"
    if "fermet" in n:
        return "Fermeture"
    if "linge" in n:
        return "Location linge"
    return "Standard"


def build_rows(raw, inv):
    """Randuri brute -> randurile aplicatiei (pret fix + ore Intermédiaire) + ajustari + avertismente."""
    out, adjust, warns = [], [], []
    def clean_name(n):
        n = re.sub(r"\s+", " ", n).strip(" ,;")
        return re.sub(r"(?i)^(ouvertures?|fermetures?)\s+(?=\S)", "", n).strip(" ,;") or n

    houses = [clean_name(r["nom"]) for r in raw if r["nom"] and not r["nom"].casefold().startswith("interm")]
    for r in raw:
        typ = _row_type(r.get("_desig") or r["nom"])
        nom, txt = clean_name(r["nom"]), r["dates_text"]
        if typ == "Location linge" and r.get("_rezumat") and out:
            nom = out[-1]["nom"]        # lenjeria e pentru casa facturii, nu o casa separata
        pu, total = money(r["prix_text"]), money(r["total_text"])
        n_q, h_q = _qty(r["qty_text"])
        if not nom:
            continue
        if not re.search(r"\d", txt) and n_q is None and h_q is None and not total:
            continue                    # rand gol (casa fara menaje luna asta: ",00€" / "-")
        for m in ADJUST.finditer(txt):   # "-100,00€ Cirindinu le 30.07": nu e un menaj
            adjust.append({"suma": money(m.group(1)), "casa": _match_house(m.group(2), houses) or nom,
                           "data": _d(m.group(3), m.group(4), inv)})
        txt = ADJUST.sub(" ", txt)
        extra = sum((1 if s == "+" else -1) * _num(v) for s, v in EXTRA.findall(txt))
        txt = EXTRA.sub(" ", txt)
        interm_list = nom.casefold().startswith("interm")
        prefix = re.sub(r"(?i)^interm\w*\.?\s*", "", nom) if interm_list else ""
        base_house = _match_house(prefix, houses) if prefix else nom

        hours = []           # (ore, casa, data)
        last_named = None
        for m in PERSONS.finditer(txt):          # "06.08(3.5hX4p)" = 3,5 h x 4 persoane
            hours.append((round(_num(m.group(3)) * int(float(m.group(4))), 2), base_house, _d(m.group(1), m.group(2), inv)))
        txt = PERSONS.sub(" ", txt)
        for m in HOURS_ITEM.finditer(txt):
            h = _num(m.group(1))
            who = (m.group(2) or "").strip()
            who = re.sub(r"(?i)\b(lits?|le|la)\b", " ", who).strip(" ,")
            who = who or (m.group(5) or "")           # "1h le 15.07E09" -> casa dupa data
            if interm_list:
                casa = _match_house(who, houses, prefix + " " if prefix else "") if who else (last_named or base_house)
                if who:
                    last_named = casa
            else:
                casa = nom
            hours.append((h, casa or nom, _d(m.group(3), m.group(4), inv)))
        rest = HOURS_ITEM.sub(" ", txt)
        if interm_list:                          # "1h Azur+6h Aria Mezana" (fara date)
            for m in HOURS_NODATE.finditer(rest):
                hours.append((_num(m.group(1)), _match_house(m.group(2).strip(), houses) or nom, None))
            rest = HOURS_NODATE.sub(" ", rest)
        loose = [_num(x) for x in re.findall(r"\+\s*(\d+(?:[.,]\d+)?)\s*h\b", rest, re.I)]   # "31.08 +6h"
        rest = re.sub(r"\+\s*\d+(?:[.,]\d+)?\s*h\b", " ", rest, flags=re.I)
        dates = [d for d in (_d(a, b, inv) for a, b in re.findall(DM, rest)) if d]
        if loose:
            hours += [(h, nom, dates[0] if len(dates) == 1 else None) for h in loose]
        h_sum = round(sum(h for h, _, _ in hours), 2)
        only_hours = interm_list or (not dates and hours) or (n_q is None and h_q is not None)
        if only_hours and not interm_list and not hours and h_q and pu and total \
                and abs(pu - total) < 0.01 and abs(pu * h_q - total) > 0.01:
            # "Ouverture Alivu | 30.06 | 14h | 364€ | 364€": pret fix, orele sunt doar informative
            only_hours, n_q, h_q = False, float(len(dates) or 1), None

        if only_hours:
            derived = False
            if not hours and h_q:                # "Clemenzia | 15.10 | 8h | 26€": orele impartite pe date
                hours = [(round(h_q / len(dates), 2), base_house if interm_list else nom, d) for d in dates] or [(h_q, nom, None)]
                h_sum, derived = h_q, True
            rate = pu if pu else (round(total / h_sum, 2) if total and h_sum else None)
            if derived and total and h_sum and rate and abs(rate * h_sum - total) > 0.01:
                warns.append(f"{nom}: {h_sum:g} h x {rate:g} € = {h_sum * rate:.2f} € ≠ total {total:.2f} € — păstrez totalul facturii")
                rate = round(total / h_sum, 4)
            if h_q is not None and abs(h_q - h_sum) > 0.01:
                warns.append(f"{nom}: {h_q:g} ore la cantitate, dar {h_sum:g} ore în listă — verificați")
            # in tabelul pe case, orele unei case = Intermédiaire; in rezumat ("Ménage villa le 12.08 | 12h") = menaj la ora
            kind = typ if r.get("_rezumat") else ("Intermédiaire" if (interm_list or typ == "Standard") else typ)
            for h, casa, d in hours:
                out.append({"nom": casa, "type": kind, "dates_raw": [d] if d else [],
                            "qty": h, "pu": rate, "total": round(h * rate, 2) if rate else None})
            if not hours and total:
                out.append({"nom": nom, "type": kind, "dates_raw": [], "qty": h_q, "pu": pu, "total": total})
            continue

        n = n_q if n_q is not None else len(dates)
        if not n and total:
            n = 1.0                     # "Etat de lieux Pinarello | | | 50€ | 50€": o singura prestatie, fara data
        if pu is None and total is not None and n:
            pu = round((total - extra) / n, 2)
        menaje_total = round((n or 0) * pu + extra, 2) if pu is not None else None
        if hours or h_q:
            h_tot = h_sum or h_q or 0
            if h_q is not None and h_sum and abs(h_q - h_sum) > 0.01:
                warns.append(f"{nom}: {h_q:g} ore la cantitate, dar {h_sum:g} ore lângă date — verificați")
            rate = round((total - menaje_total) / h_tot, 2) if (total is not None and menaje_total is not None and h_tot) else None
            out.append({"nom": nom, "type": typ, "dates_raw": dates, "qty": n, "pu": pu, "total": menaje_total})
            for h, casa, d in (hours or [(h_tot, nom, None)]):
                out.append({"nom": casa, "type": "Intermédiaire", "dates_raw": [d] if d else [],
                            "qty": h, "pu": rate, "total": round(h * rate, 2) if rate else None})
        else:
            out.append({"nom": nom, "type": typ, "dates_raw": dates, "qty": n, "pu": pu,
                        "total": total if total is not None else menaje_total})
    return out, adjust, warns


def summary_to_raw(summ, client, hint=None):
    """Randurile rezumatului -> acelasi format ca tabelul pe case."""
    raw = []
    for s in summ:
        d = s["desig"]
        nom = _house_from_desig(d, client, hint)
        inside = " ".join(re.findall(r"\((.*?)\)", d))
        outside = re.sub(r"\(.*?\)", " ", d)
        dates_text = outside + " " + inside
        n_q, h_q = _qty(s["qty_text"])
        if h_q is not None:
            # "(14h le 10.08+6h le 13.08(12h le 13.08))": orele cu data; daca suma nu iese, ignoram parantezele interioare
            items = [(m.group(0)) for m in HOURS_ITEM.finditer(d)]
            if items:
                top = re.sub(r"\([^()]*\)", " ", re.sub(r"^[^(]*\((.*)\)[^)]*$", r"\1", d))
                tot_top = sum(_num(m.group(1)) for m in HOURS_ITEM.finditer(top))
                dates_text = top if abs(tot_top - h_q) < 0.01 else d
            else:
                dates_text = outside
        raw.append({"nom": nom, "dates_text": dates_text, "qty_text": s["qty_text"],
                    "prix_text": s["prix_text"], "total_text": s["total_text"],
                    "_desig": d, "_rezumat": True})
    return raw


def loose_rows(pages_words):
    """Tabel fara antet ("Grenadier  13.05  240,00€"): rand = nume + date + suma."""
    out = []
    started = False
    for words in pages_words:
        for line in _lines(words):
            t = _txt(line)
            if re.search(r"(?i)total\s+divers\s+reprises", t):
                started = True
                continue
            if not started:
                continue
            m = re.fullmatch(r"(.+?)\s+((?:\d{1,2}[./]\d{1,2},?\s*)+)\s+([\d.\s]+,\d{2})\s*€", t)
            if m and not re.search(r"(?i)facture|réglée|virement", t):
                out.append({"nom": m.group(1), "dates_text": m.group(2), "qty_text": "", "prix_text": "",
                            "total_text": m.group(3)})
    return out


def parse(pdf_source, max_pages=6, file_name=None):
    if hasattr(pdf_source, "seek"):
        pdf_source.seek(0)
    with pdfplumber.open(pdf_source) as pdf:
        pages = pdf.pages[:max_pages]
        pages_words = [p.extract_words() for p in pages]
        text = "\n".join((p.extract_text() or "") for p in pages)
    head = parse_header(pages_words[0] if pages_words else [], text)
    inv = head.get("data_emitere")
    # suma de plata: "Total DIVERS REPRISES" (una pe factura; mai multe facturi in acelasi PDF -> suma lor)
    totals_doc = [money(x) for x in re.findall(r"Total\s+DIVERS\s+REPRISES\s*\n?\s*([\d.\s]+,\s?\d{2})\s*€", text)]
    totals_doc = [t for t in totals_doc if t is not None]
    total_doc = round(sum(totals_doc), 2) if totals_doc else None
    raw, table_totals = table_rows(pages_words)
    source = "tabel"
    if not raw:
        raw = loose_rows(pages_words)
        source = "tabel fara antet"
        if not raw or (total_doc is not None and abs(sum(money(r["total_text"]) or 0 for r in raw) - total_doc) > 0.01):
            raw = summary_to_raw(summary_rows(pages_words), head.get("client"), house_from_filename(file_name))
            source = "rezumat"
    rows, adjust, warns = build_rows(raw, inv)
    if source == "rezumat" and rows:
        warns.append("Factura nu are tabel pe case: casa e luată din descriere — verificați numele casei")
    return head, rows, adjust, warns, table_totals, total_doc
