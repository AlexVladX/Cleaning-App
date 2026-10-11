"""
Indicatori calculati doar din facturi (etapa 1):
  - filtru de perioada (toata perioada / an / sezon / luna)
  - clienti, evolutie lunara, calendar (zile aglomerate), frecventa si case inactive
  - verificarea preturilor: abateri izolate (probabil greseli) vs schimbari de pret
"""
from collections import defaultdict
from datetime import date, datetime
from statistics import median

LUNI = ["", "Ianuarie", "Februarie", "Martie", "Aprilie", "Mai", "Iunie", "Iulie",
        "August", "Septembrie", "Octombrie", "Noiembrie", "Decembrie"]
LUNI_SCURT = ["", "ian", "feb", "mar", "apr", "mai", "iun", "iul", "aug", "sep", "oct", "nov", "dec"]
ZILE = ["Luni", "Marți", "Miercuri", "Joi", "Vineri", "Sâmbătă", "Duminică"]


# ------------------------------------------------------------------ perioade

def _month_key(d):
    return d[:7]  # 'YYYY-MM'


def month_label(mk, short=False):
    y, m = int(mk[:4]), int(mk[5:7])
    return f"{LUNI_SCURT[m]} {str(y)[2:]}" if short else f"{LUNI[m]} {y}"


def months_range(first_mk, last_mk):
    y, m = int(first_mk[:4]), int(first_mk[5:7])
    ly, lm = int(last_mk[:4]), int(last_mk[5:7])
    out = []
    while (y, m) <= (ly, lm):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def period_options(ivs):
    """Lista de perioade disponibile, construita din datele interventiilor."""
    months = sorted({_month_key(iv["data_iso"]) for iv in ivs if iv.get("data_iso")})
    opts = [{"key": "all", "label": "Toată perioada", "start": None, "end": None}]
    years = sorted({m[:4] for m in months}, reverse=True)
    for y in years:
        opts.append({"key": f"y{y}", "label": f"Anul {y}", "start": f"{y}-01-01", "end": f"{y}-12-31"})
        if any(m[:4] == y and "04" <= m[5:7] <= "10" for m in months):
            opts.append({"key": f"s{y}", "label": f"Sezon {y} (apr–oct)", "start": f"{y}-04-01", "end": f"{y}-10-31"})
        if any(m[:4] == y and m[5:7] in ("07", "08") for m in months):
            opts.append({"key": f"h{y}", "label": f"Sezon fierbinte {y} (iul–aug)", "start": f"{y}-07-01", "end": f"{y}-08-31"})
    for mk in reversed(months):
        y, m = int(mk[:4]), int(mk[5:7])
        last_day = (date(y + (m == 12), m % 12 + 1, 1) - date(y, m, 1)).days
        opts.append({"key": f"m{mk}", "label": month_label(mk), "start": f"{mk}-01", "end": f"{mk}-{last_day:02d}"})
    return opts


def resolve_period(key, ivs):
    opts = period_options(ivs)
    return next((o for o in opts if o["key"] == key), opts[0]), opts


def filter_ivs(ivs, period):
    if not period or not period.get("start"):
        return list(ivs)
    s, e = period["start"], period["end"]
    return [iv for iv in ivs if iv.get("data_iso") and s <= iv["data_iso"] <= e]


def period_months(ivs_all, period):
    """Lunile perioadei, limitate la intervalul in care exista date (pentru 'pe luna')."""
    months = sorted({_month_key(iv["data_iso"]) for iv in ivs_all if iv.get("data_iso")})
    if not months:
        return []
    first, last = months[0], months[-1]
    if period and period.get("start"):
        first = max(first, period["start"][:7])
        last = min(last, period["end"][:7])
    return months_range(first, last) if first <= last else []


# ------------------------------------------------------------------ proprietati: frecventa & inactivitate

def property_activity(ivs_period, ivs_all, prop_key, n_months, today=None):
    """Pe fiecare cheie de proprietate: frecventa/luna in perioada, ultima interventie, inactivitate."""
    today = today or date.today()
    dates_all = defaultdict(set)
    for iv in ivs_all:
        if iv.get("data_iso"):
            dates_all[prop_key(iv.get("proprietate") or "")].add(iv["data_iso"])
    count_period = defaultdict(int)
    for iv in ivs_period:
        count_period[prop_key(iv.get("proprietate") or "")] += 1
    out = {}
    for key, ds in dates_all.items():
        ds = sorted(ds)
        last = datetime.fromisoformat(ds[-1]).date()
        gaps = [(datetime.fromisoformat(b) - datetime.fromisoformat(a)).days for a, b in zip(ds, ds[1:])]
        typical = median(gaps) if gaps else None
        since = (today - last).days
        threshold = max(45, 3 * typical) if typical else 60
        out[key] = {
            "frecventa_luna": round(count_period.get(key, 0) / n_months, 2) if n_months else None,
            "ultima": ds[-1],
            "zile_de_la_ultima": since,
            "interval_tipic": round(typical) if typical else None,
            "inactiva": since > threshold,
        }
    return out


# ------------------------------------------------------------------ clienti

def clients_table(ivs, norm):
    groups = {}
    total = sum(float(iv.get("suma") or 0) for iv in ivs) or 0.0
    for iv in ivs:
        name = (iv.get("client") or "Necunoscut").strip()
        k = norm(name) or "necunoscut"
        g = groups.setdefault(k, {"nume_variante": defaultdict(int), "venit": 0.0, "interventii": 0,
                                  "ore": 0.0, "case": set(), "ultima": ""})
        g["nume_variante"][name] += 1
        if iv.get("client_factura"):
            g.setdefault("membri", set()).add(iv["client_factura"].strip())
        g["venit"] += float(iv.get("suma") or 0)
        g["interventii"] += 1
        if iv.get("tip_facturare") == "heure":
            g["ore"] += float(iv.get("ore") or 0)
        g["case"].add(norm(iv.get("proprietate") or ""))
        if (iv.get("data_iso") or "") > g["ultima"]:
            g["ultima"] = iv["data_iso"]
    rows = []
    for g in groups.values():
        rows.append({
            "client": max(g["nume_variante"], key=g["nume_variante"].get),
            "venit": round(g["venit"], 2),
            "procent": round(g["venit"] / total * 100, 1) if total else 0,
            "interventii": g["interventii"],
            "medie": round(g["venit"] / g["interventii"], 2) if g["interventii"] else 0,
            "ore": round(g["ore"], 2),
            "case": len(g["case"]),
            "ultima": g["ultima"],
            "membri": sorted(g.get("membri") or [], key=str.casefold),
        })
    rows.sort(key=lambda r: -r["venit"])
    top3 = round(sum(r["procent"] for r in rows[:3]), 1)
    return rows, top3


# ------------------------------------------------------------------ evolutie lunara

def monthly_series(ivs_all, months, prop_key, today=None):
    current = (today or date.today()).isoformat()[:7]
    by = defaultdict(lambda: {"venit": 0.0, "interventii": 0, "ore": 0.0, "case": set()})
    for iv in ivs_all:
        if not iv.get("data_iso"):
            continue
        b = by[_month_key(iv["data_iso"])]
        b["venit"] += float(iv.get("suma") or 0)
        b["interventii"] += 1
        if iv.get("tip_facturare") == "heure":
            b["ore"] += float(iv.get("ore") or 0)
        b["case"].add(prop_key(iv.get("proprietate") or ""))

    def pct(a, b):
        return round((a - b) / b * 100, 1) if b else None

    rows = []
    for mk in months:
        b = by.get(mk, {"venit": 0.0, "interventii": 0, "ore": 0.0, "case": set()})
        y, m = int(mk[:4]), int(mk[5:7])
        prev = f"{y - (m == 1):04d}-{(m - 2) % 12 + 1:02d}"
        ly = f"{y - 1:04d}-{m:02d}"
        rows.append({
            "luna": mk, "eticheta": month_label(mk), "scurt": month_label(mk, short=True),
            "venit": round(b["venit"], 2), "interventii": b["interventii"],
            "ore": round(b["ore"], 2), "case_active": len(b["case"]),
            "vs_luna_trecuta": pct(b["venit"], by[prev]["venit"]) if prev in by else None,
            "vs_an_trecut": pct(b["venit"], by[ly]["venit"]) if ly in by else None,
            "in_curs": mk == current,  # luna curenta: incompleta, comparatiile nu sunt inca relevante
        })
    return rows


# ------------------------------------------------------------------ calendar

def calendar_data(ivs, prop_display):
    days = defaultdict(lambda: {"n": 0, "venit": 0.0, "case": []})
    for iv in ivs:
        if not iv.get("data_iso"):
            continue
        d = days[iv["data_iso"]]
        d["n"] += 1
        d["venit"] += float(iv.get("suma") or 0)
        d["case"].append(prop_display(iv.get("proprietate") or ""))
    out = {k: {"n": v["n"], "venit": round(v["venit"], 2), "case": sorted(set(v["case"]))} for k, v in days.items()}
    top = sorted(out.items(), key=lambda kv: (-kv[1]["n"], kv[0]))[:15]
    # media pe zi a saptamanii (doar zilele cu activitate)
    wd = defaultdict(list)
    for k, v in out.items():
        wd[datetime.fromisoformat(k).weekday()].append(v["n"])
    weekday = [{"zi": ZILE[i], "medie": round(sum(wd[i]) / len(wd[i]), 1) if wd[i] else 0,
                "zile_active": len(wd[i]), "total": sum(wd[i])} for i in range(7)]
    return {"zile": out, "top": [{"data": k, **v} for k, v in top], "pe_zi_saptamana": weekday}


# ------------------------------------------------------------------ verificarea preturilor

def price_checks(ivs_all, period, prop_key, prop_display):
    """
    Pe fiecare (casa, tip menaj, mod de facturare) urmarim pretul unitar factura dupa factura.
      - 'abatere': un pret diferit izolat, intre doua facturi cu acelasi pret obisnuit
        -> foarte probabil o greseala de redactare
      - 'schimbare': pretul se schimba si ramane la noua valoare -> schimbare de tarif (informativ)
      - 'de confirmat': ultima factura are alt pret decat cele dinainte (inca nu stim daca ramane)
    Doar facturile din perioada aleasa sunt raportate (istoricul complet e folosit ca referinta).
    """
    groups = defaultdict(dict)
    for iv in ivs_all:
        pu = iv.get("pret_unitar")
        if pu is None:
            continue
        # clientul face parte din cheie: aceeasi casa la alt client poate avea alt tarif, legitim
        g = (prop_key(iv.get("proprietate") or ""), iv.get("tip_menaj") or "standard",
             iv.get("tip_facturare"), (iv.get("client") or "").strip().casefold())
        inv = (iv.get("factura") or "", iv.get("_file_id") or iv.get("_source") or "")
        e = groups[g].setdefault(inv, {"pret": round(float(pu), 2), "data": iv.get("data_iso") or "",
                                        "factura": iv.get("factura"), "fisier": iv.get("_source"),
                                        "client": iv.get("client"), "nume": iv.get("proprietate")})
        if iv.get("data_iso") and (not e["data"] or iv["data_iso"] < e["data"]):
            e["data"] = iv["data_iso"]
    s = period.get("start") if period else None
    e_ = period.get("end") if period else None
    in_period = (lambda d: True) if not s else (lambda d: bool(d) and s <= d <= e_)

    out = []
    for (key, tip, mode, client_k), invs in groups.items():
        seq = sorted(invs.values(), key=lambda x: (x["data"] or "9999", x["factura"] or ""))
        if len(seq) < 2:
            continue
        prices = [x["pret"] for x in seq]
        base = prices[0]  # pretul "obisnuit" curent; o abatere izolata nu il schimba
        for i in range(1, len(seq)):
            p = prices[i]
            if abs(p - base) < 0.01:
                continue
            prev = base
            nxt = prices[i + 1] if i + 1 < len(seq) else None
            if nxt is not None and abs(nxt - base) < 0.01:
                kind = "abatere"            # 330, 300, 330 -> 300 e o greseala probabila
            elif nxt is None:
                kind = "de confirmat"       # ultima factura: nu stim inca daca noul pret ramane
            else:
                kind = "schimbare"          # 330, 360, 360 -> tarif nou
                base = p
            x = seq[i]
            if not in_period(x["data"]):
                continue
            out.append({
                "fel": kind, "proprietate": prop_display(x["nume"] or ""), "tip": tip,
                "facturare": mode, "factura": x["factura"], "fisier": x["fisier"], "client": x["client"],
                "data": x["data"], "pret": p, "pret_anterior": prev,
                "diferenta": round(p - prev, 2), "unitate": "€/h" if mode == "heure" else "€",
                # id stabil: aceeasi semnalare are acelasi id la fiecare calcul (pentru "E corect")
                "id": f"{key}|{tip}|{mode}|{client_k}|{x['factura'] or ''}|{p:g}|{prev:g}",
            })
    order = {"abatere": 0, "de confirmat": 1, "schimbare": 2}
    # abaterile primele; in fiecare categorie, cele mai recente primele
    out.sort(key=lambda r: (order[r["fel"]], -int((r["data"] or "0000-00-00").replace("-", ""))))
    return out
