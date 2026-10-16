"""
Cheltuieli si profit.

  - cheltuielile se introduc pe categorii (unele lunare, recurente)
  - categoriile "personale" (ex. transferuri acasa) NU scad din profitul firmei;
    apar separat, ca "ramas dupa transferuri"
  - motorina: estimata din traseele zilnice (baza -> casele zilei, de la una la alta -> baza);
    unde exista bonuri reale pentru o luna, profitul foloseste bonurile, nu estimarea
  - profitul pe casa e ESTIMAT: drumul casei + salarii/produse/mancare dupa numarul de menaje
    + restul cheltuielilor firmei dupa venit
"""
import calendar
import re
import unicodedata
from collections import defaultdict
from datetime import date

CATEGORII_IMPLICITE = [
    {"nume": "Motorină", "tip": "firma"},
    {"nume": "Salarii", "tip": "firma"},
    {"nume": "Produse curățenie", "tip": "firma"},
    {"nume": "Mâncare", "tip": "firma"},
    {"nume": "Chirie", "tip": "firma"},
    {"nume": "Taxe", "tip": "firma"},
    {"nume": "Comisioane bănci", "tip": "firma"},
    {"nume": "Amortizare investiție", "tip": "firma"},
    {"nume": "Altele", "tip": "firma"},
    {"nume": "Transferuri acasă / familie", "tip": "personal"},
]
# cheltuieli care cresc cu numarul de menaje (se impart pe case dupa interventii)
PER_MENAJ = ("salari", "produs", "mancare")


def _n(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def is_fuel(cat):
    c = _n(cat)
    return "motorin" in c or "combustib" in c or "carburant" in c


def cat_type(cat, categorii):
    for c in categorii:
        if _n(c["nume"]) == _n(cat):
            return c.get("tip") or "firma"
    return "firma"


def _add_months(mk, n):
    y, m = int(mk[:4]), int(mk[5:7]) - 1 + n
    return f"{y + m // 12:04d}-{m % 12 + 1:02d}"


def expand(entries, until_month):
    """Cheltuielile + aparitiile lunare ale celor recurente, pana la luna data (inclusiv)."""
    out = []
    for e in entries:
        d = e.get("data") or ""
        if len(d) < 10:
            continue
        if not e.get("lunar"):
            out.append({**e, "luna": d[:7], "aparitie": d, "serie": None})
            continue
        stop = min(until_month, e.get("pana_la") or until_month)
        mk, day = d[:7], int(d[8:10])
        while mk <= stop:
            y, m = int(mk[:4]), int(mk[5:7])
            dd = f"{mk}-{min(day, calendar.monthrange(y, m)[1]):02d}"
            out.append({**e, "luna": mk, "aparitie": dd, "serie": e["id"]})
            mk = _add_months(mk, 1)
    return out


# ------------------------------------------------------------------ motorina estimata

def daily_routes(ivs, house_point, base, dist):
    """
    Pe fiecare zi: baza -> casele zilei (cea mai apropiata urmatoare) -> baza.
    house_point(iv) -> (cheie casa, punct) sau None; dist(a, b) -> km.
    Intoarce: km pe zi, km pe casa (km-ii zilei impartiti dupa distanta casei de la baza), zile fara coordonate.
    """
    days = defaultdict(dict)
    missing = set()
    for iv in ivs:
        d = iv.get("data_iso")
        if not d:
            continue
        hp = house_point(iv)
        if hp:
            days[d][hp[0]] = hp[1]
        else:
            missing.add(d)
    km_day, km_house = {}, defaultdict(float)
    for d, houses in days.items():
        left, cur, total, legs = dict(houses), base, 0.0, []
        while left:
            k = min(left, key=lambda h: dist(cur, left[h]))
            leg = dist(cur, left[k])
            legs.append((k, leg))
            total += leg
            cur = left.pop(k)
        back = dist(cur, base)
        total += back
        # km-ii zilei se impart intre casele zilei proportional cu distanta fiecareia de la baza
        # (o casa departe "cere" mai mult drum). Nu depinde de ordinea in care le-ai facut:
        # inainte, casa din mijlocul unui grup de case apropiate primea doar 1-2 km.
        w = {k: max(dist(base, houses[k]), 0.5) for k, _ in legs}
        sw = sum(w.values()) or 1.0
        for k in w:
            km_house[k] += total * w[k] / sw
        km_day[d] = round(total, 1)
    return km_day, dict(km_house), len(missing - set(days))


def fuel_cost(km, consum, pret):
    return km * (consum or 0) / 100 * (pret or 0)


# ------------------------------------------------------------------ profit

def month_summary(months, venit_luna, occ, categorii, fuel_est_luna):
    """
    Pe fiecare luna: venit, cheltuieli firma (motorina reala daca exista, altfel estimata),
    profit, transferuri personale si ce ramane dupa ele.
    """
    rows = []
    for mk in months:
        firma = personal = fuel_real = 0.0
        pe_cat = defaultdict(float)
        for e in occ:
            if e["luna"] != mk:
                continue
            s = float(e.get("suma") or 0)
            pe_cat[e["categorie"]] += s
            if cat_type(e["categorie"], categorii) == "personal":
                personal += s
            elif is_fuel(e["categorie"]):
                fuel_real += s
            else:
                firma += s
        est = fuel_est_luna.get(mk, 0.0)
        fuel_used = fuel_real if fuel_real > 0 else est
        chelt = firma + fuel_used
        v = venit_luna.get(mk, 0.0)
        rows.append({
            "luna": mk, "venit": round(v, 2), "cheltuieli": round(chelt, 2),
            "motorina_reala": round(fuel_real, 2), "motorina_estimata": round(est, 2),
            "motorina_din": "bonuri" if fuel_real > 0 else "estimare",
            "profit": round(v - chelt, 2), "marja": round((v - chelt) / v * 100, 1) if v else None,
            "personal": round(personal, 2), "ramas": round(v - chelt - personal, 2),
            "pe_categorii": {k: round(x, 2) for k, x in pe_cat.items()},
        })
    return rows


def house_profit(props, km_house, fuel_per_km, pools):
    """
    props: [{key, total, interventii}]; pools: {"per_menaj": €, "general": €, "motorina_scalare": f}
    Motorina: km-ii casei x cost/km (scalat la bonuri, daca exista). Restul impartit
    dupa interventii (salarii, produse, mancare) si dupa venit (restul cheltuielilor firmei).
    """
    tot_int = sum(p["interventii"] for p in props) or 1
    tot_v = sum(p["total"] for p in props) or 1
    out = {}
    for p in props:
        fuel = km_house.get(p["key"], 0.0) * fuel_per_km * pools.get("motorina_scalare", 1.0)
        share = pools["per_menaj"] * p["interventii"] / tot_int + pools["general"] * p["total"] / tot_v
        cost = fuel + share
        out[p["key"]] = {"profit_est": round(p["total"] - cost, 2), "cost_est": round(cost, 2),
                         "motorina_est": round(fuel, 2), "km_drum": round(km_house.get(p["key"], 0.0), 1),
                         "marja_est": round((p["total"] - cost) / p["total"] * 100, 1) if p["total"] else None}
    return out


def per_menaj(cat):
    c = _n(cat)
    return any(w in c for w in PER_MENAJ)


def today_month():
    return date.today().isoformat()[:7]
