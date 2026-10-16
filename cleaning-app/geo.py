"""
Etapa 2: distante de la baza la fiecare casa.

  - coordonatele vin din Google Sheet-ul utilizatorului; se pastreaza DOAR nume/lat/lng
    (orice alta coloana, ex. coduri de acces, e ignorata imediat si nu se salveaza nicaieri)
  - traseul pe drum (km + minute) vine de la serviciul public OSRM, intr-un numar mic de cereri
    (max 90 de case pe cerere); rezultatele se salveaza in backup si se refac doar la nevoie
  - daca serviciul de rute nu raspunde: estimare din linia dreapta x 1.35 (marcata "estimat")
"""
import csv
import io
import json
import math
import time
import urllib.request

OSRM_URL = "https://router.project-osrm.org/table/v1/driving/"
ROAD_FACTOR = 1.35          # drum / linie dreapta, tipic pentru drumurile din Corsica de sud
AVG_SPEED_KMH = 45          # pentru estimarea timpului cand nu avem traseu
BATCH = 90                  # case pe cerere (serverul public accepta ~100 de puncte)

ZONES = [(10, "sub 10 km"), (20, "10–20 km"), (35, "20–35 km"), (None, "peste 35 km")]

NAME_COLS = ("nume", "name", "nom", "casa", "proprietate", "propriete")
LAT_COLS = ("lat", "latitude", "latitudine")
LNG_COLS = ("lng", "lon", "long", "longitude", "longitudine")
ORE_COLS = ("ore", "ore estimate", "ore_estimate", "ore estimat", "ore/persoana", "ore persoana")


def parse_latlng(text):
    """'41.5955, 9.2746' -> (41.5955, 9.2746) sau None."""
    try:
        a, b = [float(x.strip()) for x in str(text).replace(";", ",").split(",")[:2]]
    except (ValueError, TypeError):
        return None
    if -90 <= a <= 90 and -180 <= b <= 180:
        return (a, b)
    return None


def _num(v):
    try:
        return float(str(v).strip().replace(",", "."))
    except (ValueError, TypeError):
        return None


def parse_coords_csv(text):
    """
    Citeste CSV-ul exportat din Sheet. Intoarce lista [{nume, lat, lng}] (+ "ore" daca exista coloana)
    - nimic altceva (codurile de acces si restul coloanelor sunt ignorate).
    Coloanele sunt gasite dupa titlu (nume/lat/lng sau echivalente).
    """
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return [], "Sheet gol"
    head = [h.strip().casefold() for h in rows[0]]

    def col(options):
        return next((i for i, h in enumerate(head) if h in options), None)
    i_name, i_lat, i_lng, i_ore = col(NAME_COLS), col(LAT_COLS), col(LNG_COLS), col(ORE_COLS)
    if i_name is None or i_lat is None or i_lng is None:
        return [], f"Nu gasesc coloanele nume/lat/lng (am gasit: {', '.join(rows[0][:8])})"
    out = []
    for r in rows[1:]:
        if max(i_name, i_lat, i_lng) >= len(r):
            continue
        name, lat, lng = r[i_name].strip(), _num(r[i_lat]), _num(r[i_lng])
        if name and lat is not None and lng is not None and -90 <= lat <= 90 and -180 <= lng <= 180:
            item = {"nume": name, "lat": round(lat, 6), "lng": round(lng, 6)}
            ore = _num(r[i_ore]) if i_ore is not None and i_ore < len(r) else None
            if ore is not None and 0 < ore <= 48:
                item["ore"] = round(ore, 2)
            out.append(item)
    return out, None


def haversine_km(a, b):
    lat1, lon1, lat2, lon2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((lat2 - lat1) / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 6371.0 * 2 * math.asin(math.sqrt(h))


def coord_key(lat, lng):
    return f"{lat:.5f},{lng:.5f}"


def osrm_routes(base, points, timeout=40, pause=1.1):
    """
    base: (lat, lng); points: lista de (lat, lng).
    Intoarce lista de (km, minute) sau None pentru punctele fara traseu.
    Ridica exceptie daca serviciul nu raspunde deloc.
    """
    out = []
    for start in range(0, len(points), BATCH):
        chunk = points[start:start + BATCH]
        coords = ";".join(f"{lng:.6f},{lat:.6f}" for lat, lng in [base] + chunk)
        url = f"{OSRM_URL}{coords}?sources=0&annotations=duration,distance"
        req = urllib.request.Request(url, headers={"User-Agent": "SimNettoyage-Facturi/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read())
        if data.get("code") != "Ok":
            raise RuntimeError(f"serviciul de rute: {data.get('code')} {data.get('message', '')}")
        dist, dur = data["distances"][0][1:], data["durations"][0][1:]
        for d, t in zip(dist, dur):
            out.append((round(d / 1000, 1), round(t / 60)) if d is not None and t is not None else None)
        if start + BATCH < len(points):
            time.sleep(pause)   # serverul public cere cel mult o cerere pe secunda
    return out


def estimate_route(base, point):
    km = haversine_km(base, point) * ROAD_FACTOR
    return round(km, 1), round(km / AVG_SPEED_KMH * 60)


def zone_of(km):
    if km is None:
        return None
    for limit, label in ZONES:
        if limit is None or km < limit:
            return label
    return ZONES[-1][1]


MAX_TABLE = 100             # cate puncte accepta serverul public intr-o matrice


def _osrm_table(points, sources=None, destinations=None, timeout=60):
    coords = ";".join(f"{lng:.6f},{lat:.6f}" for lat, lng in points)
    url = f"{OSRM_URL}{coords}?annotations=distance"
    if sources is not None:
        url += "&sources=" + ";".join(map(str, sources)) + "&destinations=" + ";".join(map(str, destinations))
    req = urllib.request.Request(url, headers={"User-Agent": "SimNettoyage-Facturi/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.loads(r.read())
    if data.get("code") != "Ok":
        raise RuntimeError(f"serviciul de rute: {data.get('code')} {data.get('message', '')}")
    return [[round(d / 1000, 1) if d is not None else None for d in row] for row in data["distances"]]


def osrm_matrix(points, timeout=60, pause=1.0):
    """
    Distantele pe drum intre TOATE punctele (baza + case). Intoarce matricea km [i][j]
    (None unde nu exista traseu). Ridica exceptie daca nu merge.
    Serverul public accepta ~100 de puncte pe cerere: peste atat, matricea se face pe bucati
    (cate 50 x 50 de puncte, o cerere pe secunda).
    """
    n = len(points)
    if n <= MAX_TABLE:
        return _osrm_table(points, timeout=timeout)
    half = MAX_TABLE // 2
    blocks = [list(range(i, min(i + half, n))) for i in range(0, n, half)]
    km = [[None] * n for _ in range(n)]
    first = True
    for a in blocks:
        for b in blocks:
            if not first:
                time.sleep(pause)
            first = False
            if a == b:
                sub = _osrm_table([points[i] for i in a], timeout=timeout)
            else:
                pts = [points[i] for i in a] + [points[j] for j in b]
                sub = _osrm_table(pts, list(range(len(a))), list(range(len(a), len(a) + len(b))), timeout=timeout)
            for x, i in enumerate(a):
                for y, j in enumerate(b):
                    km[i][j] = sub[x][y]
    return km
