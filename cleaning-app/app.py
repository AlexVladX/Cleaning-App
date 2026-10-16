# Incarcam codec-ul 'idna' in firul principal, INAINTE de orice fir din fundal.
# Altfel, daca doua fire il incarca simultan (ex. scanarea + health check-ul Render),
# Python memoreaza esecul si TOATE cererile dau "LookupError: unknown encoding: idna".
import encodings.idna  # noqa: F401
"x".encode("idna")

import os
import re
import json
import io
import hmac
from functools import lru_cache
import socket
import threading
import time
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime, date, timedelta

from flask import Flask, render_template, jsonify, send_file, request, Response
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import google_auth_httplib2
import httplib2
import openpyxl
import analytics
import calendar_check
import costs
import formulas
import geo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from invoice_parser import (
    parse_invoice, compute_interventions, layout_text, parse_ddmmyyyy, ParseError, sane_invoice_date,
    norm_text, canonical_type, DEFAULT_RULES, TIP_LABEL,
)

app = Flask(__name__)

# ---------------------------------------------------------------- configurare
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")
POLL_INTERVAL = int(os.environ.get("POLL_INTERVAL", "120"))
# monitorizarea (facturi noi din Drive) porneste singura; MONITOR=0 o opreste
MONITOR_ON = os.environ.get("MONITOR", "1") != "0"

# AI folosit DOAR ca rezerva, cand o factura nu are sablonul SimNettoyage.
# Implicit: Groq (gratuit), model openai/gpt-oss-120b. Orice API compatibil OpenAI merge.
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
LLM_API_KEY = os.environ.get("LLM_API_KEY", "").strip() or GROQ_API_KEY
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.groq.com/openai/v1").rstrip("/")
LLM_MODEL = os.environ.get("LLM_MODEL", "openai/gpt-oss-120b")

# Optional: parola pentru dashboard (HTTP Basic Auth, orice utilizator)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

# Backup-ul scanarilor: un fisier JSON din Drive, creat de utilizator si partajat
# cu service account-ul ca EDITOR. Gasit dupa ID sau dupa nume in folderul facturilor.
CACHE_FILE_ID = os.environ.get("CACHE_FILE_ID", "").strip()
CACHE_FILE_NAME = os.environ.get("CACHE_FILE_NAME", "cache-facturi.json").strip()
CACHE_SAVE_EVERY = 20  # salvam progresul la fiecare N facturi procesate
# Schimbati valoarea cand se modifica PARSERUL (citirea PDF) -> se recitesc toate facturile.
# Regulile de calcul (setarile de mai jos) NU cer recitire: se recalculeaza din randurile salvate.
PARSER_VERSION = "2026-10-11c"  # c: an din 2 cifre, randuri fara fundal

cache = {"file_id": None, "loaded": False, "dirty": False, "enabled": False, "msg": "se încarcă"}

# Reguli de calcul + unificari de proprietati. Se editeaza din dashboard si se salveaza in backup.
settings = {
    **DEFAULT_RULES,
    "clienti_la_ora": [c.strip() for c in os.environ.get(
        "CLIENTI_LA_ORA", "Aquila Vacances;Esther").split(";") if c.strip()],
    "alias": {},   # cheie proprietate -> cheie proprietate canonica (unire manuala)
    "ignorate": [],  # sugestii de unire respinse ("Nu e dubla")
    "verificate": {},  # factura -> amprenta avertismentelor marcate "Am verificat"
    "preturi_ok": [],  # semnalari de pret marcate "E corect"
    "excluse": [],     # proprietati scoase manual din lista (ex. produse de curatenie facturate)
    "cuvinte_produse": ["produit", "fourniture"],  # randuri cu aceste cuvinte in nume/tip = produse, nu menaj
    "pastrate": [],    # proprietati puse inapoi manual: cuvintele de produse nu le mai scot
    # clienti grupati: [{"nume": "Client X", "potriviri": ["dupont", "martin", ...]}]
    # o factura intra in grup daca numele clientului SAU folderul din Drive contine o potrivire
    "grupuri_clienti": [],
    # incasari: id fisier factura -> {"suma": incasat, "data": "YYYY-MM-DD", "nota": ""}
    "incasari": {},
    # cheltuieli: [{"id", "data": "YYYY-MM-DD", "categorie", "suma", "nota", "lunar": bool, "pana_la": "YYYY-MM"}]
    "cheltuieli": [],
    "categorii_cheltuieli": [dict(c) for c in costs.CATEGORII_IMPLICITE],
    "consum": 8.0,             # litri / 100 km
    "pret_motorina": 1.95,     # € / litru
    "dist_matrix": {},         # distante pe drum intre toate casele (pentru motorina estimata)
    # Google Calendar cu interventiile (ID-ul se lipeste in aplicatie; evenimentele NU se salveaza)
    "calendar_id": "",
    "calendar_de_la": "",      # verificarea incepe de la data asta (optional)
    "calendar_ok": [],         # semnalari marcate "E corect"
    # ore estimate pe casa, pentru O persoana (salvate si in Sheet-ul cu coordonate, coloana "ore")
    "ore_estimate": {},
    "ore_nescrise": [],        # case modificate in aplicatie dar inca nescrise in Sheet (se reincearca)
    # indicatori proprii: [{"nume": "€ / oră", "expr": "incasari_fara_inter / ore_lucru"}]
    "formule": [dict(f) for f in formulas.DEFAULT_FORMULAS],
    # ordinea coloanelor aleasa de utilizator (trase cu mouse-ul) + indicatorii ascunsi
    # {"props": ["profit", "fx:€ / oră", ...], "clienti": [...], "ascunse": ["Marjă %"]}
    "coloane": {},
    # nume unite automat ("Appartement Marina" = "Marina") pe care utilizatorul le-a separat (case diferite)
    "separate": [],
    # intrebarile pastrate din tab-ul "Intrebari" (filtre/sortare pe indicatori + concluzii si istoric)
    "intrebari": [],
    # --- etapa 2: distante (salvate in backup, NU in cod)
    "baza": os.environ.get("BASE_COORDS", "41.595529046210864, 9.274690276554628"),
    "coord_sheet": os.environ.get("COORDS_SHEET_ID", ""),  # ID-ul Sheet-ului cu coordonate
    "coords": [],          # ultima citire: doar [{nume, lat, lng}]
    "coord_alias": {},     # cheie proprietate (facturi) -> nume din Sheet, legare manuala
    "distante": {},        # "lat,lng" -> [km, minute, sursa]
    "distante_baza": "",   # baza pentru care sunt calculate distantele
    "geo_sync": "",        # cand s-au citit ultima data coordonatele
}

# Cuvinte generice ignorate cand comparam numele proprietatilor ("Villa Blue" = "Blue")
GENERIC_PROP_WORDS = {"villa", "maison", "appartement", "appart", "appt", "apt", "app",
                      "chalet", "residence", "gite", "house", "casa", "studio", "logement",
                      "menage", "la", "le", "les", "l"}


@lru_cache(maxsize=20000)
def _auto_prop_key(name):
    words = norm_text(name).split()
    while len(words) > 1 and words[0] in GENERIC_PROP_WORDS:
        words = words[1:]
    return " ".join(words) or norm_text(name)


_sep_memo = [None, frozenset()]


def _separated():
    """Numele (normalizate) pe care utilizatorul le-a separat: nu se mai unesc automat."""
    lst = settings.get("separate")
    if lst is not _sep_memo[0]:
        _sep_memo[0], _sep_memo[1] = lst, frozenset(lst or [])
    return _sep_memo[1]


def base_prop_key(name):
    """Cheia automata (fara unirile manuale): 'Villa Blue' -> 'blue'. Numele separate isi pastreaza cheia intreaga."""
    sep = _separated()
    if sep:
        n = norm_text(name)
        if n in sep:
            return n
    return _auto_prop_key(name)


def prop_key(name):
    key = base_prop_key(name)
    seen = set()
    while key in settings["alias"] and key not in seen:  # urmam unirile manuale
        seen.add(key)
        key = settings["alias"][key]
    return key

def sheet_keys():
    """Cheile caselor cunoscute din Sheet-ul cu coordonate (inclusiv legaturile manuale)."""
    return ({prop_key(c["nume"]) for c in settings.get("coords") or []}
            | set(settings.get("coord_alias") or {}))


def _excl_cfg():
    kept = set(settings.get("pastrate") or [])
    return (set(settings.get("excluse") or []),
            [norm_text(w) for w in settings.get("cuvinte_produse") or [] if w.strip()],
            kept | sheet_keys(), kept)


def exclusion_reason(iv, cfg=None):
    """
    'manual' (scos cu ✕) / 'produs' (cuvant din Reguli) / None.
    Siguranta: cuvintele NU scot niciodata o casa din Sheet-ul cu coordonate
    sau una pusa inapoi manual - doar butonul ✕ poate face asta.
    """
    excl, words, safe, kept = cfg or _excl_cfg()
    name = re.sub(r"\s+", " ", (iv.get("proprietate") or "").strip()) or "Necunoscut"
    key = prop_key(name)
    if key in excl:
        return "manual"
    if words:
        # tipul randului e "produs" (ex. Villa Emma · Produits) -> doar randul acela iese, casa ramane
        tip = norm_text(iv.get("tip_serviciu") or "")
        if tip and key not in kept and any(w in tip for w in words):
            return "produs"
        # numele e un produs -> iese, dar NICIODATA o casa din Sheet sau una pusa inapoi manual
        if key not in safe and any(w in norm_text(name) for w in words):
            return "produs"
    return None


def split_ivs(ivs):
    """(interventii de menaj, randuri scoase din lista)."""
    cfg = _excl_cfg()
    if not cfg[0] and not cfg[1]:
        return ivs, []
    keep, out = [], []
    for iv in ivs:
        (out if exclusion_reason(iv, cfg) else keep).append(iv)
    return keep, out


def client_group_rules():
    out = []
    for g in settings.get("grupuri_clienti") or []:
        words = [norm_text(w) for w in g.get("potriviri") or [] if norm_text(w)]
        if g.get("nume") and words:
            out.append((g["nume"], words))
    return out


def group_of(client, folder_norm, rules):
    c = norm_text(client or "")
    return next((name for name, words in rules
                 if any(w in c or (folder_norm and w in folder_norm) for w in words)), None)


def group_clients(ivs):
    """
    Inlocuieste clientul facturat cu numele grupului (ex. persoane fizice aduse de acelasi
    colaborator). Clientul de pe factura ramane in 'client_factura'. Intoarce copii.
    """
    rules = client_group_rules()
    if not rules:
        return ivs
    with state_lock:
        folders = {fid: norm_text(f.get("folder") or "") for fid, f in state["files"].items()}
    out, memo = [], {}
    for iv in ivs:
        client = iv.get("client") or ""
        folder = folders.get(iv.get("_file_id"), "")
        mk = (client, folder)
        if mk not in memo:
            memo[mk] = group_of(client, folder, rules)
        grp = memo[mk]
        out.append({**iv, "client": grp, "client_factura": client} if grp else iv)
    return out


def analysis_ivs(ivs):
    """Interventiile folosite in indicatori: fara produse, cu clientii grupati."""
    keep, ex = split_ivs(ivs)
    return group_clients(keep), ex


import copy as _copy
DEFAULT_SETTINGS = _copy.deepcopy(settings)   # setarile implicite (pentru restaurarea din backup vechi)

state = {
    "files": {},
    "interventii": [],
    "last_scan": None,
    "monitoring": MONITOR_ON,
    "log": [],
    "scan": {"active": False, "phase": None, "done": 0, "total": 0, "eta": None},
    "calendar": {"events": [], "citit": None, "eroare": None},   # doar in memorie
}
state_lock = threading.Lock()
scan_lock = threading.Lock()


def add_log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)  # apare si in Render → Logs
    with state_lock:
        state["log"].append(f"[{ts}] {msg}")
        if len(state["log"]) > 200:
            state["log"] = state["log"][-200:]


def set_stage(txt):
    """Pasul curent al scanarii — afisat in dashboard si raportat daca dureaza prea mult."""
    with state_lock:
        state["scan"]["stage"] = txt
        state["scan"]["stage_since"] = time.time()


def _watchdog(scan_id):
    """Daca un pas dureaza > 20s, scrie in jurnal la ce anume asteapta aplicatia."""
    last = 0.0
    while True:
        time.sleep(5)
        with state_lock:
            sc = state["scan"]
            if not sc.get("active") or sc.get("id") != scan_id:
                return
            stage, since = sc.get("stage"), sc.get("stage_since") or time.time()
        waited = time.time() - since
        if waited >= 20 and time.time() - last >= 20:
            add_log(f"… inca astept: {stage} ({waited:.0f}s)")
            last = time.time()


# ---------------------------------------------------------------- autentificare
@app.before_request
def require_password():
    if not APP_PASSWORD or request.path == "/healthz":
        return None
    auth = request.authorization
    if auth and auth.password and hmac.compare_digest(auth.password, APP_PASSWORD):
        return None
    return Response("Autentificare necesara", 401,
                    {"WWW-Authenticate": 'Basic realm="Facturi"'})


# ---------------------------------------------------------------- retea
# Pe unele servere (ex. Render gratuit) conexiunile IPv6 spre Google nu raspund si
# raman agatate in loc sa esueze rapid. Folosim doar IPv4 (Google are IPv4 peste tot).
# originalul e pastrat pe modulul socket, ca o reincarcare a app.py sa nu-l "infasoare" de doua ori
if not hasattr(socket, "_getaddrinfo_original"):
    socket._getaddrinfo_original = socket.getaddrinfo
_orig_getaddrinfo = socket._getaddrinfo_original
FORCE_IPV4 = os.environ.get("FORCE_IPV4", "1") == "1"


def _getaddrinfo_ipv4(host, port, family=0, type=0, proto=0, flags=0):
    res = _orig_getaddrinfo(host, port, family, type, proto, flags)
    v4 = [r for r in res if r[0] == socket.AF_INET]
    return v4 or res


if FORCE_IPV4:
    socket.getaddrinfo = _getaddrinfo_ipv4


def _run_with_timeout(fn, seconds, what):
    """Ruleaza fn() cu limita stricta de timp (si pentru DNS, unde timeout-ul socket nu ajuta)."""
    box = {}

    def target():
        try:
            box["ok"] = fn()
        except Exception as e:  # noqa: BLE001
            box["err"] = e
    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join(seconds)
    if t.is_alive():
        raise TimeoutError(f"{what}: niciun raspuns in {seconds}s")
    if "err" in box:
        raise box["err"]
    return box.get("ok")


def net_check(host="oauth2.googleapis.com", port=443):
    """Test de retea: DNS + conexiune TCP pe fiecare adresa (IPv4 si IPv6)."""
    parts = []
    t = time.time()
    try:
        addrs = _run_with_timeout(lambda: _orig_getaddrinfo(host, port, 0, socket.SOCK_STREAM), 10, "DNS")
    except Exception as e:  # noqa: BLE001
        return f"{host}: DNS esuat ({e})"
    parts.append(f"DNS {(time.time() - t) * 1000:.0f}ms")
    seen = []
    for fam, _, _, _, sa in addrs:
        if sa[0] in seen or len(seen) >= 4:
            continue
        seen.append(sa[0])
        kind = "IPv6" if fam == socket.AF_INET6 else "IPv4"
        s = None
        t = time.time()
        try:
            s = socket.socket(fam, socket.SOCK_STREAM)  # poate esua daca serverul nu are IPv6 deloc
            s.settimeout(4)
            s.connect(sa)
            parts.append(f"{kind} ok {(time.time() - t) * 1000:.0f}ms")
        except Exception as e:  # noqa: BLE001
            parts.append(f"{kind} ESUAT ({type(e).__name__})")
        finally:
            if s is not None:
                s.close()
    return f"{host}: " + " · ".join(parts)


# ---------------------------------------------------------------- Google Drive
_drive_local = threading.local()
_net_checked = False


def get_drive_service():
    """
    Conexiunea la Drive se creeaza o singura data per fir de executie si se refoloseste.
    (Inainte se reconecta pentru FIECARE factura — foarte lent pe serverele gratuite.)
    Timeout de 60s ca o cerere blocata sa nu tina scanarea "in curs" la nesfarsit.
    """
    svc = getattr(_drive_local, "service", None)
    if svc is not None:
        return svc
    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("GOOGLE_CREDENTIALS_JSON nu e setat in Render → Environment")
    t0 = time.time()
    set_stage("citesc cheia Google (GOOGLE_CREDENTIALS_JSON)")
    try:
        info = json.loads(GOOGLE_CREDENTIALS_JSON)
    except ValueError as e:
        raise ValueError(f"GOOGLE_CREDENTIALS_JSON nu e JSON valid ({e}) — lipiti tot continutul fisierului cheii")
    creds = service_account.Credentials.from_service_account_info(
        info,
        # 'drive' (nu doar readonly) ca sa putem actualiza fisierul de backup.
        # Service account-ul vede/modifica tot doar ce i-ati partajat explicit.
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    global _net_checked
    if not _net_checked:
        _net_checked = True
        set_stage("test de retea spre Google")
        add_log(f"Retea: {net_check()}" + (" · folosesc doar IPv4" if FORCE_IPV4 else ""))
    # autentificare explicita, separat de prima cerere -> erorile de cheie se vad clar
    set_stage("autentificare la Google (oauth2.googleapis.com)")
    try:
        _run_with_timeout(lambda: creds.refresh(google_auth_httplib2.Request(httplib2.Http(timeout=20))),
                          45, "Autentificarea la Google")
    except TimeoutError as e:
        raise TimeoutError(f"{e} — conexiunea de la server la Google nu raspunde (vezi linia 'Retea:')")
    add_log(f"Autentificat ca {info.get('client_email')} ({time.time() - t0:.1f}s)")
    set_stage("incarc API-ul Google Drive")
    http = google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=60))
    svc = build("drive", "v3", http=http, cache_discovery=False)
    add_log(f"Conectat la Google Drive ({time.time() - t0:.1f}s)")
    _drive_local.service = svc
    return svc


# erorile de moment ale Google (500 "Internal Error", 503, 429) se reincearca singure, cu pauze crescatoare
DRIVE_RETRIES = 4


def list_pdf_files(folder_id):
    service = get_drive_service()
    all_files = []
    folders = [0]

    def search_folder(fid, label):
        page_token, here, subfolders = None, 0, []
        folders[0] += 1
        set_stage(f"listez folderul '{label}' (folderul nr. {folders[0]})")
        if folders[0] % 10 == 0:
            add_log(f"Foldere parcurse: {folders[0]} · {len(all_files)} PDF gasite")
        while True:  # paginare: altfel se pierd fisierele peste primele 100
            results = service.files().list(
                q=f"'{fid}' in parents and trashed=false",
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, size, md5Checksum)",
                pageSize=1000,
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute(num_retries=DRIVE_RETRIES)
            for f in results.get("files", []):
                if f["mimeType"] == "application/pdf":
                    f["folder"] = "" if label == "." else label
                    all_files.append(f)
                    here += 1
                elif f["mimeType"] == "application/vnd.google-apps.folder":
                    subfolders.append(f)
            page_token = results.get("nextPageToken")
            if not page_token:
                break
        with state_lock:
            state["scan"]["found"] = len(all_files)
        if here:
            add_log(f"📁 {label}: {here} PDF · {len(all_files)} gasite pana acum")
        for sf in subfolders:
            search_folder(sf["id"], f"{label}/{sf['name']}" if label != "." else sf["name"])

    add_log("Citesc lista de facturi din Drive...")
    search_folder(folder_id, ".")
    if not all_files:
        add_log(f"⚠ 0 PDF-uri in {folders[0]} foldere. Verificati DRIVE_FOLDER_ID si ca folderul "
                f"e partajat (Share) cu email-ul service account-ului de mai sus.")
    return all_files


def download_pdf(file_id):
    service = get_drive_service()
    req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=DRIVE_RETRIES)
    buf.seek(0)
    return buf


def file_signature(f):
    """Amprenta continutului PDF: se schimba doar daca se schimba fisierul (nu la redenumire)."""
    return f.get("md5Checksum") or f.get("modifiedTime") or ""


# ---------------------------------------------------------------- backup in Drive
CACHE_KEYS = ("id", "name", "sig", "status", "error", "avertismente", "interventii",
              "numar", "client", "data_emitere", "total", "metoda", "processed_at",
              "randuri", "antet", "av_factura", "folder")
RULE_KEYS = ("prag_forfait", "tipuri_la_ora", "clienti_la_ora")
PERSIST_KEYS = ("alias", "ignorate", "verificate", "preturi_ok", "excluse", "cuvinte_produse", "pastrate", "grupuri_clienti", "incasari", "cheltuieli",
                "categorii_cheltuieli", "consum", "pret_motorina", "dist_matrix",
                "calendar_id", "calendar_de_la", "calendar_ok", "ore_estimate", "ore_nescrise", "formule", "coloane", "intrebari", "separate", "baza", "coord_sheet", "coords",
                "coord_alias", "distante", "distante_baza", "geo_sync")


def current_rules():
    return {k: settings[k] for k in RULE_KEYS}


def recompute_all():
    """
    Recalculeaza interventiile tuturor facturilor din randurile salvate, cu regulile curente.
    Nu descarca si nu reciteste niciun PDF — dureaza o fractiune de secunda.
    """
    rules = current_rules()
    with state_lock:
        # reevaluam si duplicatele (dupa continut), in ordinea in care au fost citite facturile
        active = sorted((f for f in state["files"].values()
                         if f.get("status") in ("done", "warning", "duplicate") and f.get("randuri")),
                        key=lambda f: (f.get("processed_at") or 0, f.get("name") or ""))
        seen = {}
        for f in active:
            k = entry_dup_key(f)
            if k in seen:
                f["status"] = "duplicate"
                f["avertismente"] = [f"Duplicat al fisierului {seen[k]['name']} (acelasi continut) — ignorat"]
                f["interventii"] = []
                continue
            seen[k] = f
            ivs, calc_w = compute_interventions(f["randuri"], f.get("antet") or {}, rules)
            for iv in ivs:
                iv["_source"], iv["_file_id"] = f["name"], f["id"]
            f["interventii"] = ivs
            f["avertismente"] = calc_w + list(f.get("av_factura") or [])
            f["status"] = "warning" if f["avertismente"] else "done"
        state["interventii"] = [iv for f in state["files"].values()
                                if f.get("status") in ("done", "warning")
                                for iv in f.get("interventii") or []]
    cache["dirty"] = True


def apply_saved_settings(saved):
    if not isinstance(saved, dict):
        return
    for k in RULE_KEYS + PERSIST_KEYS:
        if k in saved:
            settings[k] = saved[k]


def merge_old_settings(old):
    """
    Pune inapoi setarile dintr-un backup vechi FARA sa strice ce ai acum:
    ce e gol / implicit acum se ia din backup-ul vechi; listele si dictionarele se UNESC
    (la conflict castiga valoarea de acum). Intoarce ce s-a recuperat, pe scurt.
    """
    defaults = DEFAULT_SETTINGS or {}
    got = []

    def empty_or_default(k):
        v = settings.get(k)
        return v in (None, "", [], {}) or (k in defaults and v == defaults[k])

    for k in ("coord_sheet", "baza", "calendar_id", "calendar_de_la", "consum", "pret_motorina",
              "prag_forfait", "tipuri_la_ora", "clienti_la_ora", "categorii_cheltuieli", "grupuri_clienti",
              "coords", "dist_matrix", "distante_baza", "geo_sync", "cuvinte_produse", "formule", "coloane", "intrebari", "separate"):
        if k in old and old[k] not in (None, "", [], {}) and empty_or_default(k) and old[k] != settings.get(k):
            settings[k] = old[k]
            got.append(k)
    for k in ("alias", "coord_alias", "verificate", "incasari", "ore_estimate", "distante"):
        o = old.get(k)
        if isinstance(o, dict) and o:
            cur = settings.get(k) or {}
            new = {**o, **cur}
            if len(new) > len(cur):
                got.append(f"{k} (+{len(new) - len(cur)})")
            settings[k] = new
    for k in ("ignorate", "preturi_ok", "excluse", "pastrate", "calendar_ok"):
        o = old.get(k)
        if isinstance(o, list) and o:
            cur = list(settings.get(k) or [])
            add = [x for x in o if x not in cur]
            if add:
                settings[k] = cur + add
                got.append(f"{k} (+{len(add)})")
    o = old.get("cheltuieli")
    if isinstance(o, list) and o:
        cur = list(settings.get("cheltuieli") or [])
        ids = {e.get("id") for e in cur}
        add = [e for e in o if isinstance(e, dict) and e.get("id") not in ids]
        if add:
            settings["cheltuieli"] = cur + add
            got.append(f"cheltuieli (+{len(add)})")
    return got


def find_cache_file():
    if CACHE_FILE_ID:
        return CACHE_FILE_ID
    name = CACHE_FILE_NAME.replace("'", "\\'")
    svc = get_drive_service()
    set_stage(f"caut fisierul de backup '{CACHE_FILE_NAME}'")
    res = svc.files().list(
        q=f"name = '{name}' and '{FOLDER_ID}' in parents and trashed = false",
        fields="files(id, name, size, capabilities/canEdit)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute(num_retries=DRIVE_RETRIES)
    found = res.get("files", [])
    if not found:
        return None
    if not found[0].get("capabilities", {}).get("canEdit"):
        cache["msg"] = f"{CACHE_FILE_NAME} există, dar service account-ul nu e Editor pe el"
        return None
    cache["file_size"] = int(found[0].get("size") or 0)
    return found[0]["id"]


def load_cache():
    """Incarca scanarile salvate. Fara backup -> aplicatia merge ca inainte (doar in memorie)."""
    cache["loaded"] = True
    try:
        fid = find_cache_file()
    except Exception as e:
        # retea / Google: NU continuam cu regulile implicite (s-ar pierde setarile), reincercam
        cache["loaded"] = False
        cache["msg"] = f"nu am putut căuta backup-ul în Drive ({e}) — reîncerc automat"
        raise RuntimeError(cache["msg"]) from e
    if not fid:
        msg = cache["msg"] if "Editor" in (cache.get("msg") or "") else f"nu găsesc '{CACHE_FILE_NAME}' în folderul facturilor"
        if msg != cache.get("msg"):
            add_log(f"⚠ Backup dezactivat: {msg} — reîncerc la fiecare scanare")
        cache["msg"] = msg
        return
    cache.update({"file_id": fid, "enabled": True})
    try:
        set_stage("descarc backup-ul")
        raw = download_pdf(fid).read().decode("utf-8").strip()
    except Exception as e:
        # problema de retea / Google: NU refacem backup-ul de la zero (s-ar pierde tot), reincercam
        cache["loaded"] = cache["enabled"] = False
        cache["msg"] = f"nu am putut descărca backup-ul ({e}) — reîncerc automat"
        raise RuntimeError(cache["msg"]) from e
    try:
        data = json.loads(raw) if raw else {}
        if not isinstance(data, dict):
            raise ValueError("format neasteptat")
    except ValueError as e:
        if raw or cache.get("file_size"):
            # fisierul are continut, dar l-am primit stricat: reincercam, NU il suprascriem
            cache["loaded"] = cache["enabled"] = False
            cache["msg"] = f"backup-ul s-a descărcat incomplet ({e}) — reîncerc automat"
            raise RuntimeError(cache["msg"]) from e
        data = {}
    if not data and cache.get("file_size"):
        cache["loaded"] = cache["enabled"] = False
        cache["msg"] = "backup-ul a venit gol din Drive, deși fișierul are conținut — reîncerc automat"
        raise RuntimeError(cache["msg"])
    apply_saved_settings(data.get("setari"))  # setarile raman valabile si cand se schimba parserul
    cache["msg"] = ""
    add_log(f"Reguli și setări încărcate din backup ({len(data.get('setari') or {})} chei)")
    if data.get("parser_version") != PARSER_VERSION:
        if data.get("files"):
            add_log("Parserul s-a actualizat — recitesc toate facturile o singura data")
        cache["dirty"] = True
        return
    files = data.get("files") or {}
    with state_lock:
        for fid_, entry in files.items():
            entry.setdefault("interventii", [])
            state["files"][fid_] = entry
        state["interventii"] = [iv for f in state["files"].values()
                                if f.get("status") in ("done", "warning")
                                for iv in f.get("interventii") or []]
    recompute_all()
    cache["dirty"] = False
    cache["known_files"] = len(files)
    add_log(f"Backup incarcat: {len(files)} facturi (salvat {data.get('saved_at', '?')})")


def save_cache(force=False, allow_empty=False):
    if not cache["enabled"] or not (cache["dirty"] or force):
        return
    if not cache["loaded"]:
        return          # backup-ul nu a fost inca citit: nu-l suprascriem cu o stare incompleta
    with state_lock:
        files = {fid: {k: f.get(k) for k in CACHE_KEYS}
                 for fid, f in state["files"].items()
                 if f.get("status") in ("done", "warning", "duplicate", "error")}
        refuse = not files and cache.get("known_files") and not allow_empty
        if not refuse:
            cache["dirty"] = False
    if refuse:   # (mesajul in afara lacatului: add_log il foloseste si el)
        add_log("⚠ Protecție: nu suprascriu backup-ul cu o listă goală de facturi")
        return
    payload = json.dumps({"parser_version": PARSER_VERSION,
                          "saved_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                          "setari": {k: settings[k] for k in RULE_KEYS + PERSIST_KEYS},
                          "files": files}, ensure_ascii=False, default=str)
    try:
        get_drive_service().files().update(
            fileId=cache["file_id"],
            media_body=MediaIoBaseUpload(io.BytesIO(payload.encode("utf-8")),
                                         mimetype="application/json", resumable=False),
            supportsAllDrives=True,
        ).execute(num_retries=DRIVE_RETRIES)
    except Exception as e:
        cache["dirty"] = True
        add_log(f"⚠ Nu am putut salva backup-ul: {e}")
        return
    cache["known_files"] = len(files)


# ---------------------------------------------------------------- rezerva AI
def llm_extract(pdf_bytes, max_retries=3):
    """
    Pentru facturi care NU au sablonul SimNettoyage. AI-ul doar copiaza datele
    brute (client, randuri, date); calculul interventiilor se face tot in cod.
    """
    text = layout_text(pdf_bytes)
    text = re.sub(r"[ \t]{3,}", " | ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)[:12000]

    prompt = (
        "Extract the data from this French cleaning-company invoice. "
        "The issuer is SimNettoyage — ignore its address and bank details.\n"
        "Return ONLY a JSON object, no markdown, with this shape:\n"
        '{"client": "name right under \'Facture pour\' (first line only, no address)",'
        ' "numar": "invoice number after \'Facture #\'",'
        ' "data_emitere": "DD/MM/YYYY from \'Soumis le\'",'
        ' "randuri": [{"nom": "property name", "type": "service type", '
        '"dates": ["dates exactly as written"], "qty": number, "pu": unit price number, '
        '"total": line total number}]}\n'
        "Use the 'Détail des prestations' table if present (one entry per property row; "
        "a row can have several dates). Skip empty €0.00 rows. Copy numbers, do not compute.\n\n"
        "INVOICE:\n" + text
    )
    body = json.dumps({
        "model": LLM_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0,
        "max_tokens": 4000,
    }).encode()

    for attempt in range(1, max_retries + 1):
        req = urllib.request.Request(
            f"{LLM_BASE_URL}/chat/completions", data=body,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {LLM_API_KEY}",
                     "User-Agent": "Cleaning-App/2.0"},
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = json.loads(resp.read())
            raw = (data["choices"][0]["message"].get("content") or "").strip()
            s, e = raw.find("{"), raw.rfind("}") + 1
            if s < 0 or e <= s:
                raise ValueError("raspuns AI fara JSON")
            return json.loads(raw[s:e])
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < max_retries:
                wait = int(e.headers.get("retry-after") or 20)
                add_log(f"Limita AI atinsa, astept {wait}s (incercarea {attempt}/{max_retries})")
                time.sleep(min(wait, 60))
                continue
            raise
        except (ValueError, KeyError) as e:
            if attempt < max_retries:
                add_log(f"Raspuns AI invalid ({e}), reincerc...")
                continue
            raise
    raise RuntimeError("AI indisponibil")


def parse_with_ai(pdf_bytes):
    data = llm_extract(pdf_bytes)
    header = {
        "client": (data.get("client") or "").strip() or None,
        "numar": (data.get("numar") or "").strip() or None,
        "data_emitere": parse_ddmmyyyy(data.get("data_emitere") or ""),
    }
    rows = []
    for r in data.get("randuri") or []:
        def num(v):
            try:
                return float(str(v).replace("€", "").replace(",", ".").strip()) if v not in (None, "") else None
            except ValueError:
                return None
        rows.append({"nom": (r.get("nom") or "").strip(), "type": r.get("type"),
                     "dates_raw": [str(d) for d in (r.get("dates") or [])],
                     "qty": num(r.get("qty")), "pu": num(r.get("pu")), "total": num(r.get("total"))})
    rows = [r for r in rows if r["nom"] or r["total"]]
    interventii, warnings = compute_interventions(rows, header, current_rules())
    total = round(sum(r["total"] or 0 for r in rows), 2)
    ai_note = ["Citita cu AI (sablon necunoscut) — verificati manual"]
    return {
        "metoda": "ai",
        "antet": {**header, "data_emitere": header["data_emitere"].strftime("%d/%m/%Y")
                  if header["data_emitere"] else None},
        "randuri": rows,
        "totaluri": {"subtotal": total, "ajustari": 0.0, "total": total},
        "interventii": interventii,
        "avertismente": ai_note + warnings,
        "avertismente_factura": ai_note,
    }


# ---------------------------------------------------------------- procesare
def content_key(client, total, randuri, numar=None, data_emitere=None):
    """
    Duplicat = aceeasi factura urcata de doua ori: acelasi numar, aceeasi data de emitere,
    acelasi client, acelasi total si ACELEASI randuri (casa, date de interventie, suma).
    Facturi diferite cu acelasi numar (numerotare refolosita) NU sunt duplicate.
    """
    c = norm_text(client)
    try:
        tot = round(float(total or 0), 2)
    except (TypeError, ValueError):
        tot = 0.0
    if randuri:
        rows = tuple(sorted((norm_text(r.get("nom")), tuple(sorted(r.get("dates_raw") or [])),
                             round(float(r.get("total") or 0), 2)) for r in randuri))
        return ("continut", (numar or "").upper(), data_emitere or "", c, tot, rows)
    return ("numar", (numar or "").upper(), c, tot)  # fara randuri (rar): rezerva veche


def entry_dup_key(f):
    return content_key(f.get("client"), f.get("total"), f.get("randuri"), f.get("numar"),
                       f.get("data_emitere"))


MAX_PDF_MB = 12        # PDF-uri mai mari sunt aproape sigur scanari/imagini: nu le deschidem
MEM_LIMIT_MB = 400     # serverul gratuit are 512 MB; peste prag oprim citirea si continuam data viitoare


def rss_mb():
    try:
        with open("/proc/self/status") as fh:
            for line in fh:
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) // 1024
    except OSError:
        pass
    return 0


def process_file(file_info, idx=None, total=None):
    file_id, name = file_info["id"], file_info["name"]
    pre = f"[{idx}/{total}] " if idx else ""
    rest = f" · mai sunt {total - idx}" if idx and total - idx > 0 else (" · ultima" if idx else "")
    with state_lock:
        state["scan"]["current"] = name
        entry = state["files"].setdefault(file_id, {})
        entry.update({
            "id": file_id, "name": name, "status": "processing",
            "sig": file_signature(file_info) or entry.get("sig"),
            "error": None, "avertismente": [],
        })
        entry.setdefault("interventii", [])
    used_ai = False
    try:
        size_mb = int(file_info.get("size") or 0) / 1024 / 1024
        if size_mb > MAX_PDF_MB:
            raise ValueError(f"PDF prea mare ({size_mb:.0f} MB) — probabil scanat; nu îl deschid ca să nu "
                             f"blochez serverul (limita {MAX_PDF_MB} MB)")
        set_stage(f"descarc {name}")
        pdf_bytes = download_pdf(file_id)
        set_stage(f"citesc {name}")
        try:
            result = parse_invoice(pdf_bytes, current_rules())
        except ParseError as pe:
            if not LLM_API_KEY:
                raise ValueError(f"{pe} (nicio cheie AI pentru rezerva)")
            add_log(f"{name}: {pe} — incerc cu AI ({LLM_MODEL})")
            used_ai = True
            result = parse_with_ai(pdf_bytes)

        interventii = result["interventii"]
        for iv in interventii:
            iv["_source"] = name
            iv["_file_id"] = file_id
        a, t = result["antet"], result["totaluri"]
        warnings = list(result["avertismente"])

        with state_lock:
            others = [f for fid, f in state["files"].items()
                      if fid != file_id and f.get("status") in ("done", "warning")]
            key = content_key(a.get("client"), t.get("total"), result["randuri"], a.get("numar"),
                              a.get("data_emitere"))
            dup = next((f for f in others if entry_dup_key(f) == key), None)

            entry = state["files"][file_id]
            entry.update({
                "numar": a.get("numar"), "client": a.get("client"),
                "data_emitere": a.get("data_emitere"), "total": t.get("total"),
                "metoda": result["metoda"],
                # date brute pastrate in backup -> regulile se pot schimba fara recitirea PDF-ului
                "randuri": result["randuri"], "antet": a,
                "av_factura": result.get("avertismente_factura", []),
            })
            if dup:
                entry["status"] = "duplicate"
                entry["avertismente"] = [f"Duplicat al fisierului {dup['name']} (acelasi continut) — ignorat"]
                entry["interventii"] = []
                state["interventii"] = [iv for iv in state["interventii"] if iv.get("_file_id") != file_id]
            else:
                entry["status"] = "warning" if warnings else "done"
                entry["avertismente"] = warnings
                entry["interventii"] = interventii
                state["interventii"] = [
                    iv for iv in state["interventii"] if iv.get("_file_id") != file_id
                ] + interventii

        if dup:
            add_log(f"{pre}= {name}: duplicat al {dup['name']}, ignorat{rest}")
        else:
            add_log(f"{pre}✓ {name}: {a.get('numar')} · {a.get('client')} · "
                    f"{len(interventii)} interventii · {t.get('total')} €"
                    + (f" · {len(warnings)} avertismente" if warnings else "") + rest)
            for w in warnings:
                add_log(f"  ⚠ {w}")
    except Exception as e:
        with state_lock:
            state["files"][file_id]["status"] = "error"
            state["files"][file_id]["error"] = str(e)
            state["files"][file_id]["interventii"] = []
            state["interventii"] = [iv for iv in state["interventii"] if iv.get("_file_id") != file_id]
        add_log(f"{pre}✗ Eroare la {name}: {e}{rest}")
    with state_lock:
        state["files"][file_id]["processed_at"] = time.time()
    cache["dirty"] = True
    return used_ai


def process_queue_background(new_files):
    n = len(new_files)
    t0 = time.time()
    add_log(f"Incep citirea a {n} facturi...")
    with state_lock:
        state["scan"].update({"done": 0, "total": n, "eta": None})
    import gc
    for i, f in enumerate(new_files):
        if rss_mb() > MEM_LIMIT_MB:
            gc.collect()
            if rss_mb() > MEM_LIMIT_MB:
                add_log(f"⚠ Memorie aproape plină ({rss_mb()} MB din 512) — mă opresc aici ca serverul să nu "
                        f"cadă; continui cu restul de {n - i} facturi la următoarea scanare")
                with state_lock:
                    for g in new_files[i:]:
                        e = state["files"].get(g["id"])
                        if e and e.get("status") == "new":
                            e["sig"] = None     # ramane "de citit" la scanarea urmatoare
                break
        used_ai = process_file(f, i + 1, n)
        gc.collect()
        elapsed = time.time() - t0
        eta = round(elapsed / (i + 1) * (n - i - 1))
        with state_lock:
            state["scan"].update({"done": i + 1, "total": n, "eta": eta})
        if (i + 1) % 10 == 0 and i + 1 < n:
            add_log(f"— Progres {i + 1}/{n} · mai sunt {n - i - 1} · "
                    f"{elapsed / (i + 1):.1f}s per factura · ~{max(1, round(eta / 60))} min ramase · memorie {rss_mb()} MB")
        if (i + 1) % CACHE_SAVE_EVERY == 0 and i + 1 < n:
            save_cache()  # daca serverul adoarme in mijlocul scanarii, nu pierdem progresul
        if used_ai and i < n - 1:
            time.sleep(3)  # menajam limita gratuita a AI-ului
    add_log(f"Toate fisierele procesate in {time.time() - t0:.0f}s.")


def scan_folder(retry_errors=False):
    """
    retry_errors: reincearca si facturile care au dat eroare (la pornire si la scanare
    manuala). Monitorizarea automata nu le reincearca, ca sa nu le descarce la fiecare minut.
    """
    if not FOLDER_ID:
        add_log("DRIVE_FOLDER_ID nu e setat")
        return False
    if not scan_lock.acquire(blocking=False):
        with state_lock:
            stage, since = state["scan"].get("stage"), state["scan"].get("stage_since") or time.time()
        add_log(f"Scanare deja in curs — pasul curent: {stage} ({time.time() - since:.0f}s)")
        return None
    scan_id = time.time()
    with state_lock:
        state["scan"] = {"active": True, "phase": "listare", "done": 0, "total": 0,
                         "eta": None, "found": 0, "current": None, "id": scan_id,
                         "stage": "pornire", "stage_since": time.time(),
                         "thread": threading.get_ident()}
    threading.Thread(target=_watchdog, args=(scan_id,), daemon=True).start()
    add_log("Pornesc scanarea...")
    try:
        if not cache["loaded"] or not cache["enabled"]:
            load_cache()      # si cand backup-ul lipsea: il cautam din nou la fiecare scanare
        files = list_pdf_files(FOLDER_ID)
        with state_lock:
            state["scan"]["phase"] = "procesare"
            known = {fid: dict(f) for fid, f in state["files"].items()}  # copie, nu referinte
        in_drive = {f["id"] for f in files}
        if not files and known:
            # Drive a intors 0 PDF-uri desi aveam facturi: aproape sigur o problema temporara de acces
            add_log("⚠ Drive a întors 0 facturi (acces sau rețea) — nu șterg nimic din aplicație, reîncerc")
            return False

        # 1) facturi sterse din Drive -> le scoatem si din aplicatie
        removed = [fid for fid, f in known.items()
                   if fid not in in_drive and f.get("status") != "processing"]
        # 2) facturi redenumite (acelasi continut) -> doar actualizam numele
        renamed = [f for f in files if f["id"] in known
                   and known[f["id"]].get("name") != f["name"]
                   and known[f["id"]].get("sig") == file_signature(f)]
        # 3) facturi noi / cu continut schimbat / cu eroare (doar la cerere)
        todo = [f for f in files
                if f["id"] not in known
                or known[f["id"]].get("sig") != file_signature(f)
                or (retry_errors and known[f["id"]].get("status") == "error")]
        # un duplicat al unei facturi sterse redevine factura valida
        if removed:
            todo_ids = {f["id"] for f in todo}
            todo += [f for f in files if f["id"] not in todo_ids
                     and known.get(f["id"], {}).get("status") == "duplicate"]

        with state_lock:
            for fid in removed:
                state["files"].pop(fid, None)
            if removed:
                gone = set(removed)
                state["interventii"] = [iv for iv in state["interventii"]
                                        if iv.get("_file_id") not in gone]
            for f in files:   # folderul din Drive (pentru grupurile de clienti); se actualizeaza la mutare
                e = state["files"].get(f["id"])
                if e is not None and e.get("folder") != f.get("folder"):
                    e["folder"] = f.get("folder")
                    cache["dirty"] = True
            for f in renamed:
                state["files"][f["id"]]["name"] = f["name"]
                for iv in state["files"][f["id"]].get("interventii") or []:
                    iv["_source"] = f["name"]
            for f in todo:
                prev = state["files"].get(f["id"], {})
                state["files"][f["id"]] = {
                    **prev, "id": f["id"], "name": f["name"], "status": "new", "folder": f.get("folder"),
                    "interventii": prev.get("interventii", []), "error": None,
                    "avertismente": prev.get("avertismente", []),
                }
            state["last_scan"] = datetime.now().isoformat()

        for fid in removed:
            add_log(f"− Stearsa din Drive: {known[fid].get('name')}")
        for f in renamed:
            add_log(f"~ Redenumita: {known[f['id']].get('name')} → {f['name']}")
        if removed or renamed:
            cache["dirty"] = True

        unchanged = len(files) - len(todo)
        add_log(f"{len(files)} PDF-uri in Drive · {unchanged} deja citite · {len(todo)} de procesat")
        if todo:
            process_queue_background(todo)
        if todo or removed:
            recompute_all()  # avertismentele "acelasi numar" devin simetrice intre facturi
        try:
            sync_geo()
        except Exception as e:  # distantele nu au voie sa strice scanarea facturilor
            add_log(f"⚠ Distante: {e}")
        if settings.get("calendar_id"):
            read_calendar()
        save_cache()
        return True
    except Exception as e:
        add_log(f"Eroare scanare: {e}")
        _drive_local.service = None  # la urmatoarea scanare ne reconectam de la zero
        save_cache()
        return False
    finally:
        with state_lock:
            state["scan"]["active"] = False
        scan_lock.release()


def monitor_loop():
    import gc
    while True:
        time.sleep(POLL_INTERVAL)
        mem = rss_mb()
        if mem > 300:
            gc.collect()
            add_log(f"⚠ Memorie folosită: {mem} MB din 512 (după curățare: {rss_mb()} MB)")
        with state_lock:
            monitoring = state["monitoring"]
        if monitoring:
            scan_folder()


_bg_started = False


def start_background():
    """Porneste monitorizarea si scanarea initiala — si sub gunicorn, nu doar la `python app.py`."""
    global _bg_started
    if _bg_started:
        return
    _bg_started = True
    threading.Thread(target=monitor_loop, daemon=True).start()
    add_log("Aplicatie pornita (parser direct, AI doar ca rezerva).")
    if FOLDER_ID:
        # asteptam ca serverul sa termine de pornit inainte de prima conexiune la Google
        t = threading.Timer(5.0, boot_scan)
        t.daemon = True
        t.start()


BOOT_RETRY = [30, 90, 180, 300, 600]   # secunde intre reincercari, apoi la 15 minute


def boot_scan(attempt=0):
    """
    Scanarea de la pornire (dupa somnul Render sau un deploy). Daca Google nu raspunde,
    reincearca singura — inainte ramaneai cu 0 peste tot pana dadeai Scanare manuala.
    """
    ok = scan_folder(retry_errors=False)   # facturile cu eroare se reincearca doar la Scanare manuala
    if (ok or (ok is None and cache["loaded"])) and cache["enabled"]:
        return      # facturile si backup-ul (regulile) sunt incarcate
    delay = BOOT_RETRY[attempt] if attempt < len(BOOT_RETRY) else 900
    add_log(f"Datele nu s-au putut încărca — reîncerc automat în {delay // 60 or delay} "
            f"{'min' if delay >= 60 else 's'} (încercarea {attempt + 2})")
    with state_lock:
        state["scan"]["retry_at"] = time.time() + delay
    t = threading.Timer(delay, boot_scan, kwargs={"attempt": attempt + 1})
    t.daemon = True
    t.start()


# ---------------------------------------------------------------- agregari
def aggregate_props(interventii):
    """
    Clasament pe proprietati. Variantele de nume ("Villa Blue", "Blue") sunt grupate automat
    (prop_key) si prin unirile manuale din setari.
    """
    props = {}
    for iv in interventii:
        raw = re.sub(r"\s+", " ", (iv.get("proprietate") or "Necunoscut").strip())
        key = prop_key(raw)
        p = props.setdefault(key, {"key": key, "interventii": 0, "forfait": 0, "ore": 0.0,
                                   "total": 0.0, "venit_ora": 0.0, "clients": set(), "ultima": "",
                                   "variante": {}, "tipuri": {}, "venit_tip": {}, "ore_tip": {}})
        p["interventii"] += 1
        p["variante"][raw] = p["variante"].get(raw, 0) + 1
        tip = iv.get("tip_menaj") or canonical_type(iv.get("tip_serviciu"))
        p["tipuri"][tip] = p["tipuri"].get(tip, 0) + 1
        suma = float(iv.get("suma") or 0)
        p["venit_tip"][tip] = p["venit_tip"].get(tip, 0.0) + suma
        if iv.get("tip_facturare") == "heure":
            p["ore_tip"][tip] = p["ore_tip"].get(tip, 0.0) + float(iv.get("ore") or iv.get("qty") or 0)
        if iv.get("tip_facturare") == "heure":
            p["ore"] += float(iv.get("ore") or iv.get("qty") or 0)
            p["venit_ora"] += suma
        else:
            p["forfait"] += 1
        p["total"] += suma
        if iv.get("client"):
            p["clients"].add(iv["client"])
        if (iv.get("data_iso") or "") > p["ultima"]:
            p["ultima"] = iv["data_iso"]
    out = list(props.values())
    max_c = max((p["interventii"] for p in out), default=1) or 1
    for p in out:
        # numele afisat: o varianta a proprietatii in care s-a unit (daca e cazul),
        # apoi cea mai des intalnita, apoi cea mai lunga
        p["name"] = max(p["variante"], key=lambda v: (base_prop_key(v) == p["key"], p["variante"][v], len(v)))
        p["variante"] = sorted(p["variante"])
        if len(p["variante"]) > 1:
            # cum a ajuns fiecare nume aici: unire manuala (se anuleaza) sau automata (se separa)
            info = []
            for v in p["variante"]:
                bk = base_prop_key(v)
                if bk != p["key"]:
                    info.append([v, "manual", bk])
                elif norm_text(v) != p["key"]:
                    info.append([v, "auto", ""])
            p["variante_info"] = info
        p["client"] = ", ".join(sorted(p.pop("clients")))
        p["ore"] = round(p["ore"], 2)
        p["total"] = round(p["total"], 2)
        p["venit_tip"] = {k: round(v, 2) for k, v in p["venit_tip"].items()}
        p["ore_tip"] = {k: round(v, 2) for k, v in p["ore_tip"].items()}
        p["medie"] = round(p["total"] / p["interventii"], 2) if p["interventii"] else 0
        p["pret_ora"] = round(p.pop("venit_ora") / p["ore"], 2) if p["ore"] else None
        p["pct"] = round(p["interventii"] / max_c * 100)
        p["perf"] = ("Performant" if p["interventii"] >= max_c * 0.7
                     else "Mediu" if p["interventii"] >= max_c * 0.4 else "Rar")
    return sorted(out, key=lambda p: (p["interventii"], p["total"]), reverse=True)


def tip_label(t):
    return TIP_LABEL.get(t, (t or "").capitalize())


SORT_KEYS = {
    "interventii": lambda p: (p["interventii"], p["total"]),
    "total": lambda p: (p["total"], p["interventii"]),
    "ore": lambda p: (p["ore"], p["total"]),
    "medie": lambda p: (p["medie"], p["total"]),
}


# ---------------------------------------------------------------- distante (etapa 2)
def sheet_id_from(text):
    """Accepta ID-ul sau linkul complet al Sheet-ului."""
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", text or "")
    return m.group(1) if m else (text or "").strip()


def read_coords_sheet(sheet_id):
    """CSV-ul Sheet-ului: intai prin service account; daca nu are acces, prin linkul public (daca exista)."""
    try:
        data = get_drive_service().files().export_media(fileId=sheet_id, mimeType="text/csv").execute(num_retries=DRIVE_RETRIES)
        return data.decode("utf-8-sig"), "service account"
    except Exception as e:  # noqa: BLE001
        first_err = e
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30) as r:
            body = r.read().decode("utf-8-sig")
        if body.lstrip().startswith("<"):
            raise ValueError("Sheet-ul nu e accesibil")
        return body, "link public"
    except Exception:
        raise RuntimeError(f"Nu pot citi Sheet-ul cu coordonate. Partajati-l (Share, Viewer) cu "
                           f"service account-ul. Detalii: {first_err}")


def sync_geo(force_routes=False):
    """Citeste coordonatele si calculeaza traseele lipsa. Pastreaza DOAR nume/lat/lng."""
    sid = sheet_id_from(settings.get("coord_sheet"))
    base = geo.parse_latlng(settings.get("baza"))
    if not sid or not base:
        return
    set_stage("citesc coordonatele caselor")
    text, how = read_coords_sheet(sid)
    coords, err = geo.parse_coords_csv(text)
    del text  # CSV-ul complet (cu eventuale coduri) nu e pastrat nicaieri
    if err:
        raise RuntimeError(err)
    # orele din Sheet (coloana "ore") sunt sursa de adevar, mai putin cele modificate acum in aplicatie
    pending = set(settings.get("ore_nescrise") or [])
    ore = settings.setdefault("ore_estimate", {})
    for c in coords:
        k = prop_key(c["nume"])
        if c.get("ore") and k not in pending:
            ore[k] = c["ore"]
    try:
        tab = read_hours_tab(sid)
        if tab is not None:
            vals, names = tab
            for nume, h in vals.items():   # foaia separata "Ore" are prioritate
                k = prop_key(nume)
                if k not in pending:
                    ore[k] = h
            # case cu ore in aplicatie, dar fara rand in foaia "Ore" (ex. randuri sterse): le rescriem
            in_tab = {prop_key(n) for n in names}
            lipsa = [k for k, h in ore.items() if h and k not in in_tab]
            if lipsa:
                settings["ore_nescrise"] = sorted(set(settings.get("ore_nescrise") or []) | set(lipsa))
                pending |= set(lipsa)
                add_log(f"Foaia „{ORE_TAB}”: {len(lipsa)} case lipsă — le rescriu din aplicație")
    except Exception as e:  # noqa: BLE001
        add_log(f"⚠ Foaia „{ORE_TAB}”: {e}")
    for c in coords:
        c.pop("ore", None)          # in "coords" pastram doar nume/lat/lng
    settings["coords"] = coords
    if pending:
        _save_async()
    settings["geo_sync"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    base_k = geo.coord_key(*base)
    if settings.get("distante_baza") != base_k or force_routes:
        settings["distante"], settings["distante_baza"] = {}, base_k
    dist = settings["distante"]
    todo = []
    for c in coords:
        k = geo.coord_key(c["lat"], c["lng"])
        if k not in dist or dist[k][2] == "estimat":   # estimarile se reincearca pe drum
            todo.append((k, (c["lat"], c["lng"])))
    todo = list(dict(todo).items())
    if todo:
        set_stage(f"calculez traseele pe drum pentru {len(todo)} case")
        try:
            routes = geo.osrm_routes(base, [p for _, p in todo])
            for (k, p), rt in zip(todo, routes):
                dist[k] = [*rt, "drum"] if rt else [*geo.estimate_route(base, p), "estimat"]
            add_log(f"Distante: {len(todo)} trasee noi calculate pe drum")
        except Exception as e:  # noqa: BLE001
            for k, p in todo:
                dist[k] = [*geo.estimate_route(base, p), "estimat"]
            add_log(f"⚠ Serviciul de rute nu a raspuns ({e}) — distante estimate din linie dreapta x{geo.ROAD_FACTOR}")
    add_log(f"Coordonate citite ({how}): {len(coords)} case")
    try:
        sync_dist_matrix(base, coords)
    except Exception as e:  # noqa: BLE001 - motorina estimata nu are voie sa strice restul
        add_log(f"⚠ Matricea de distante: {e}")
    cache["dirty"] = True


def sync_dist_matrix(base, coords):
    """Distantele pe drum intre toate casele (o singura cerere); reface doar daca s-au schimbat punctele."""
    pts = [geo.coord_key(*base)] + sorted({geo.coord_key(c["lat"], c["lng"]) for c in coords})
    cur = settings.get("dist_matrix") or {}
    if cur.get("pts") == pts and cur.get("sursa") == "drum":
        return
    points = [tuple(map(float, k.split(","))) for k in pts]
    set_stage(f"calculez distantele intre {len(pts)} puncte (motorina estimata)")
    try:
        km = geo.osrm_matrix(points)
        sursa = "drum"
        add_log(f"Distante intre case: matrice {len(pts)}×{len(pts)} calculata pe drum")
    except Exception as e:  # noqa: BLE001
        km = [[round(geo.haversine_km(a, b) * geo.ROAD_FACTOR, 1) for b in points] for a in points]
        sursa = "estimat"
        add_log(f"⚠ Distante intre case estimate din linie dreapta x{geo.ROAD_FACTOR} ({e})")
    settings["dist_matrix"] = {"pts": pts, "km": km, "sursa": sursa}
    settings["geo_v"] = settings.get("geo_v", 0) + 1


def geo_for_props(props, all_names=None):
    """Ataseaza distanta/zona fiecarei proprietati; intoarce si lista caselor fara coordonate."""
    sheet = {}
    for c in settings.get("coords") or []:
        sheet.setdefault(prop_key(c["nume"]), c)
    by_name = {c["nume"]: c for c in settings.get("coords") or []}
    dist, manual = settings.get("distante") or {}, settings.get("coord_alias") or {}
    missing = []
    for p in props:
        c = by_name.get(manual.get(p["key"])) if p["key"] in manual else sheet.get(p["key"])
        p["in_sheet"] = bool(c)
        d = dist.get(geo.coord_key(c["lat"], c["lng"])) if c else None
        if d:
            p.update({"distanta_km": d[0], "timp_min": d[1], "dist_sursa": d[2],
                      "zona": geo.zone_of(d[0]), "coord_nume": c["nume"]})
        else:
            p.update({"distanta_km": None, "timp_min": None, "zona": None})
            if not c:
                missing.append(p)
    return missing


def zones_table(props):
    rows = {}
    order = [z[1] for z in geo.ZONES] + ["fără coordonate"]
    total = sum(p["total"] for p in props) or 0
    for p in props:
        z = p.get("zona") or "fără coordonate"
        r = rows.setdefault(z, {"zona": z, "case": 0, "interventii": 0, "venit": 0.0, "km": 0.0, "min": 0.0})
        r["case"] += 1
        r["interventii"] += p["interventii"]
        r["venit"] += p["total"]
        if p.get("distanta_km") is not None:
            r["km"] += 2 * p["distanta_km"] * p["interventii"]     # dus-intors la fiecare menaj
            r["min"] += 2 * (p.get("timp_min") or 0) * p["interventii"]
    out = []
    for z in order:
        if z in rows:
            r = rows[z]
            out.append({**r, "venit": round(r["venit"], 2), "km": round(r["km"]),
                        "ore_drum": round(r["min"] / 60, 1),
                        "procent": round(r["venit"] / total * 100, 1) if total else 0,
                        "eur_km": round(r["venit"] / r["km"], 2) if r["km"] else None,
                        "medie": round(r["venit"] / r["interventii"], 2) if r["interventii"] else 0})
    return out


def costly_houses(props, min_km=15, limit=12):
    """Case departe de baza care aduc putin raportat la drum: € pe km parcurs (dus-intors)."""
    rows = []
    for p in props:
        km = p.get("distanta_km")
        if km and km >= min_km and p["interventii"]:
            rows.append({"name": p["name"], "client": p["client"], "km": km, "timp_min": p.get("timp_min"),
                         "medie": p["medie"], "interventii": p["interventii"],
                         "eur_km": round(p["medie"] / (2 * km), 2)})
    rows.sort(key=lambda r: r["eur_km"])
    return rows[:limit]


# ---------------------------------------------------------------- indicatori (etapa 1)
_an_cache = {}   # (data_key, perioada, zi) -> (lista de interventii, rezultat); cateva intrari
_an_lock = threading.Lock()


def data_key():
    """Se schimba de fiecare data cand se schimba interventiile sau unirile de proprietati."""
    ivs = state["interventii"]
    alias = hash(tuple(sorted(settings["alias"].items())))
    ok = hash((tuple(sorted(settings.get("preturi_ok") or [])), tuple(sorted(settings.get("excluse") or [])),
               tuple(settings.get("cuvinte_produse") or []), tuple(sorted(settings.get("pastrate") or [])),
               json.dumps(settings.get("grupuri_clienti") or [], sort_keys=True), settings.get("folder_v"),
               settings.get("cost_v"), settings.get("consum"), settings.get("pret_motorina"), settings.get("ore_v"),
               json.dumps(settings.get("ore_estimate") or {}, sort_keys=True), tuple(settings.get("separate") or [])))
    g = hash((settings.get("baza"), len(settings.get("coords") or []), len(settings.get("distante") or {}),
              tuple(sorted((settings.get("coord_alias") or {}).items())), settings.get("geo_sync"), settings.get("geo_v")))
    return f"{id(ivs)}-{len(ivs)}-{alias}-{ok}-{g}"


# ---------------------------------------------------------------- cheltuieli & profit
_fuel_cache = {}


def house_points():
    """cheie proprietate -> (lat, lng), dupa aceleasi reguli ca distantele de la baza."""
    coords = settings.get("coords") or []
    sheet = {}
    for c in coords:
        sheet.setdefault(prop_key(c["nume"]), (c["lat"], c["lng"]))
    by_name = {c["nume"]: (c["lat"], c["lng"]) for c in coords}
    out = dict(sheet)
    for k, nume in (settings.get("coord_alias") or {}).items():
        if nume in by_name:
            out[k] = by_name[nume]
    return out


def road_dist():
    m = settings.get("dist_matrix") or {}
    idx = {k: i for i, k in enumerate(m.get("pts") or [])}
    km = m.get("km") or []
    memo = {}

    def dist(a, b):
        if a == b:
            return 0.0
        ka, kb = geo.coord_key(*a), geo.coord_key(*b)
        i, j = idx.get(ka), idx.get(kb)
        if i is not None and j is not None and km[i][j] is not None:
            return km[i][j]
        if (ka, kb) not in memo:
            memo[(ka, kb)] = geo.haversine_km(a, b) * geo.ROAD_FACTOR
        return memo[(ka, kb)]
    return dist


def fuel_routes(ivs):
    """Km pe zi si pe casa (traseul zilnic baza -> case -> baza), pastrat in memorie pe set de date."""
    base = geo.parse_latlng(settings.get("baza"))
    if not base or not settings.get("coords"):
        return {}, {}, 0
    key = (id(ivs), len(ivs), settings.get("geo_v"), json.dumps(settings.get("coord_alias") or {}, sort_keys=True),
           tuple(sorted(settings.get("alias", {}).items())), tuple(settings.get("separate") or []))
    hit = _fuel_cache.get("k")
    if hit == key:
        return _fuel_cache["v"]
    hp = house_points()

    def house_point(iv):
        k = prop_key(re.sub(r"\s+", " ", (iv.get("proprietate") or "").strip()) or "Necunoscut")
        return (k, hp[k]) if k in hp else None
    val = costs.daily_routes(ivs, house_point, base, road_dist())
    _fuel_cache.update({"k": key, "v": val})
    return val


def fuel_per_km():
    return (float(settings.get("consum") or 0) / 100) * float(settings.get("pret_motorina") or 0)


def cost_months(period, ivs_all):
    today_m = costs.today_month()
    if period and period.get("start"):
        first, last = period["start"][:7], min(period["end"][:7], today_m)
    else:
        ms = [iv["data_iso"][:7] for iv in ivs_all if iv.get("data_iso")]
        ms += [e["data"][:7] for e in settings.get("cheltuieli") or [] if e.get("data")]
        if not ms:
            return []
        first, last = min(ms), today_m
    return analytics.months_range(first, last) if first <= last else []


def _first_cost_month(rows):
    entries = settings.get("cheltuieli") or []
    if not entries:
        return None
    first = min(e["data"][:7] for e in entries if e.get("data"))
    with_rev = [r["luna"] for r in rows if r["venit"] > 0]
    return first if with_rev and with_rev[0] < first else None


def profit_block(ivs_period, ivs_all, period, props):
    """Venit, cheltuieli, profit pe perioada + pe luni + profit estimat pe casa (adaugat in props)."""
    months = cost_months(period, ivs_all)
    cats = settings.get("categorii_cheltuieli") or costs.CATEGORII_IMPLICITE
    occ = [e for e in costs.expand(settings.get("cheltuieli") or [], costs.today_month()) if e["luna"] in set(months)]
    venit_luna = {}
    for iv in ivs_period:
        if iv.get("data_iso"):
            venit_luna[iv["data_iso"][:7]] = venit_luna.get(iv["data_iso"][:7], 0.0) + float(iv.get("suma") or 0)
    km_day, km_house_all, fara = fuel_routes(ivs_all)
    fkm = fuel_per_km()
    s, e = (period or {}).get("start"), (period or {}).get("end")
    fuel_est_luna, km_total = {}, 0.0
    for d, km in km_day.items():
        if s and not (s <= d <= e):
            continue
        fuel_est_luna[d[:7]] = fuel_est_luna.get(d[:7], 0.0) + km * fkm
        km_total += km
    rows = costs.month_summary(months, venit_luna, occ, cats, fuel_est_luna)
    # km pe casa doar in perioada aleasa
    if s and km_day:
        hp_map = house_points()

        def house_point(iv):
            k = prop_key(re.sub(r"\s+", " ", (iv.get("proprietate") or "").strip()) or "Necunoscut")
            return (k, hp_map[k]) if k in hp_map else None
        _, km_house, _ = costs.daily_routes(ivs_period, house_point, geo.parse_latlng(settings.get("baza")), road_dist())
    else:
        km_house = km_house_all
    tot = lambda k: round(sum(r[k] for r in rows), 2)  # noqa: E731
    fuel_real, fuel_est = tot("motorina_reala"), round(sum(fuel_est_luna.values()), 2)
    real_months = [r for r in rows if r["motorina_din"] == "bonuri"]
    est_on_real = sum(r["motorina_estimata"] for r in real_months)
    scale = (sum(r["motorina_reala"] for r in real_months) / est_on_real) if est_on_real > 0 else 1.0
    pm = gen = 0.0
    pe_cat = {}
    for o in occ:
        sm = float(o.get("suma") or 0)
        pe_cat[o["categorie"]] = pe_cat.get(o["categorie"], 0.0) + sm
        if costs.cat_type(o["categorie"], cats) == "personal" or costs.is_fuel(o["categorie"]):
            continue
        if costs.per_menaj(o["categorie"]):
            pm += sm
        else:
            gen += sm
    hp = costs.house_profit(props, km_house, fkm, {"per_menaj": pm, "general": gen, "motorina_scalare": scale})
    known = house_points() if settings.get("coords") else {}
    for p in props:
        p.update(hp.get(p["key"], {}))
        if p["key"] not in known:
            # fara coordonate nu stim drumul: nu afisam 0 € (ar parea gratis), ci „necunoscut”
            p["motorina_est"] = p["km_drum"] = None
    venit, chelt = tot("venit"), tot("cheltuieli")
    personal = tot("personal")
    tip = {c["nume"]: c.get("tip", "firma") for c in cats}
    cat_rows = sorted(({"categorie": k, "suma": round(v, 2), "tip": tip.get(k, "firma")} for k, v in pe_cat.items()),
                      key=lambda r: -r["suma"])
    if fuel_real <= 0 and fuel_est > 0:
        cat_rows.append({"categorie": "Motorină (estimată)", "suma": fuel_est, "tip": "estimat"})
        cat_rows.sort(key=lambda r: -r["suma"])
    return {
        "luni": rows, "venit": venit, "cheltuieli": chelt, "profit": round(venit - chelt, 2),
        "marja": round((venit - chelt) / venit * 100, 1) if venit else None,
        "personal": personal, "ramas": round(venit - chelt - personal, 2),
        "motorina": {"reala": fuel_real, "estimata": fuel_est, "km": round(km_total), "zile_fara_coordonate": fara,
                     "consum": settings.get("consum"), "pret": settings.get("pret_motorina"),
                     "sursa_distante": (settings.get("dist_matrix") or {}).get("sursa")},
        "categorii": cat_rows,
        "are_cheltuieli": bool(settings.get("cheltuieli")),
        # perioada incepe inainte de prima cheltuiala trecuta -> profitul acelor luni e supraestimat
        "acoperire": _first_cost_month(rows),
    }


def monthly_profit(months, ivs_all):
    cats = settings.get("categorii_cheltuieli") or costs.CATEGORII_IMPLICITE
    ms = set(months)
    occ = [e for e in costs.expand(settings.get("cheltuieli") or [], costs.today_month()) if e["luna"] in ms]
    venit_luna = {}
    for iv in ivs_all:
        m = (iv.get("data_iso") or "")[:7]
        if m in ms:
            venit_luna[m] = venit_luna.get(m, 0.0) + float(iv.get("suma") or 0)
    km_day, _, _ = fuel_routes(ivs_all)
    fkm, fuel = fuel_per_km(), {}
    for d, km in km_day.items():
        if d[:7] in ms:
            fuel[d[:7]] = fuel.get(d[:7], 0.0) + km * fkm
    return costs.month_summary(months, venit_luna, occ, cats, fuel)


def with_monthly_profit(rows, ivs_all):
    by = {r["luna"]: r for r in monthly_profit([r["luna"] for r in rows], ivs_all)}
    for r in rows:
        m = by.get(r["luna"]) or {}
        r.update({"cheltuieli": m.get("cheltuieli", 0), "profit": m.get("profit"), "marja": m.get("marja"),
                  "motorina_din": m.get("motorina_din")})
    return rows


def add_hours(props):
    """
    Ore estimate = timpul REAL al unei persoane la un menaj complet (Standard / Ouverture / Fermeture),
    indiferent daca menajul e facturat forfait sau la ora (orele facturate sunt dupa piata, nu timpul real).
    € / ora = incasarile din menajele complete / (menaje complete x ore estimate).
    Intermédiaire-urile sunt alt tip de munca: se arata separat si nu intra in € / ora.
    """
    ore = settings.get("ore_estimate") or {}
    for p in props:
        h = ore.get(p["key"])
        tipuri, venit_tip = p.get("tipuri") or {}, p.get("venit_tip") or {}
        n_int = tipuri.get("intermediaire", 0)
        n_complet = sum(n for t, n in tipuri.items() if t != "intermediaire")
        v_int = venit_tip.get("intermediaire", 0.0)
        v_complet = round(p["total"] - v_int, 2)
        p.update({"ore_est": h, "menaje_complete": n_complet, "venit_complet": v_complet,
                  "interm_n": n_int, "interm_venit": round(v_int, 2),
                  "interm_ore": (p.get("ore_tip") or {}).get("intermediaire"),
                  "ore_lipsa": n_complet > 0 and not h})
        if n_complet and h:
            lucru = n_complet * h
            p["ore_lucru"] = round(lucru, 1)
            p["eur_ora_lucru"] = round(v_complet / lucru, 2)
        else:                       # fara ore estimate (sau doar intermediaire): nu calculam
            p["ore_lucru"] = p["eur_ora_lucru"] = None


def compute_analytics(period_key="all", client=None):
    """Indicatorii perioadei; cu `client`, clasamentul caselor e doar pentru clientul acela."""
    base = _compute_analytics_base(period_key)
    client = (client or "").strip()
    if not client:
        return base
    key = (data_key(), period_key, date.today().isoformat(), "client", norm_text(client))
    with _an_lock:
        hit = _an_cache.get(key)
        if hit:
            return hit[1]
    with state_lock:
        ivs_raw = state["interventii"]
    ivs_all = analysis_ivs(ivs_raw)[0]
    ck = norm_text(client)
    ivs_all_c = [iv for iv in ivs_all if norm_text(iv.get("client") or "") == ck]
    ivs_c = analytics.filter_ivs(ivs_all_c, base["period"])
    months = analytics.period_months(ivs_all, base["period"])
    props = aggregate_props(ivs_c)
    activity = analytics.property_activity(ivs_c, ivs_all_c, prop_key, len(months))
    full = {p["key"]: p for p in base["props"]}
    for p in props:
        p.update(activity.get(p["key"], {}))
        f = full.get(p["key"]) or {}
        # profitul estimat al casei, proportional cu partea clientului din incasarile casei
        if f.get("cost_est") is not None and f.get("total"):
            cost = f["cost_est"] * p["total"] / f["total"]
            p.update({"profit_est": round(p["total"] - cost, 2), "cost_est": round(cost, 2),
                      "marja_est": round((p["total"] - cost) / p["total"] * 100, 1) if p["total"] else None})
    geo_for_props(props)
    add_hours(props)
    tot = round(sum(p["total"] for p in props), 2)
    val = {**base, "props": props, "client_sel": client,
           "client_info": {"venit": tot, "interventii": sum(p["interventii"] for p in props), "case": len(props),
                           "procent": round(tot / base["metrics"]["venit"] * 100, 1) if base["metrics"]["venit"] else 0,
                           "inactive": sum(1 for p in props if p.get("inactiva"))}}
    with _an_lock:
        while len(_an_cache) >= 12:
            del _an_cache[next(iter(_an_cache))]
        _an_cache[key] = (ivs_raw, val)
    return val


def _compute_analytics_base(period_key="all"):
    """
    Toti indicatorii pentru perioada aleasa. Rezultatul e pastrat in memorie pana se schimba
    datele/perioada/ziua — dashboard-ul il cere des, dar calculul se face doar cand e nevoie.
    """
    with state_lock:
        ivs_raw = state["interventii"]
    key = (data_key(), period_key, date.today().isoformat())
    with _an_lock:
        hit = _an_cache.get(key)
        if hit:
            return hit[1]
    ivs_all, ex_all = analysis_ivs(ivs_raw)
    period, options = analytics.resolve_period(period_key, ivs_all)
    ivs = analytics.filter_ivs(ivs_all, period)
    ex = analytics.filter_ivs(ex_all, period)
    excl_set = set(settings.get("excluse") or [])
    sk = sheet_keys()
    ex_rows = {}
    for iv in ex:
        raw = re.sub(r"\s+", " ", (iv.get("proprietate") or "Necunoscut").strip())
        k = prop_key(raw)
        r = ex_rows.setdefault(k, {"key": k, "name": raw, "randuri": 0, "total": 0.0, "ultima": "",
                                   "manual": k in excl_set, "clienti": set(), "in_sheet": k in sk})
        r["randuri"] += 1
        r["ultima"] = max(r["ultima"], iv.get("data_iso") or "")
        r["total"] += float(iv.get("suma") or 0)
        if iv.get("client"):
            r["clienti"].add(iv["client"])
    excluse = sorted(ex_rows.values(), key=lambda r: -r["total"])
    for r in excluse:
        r["total"] = round(r["total"], 2)
        r["clienti"] = ", ".join(sorted(r["clienti"]))
    # proprietatile scoase manual, dar fara randuri in perioada aleasa (ca sa poata fi puse inapoi)
    for k in sorted(excl_set - set(ex_rows)):
        excluse.append({"key": k, "name": k, "randuri": 0, "total": 0, "manual": True, "clienti": "",
                        "ultima": "", "in_sheet": k in sk})

    names_all = {p["key"]: p["name"] for p in aggregate_props(ivs_all)}

    def prop_display(name):
        return names_all.get(prop_key(name), name)

    months = analytics.period_months(ivs_all, period)
    activity = analytics.property_activity(ivs, ivs_all, prop_key, len(months))
    props = aggregate_props(ivs)
    for p in props:
        p.update(activity.get(p["key"], {}))
    missing = geo_for_props(props)
    profit = profit_block(ivs, ivs_all, period, props)
    add_hours(props)
    sheet_names = sorted({c["nume"] for c in settings.get("coords") or []}, key=str.casefold)

    clients, top3 = analytics.clients_table(ivs, norm_text)
    # o singura luna aleasa -> graficul arata ultimele 12 luni pana la ea (o bara singura nu spune nimic)
    if period.get("key", "").startswith("m") and months:
        all_months = analytics.period_months(ivs_all, None)
        end = all_months.index(months[-1]) + 1 if months[-1] in all_months else len(all_months)
        chart_months = all_months[max(0, end - 12):end]
    else:
        chart_months = months
    prices = analytics.price_checks(ivs_all, period, prop_key, prop_display)
    ok_set = set(settings.get("preturi_ok") or [])
    for r in prices:
        r["ok"] = r["id"] in ok_set

    venit = round(sum(float(iv.get("suma") or 0) for iv in ivs), 2)
    val = {
        "period": period, "options": options,
        "metrics": {
            "venit": venit, "interventii": len(ivs), "proprietati": len(props),
            "ore": round(sum(float(iv.get("ore") or 0) for iv in ivs if iv.get("tip_facturare") == "heure"), 2),
            "medie": round(venit / len(ivs), 2) if ivs else 0,
            "luni": len(months),
            "abateri": sum(1 for r in prices if r["fel"] == "abatere" and not r["ok"]),
            "inactive": sum(1 for p in props if p.get("inactiva")),
            "venit_inactive": round(sum(p["total"] for p in props if p.get("inactiva")), 2),
            "fara_data": sum(1 for iv in ivs if not iv.get("data_iso")),
            "venit_excluse": round(sum(float(iv.get("suma") or 0) for iv in ex), 2),
            "cheltuieli": profit["cheltuieli"], "profit": profit["profit"], "marja": profit["marja"],
        },
        "profit": profit,
        "ore_status": dict(hours_status), "ore_sheet": bool(settings.get("coord_sheet")),
        "formule": settings.get("formule") or [],
        "coloane": settings.get("coloane") or {},
        "excluse": excluse,
        "props": props,
        "clients": clients, "top3": top3,
        "monthly": with_monthly_profit(analytics.monthly_series(ivs_all, chart_months, prop_key), ivs_all),
        "luna_selectata": months[-1] if period.get("key", "").startswith("m") and months else None,
        "calendar": analytics.calendar_data(ivs, prop_display),
        "prices": prices,
        "zones": zones_table(props),
        "costly": costly_houses(props),
        "geo": {
            "configurat": bool(settings.get("coord_sheet")),
            "baza": settings.get("baza"),
            "case_cu_distanta": sum(1 for p in props if p.get("distanta_km") is not None),
            "estimate": sum(1 for p in props if p.get("dist_sursa") == "estimat"),
            "fara_coordonate": [{"key": p["key"], "name": p["name"], "interventii": p["interventii"],
                                 "total": p["total"]} for p in sorted(missing, key=lambda x: -x["total"])],
            "nume_sheet": sheet_names,
            "ultima_citire": settings.get("geo_sync"),
        },
    }
    with _an_lock:
        # pastram referinta la lista: cat timp e in cache, id-ul ei nu poate fi refolosit
        for k in [k for k in _an_cache if k[0] != key[0]]:
            del _an_cache[k]            # date vechi -> afara
        while len(_an_cache) >= 8:
            del _an_cache[next(iter(_an_cache))]
        _an_cache[key] = (ivs_raw, val)
    return val


def build_excel(sort="interventii", period_key="all", client=None):
    with state_lock:
        files = [dict(f) for f in state["files"].values()]
    an = compute_analytics(period_key, client)
    interventii = analytics.filter_ivs(analysis_ivs(state["interventii"])[0], an["period"])
    props = sorted(an["props"], key=SORT_KEYS.get(sort, SORT_KEYS["interventii"]), reverse=True)

    wb = openpyxl.Workbook()
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    side = Side(style="thin", color="D1D5DB")
    thin = Border(left=side, right=side, top=side, bottom=side)

    def sheet(ws, headers, widths):
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            c = ws.cell(1, col)
            c.fill, c.font, c.border = header_fill, header_font, thin
            c.alignment = Alignment(horizontal="center")
        for i, w in enumerate(widths):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
        ws.freeze_panes = "A2"

    ws1 = wb.active
    ws1.title = "Clasament proprietati"
    sheet(ws1, ["Rang", "Proprietate", "Client", "Interventii", "Standard", "Intermédiaire",
                "Ouverture", "Ore (la ora)", "€/ora", "Venit (EUR)", "Medie / interventie",
                "Interventii / luna", "Ultima interventie", "Zile de la ultima", "Inactiva",
                "Distanta (km)", "Timp drum (min)", "Zona",
                "Performanta", "Apare in facturi ca", "Profit estimat (EUR)", "Marja estimata %",
                "Ore / menaj (1 pers.)", "Ore de lucru (menaje complete)", "EUR / ora (menaje complete)",
                "Venit Intermediaire", "Ore Intermediaire"],
          [6, 26, 24, 11, 10, 13, 11, 11, 8, 13, 17, 15, 16, 15, 9, 12, 14, 12, 13, 36, 16, 14, 14, 16, 16, 14, 13])
    fcols = []
    for fdef in settings.get("formule") or []:
        try:
            fcols.append((fdef["nume"], formulas.compile_expr(fdef["expr"])))
        except formulas.FormulaError:
            pass
    for j, (fname, _) in enumerate(fcols):
        c = ws1.cell(1, 28 + j, fname)
        c.font, c.fill = header_font, header_fill
        ws1.column_dimensions[openpyxl.utils.get_column_letter(28 + j)].width = 14
    for i, p in enumerate(props, 1):
        ultima = datetime.fromisoformat(p["ultima"]).strftime("%d/%m/%Y") if p.get("ultima") else ""
        t = p["tipuri"]
        ws1.append([i, p["name"], p["client"], p["interventii"], t.get("standard") or None,
                    t.get("intermediaire") or None, t.get("ouverture") or None, p["ore"] or None,
                    p["pret_ora"], p["total"], p["medie"], p.get("frecventa_luna"), ultima,
                    p.get("zile_de_la_ultima"), "da" if p.get("inactiva") else "",
                    p.get("distanta_km"), p.get("timp_min"), p.get("zona") or "", p["perf"],
                    ", ".join(p["variante"]) if len(p["variante"]) > 1 else "",
                    p.get("profit_est"), p.get("marja_est"), p.get("ore_est"), p.get("ore_lucru"), p.get("eur_ora_lucru"),
                    p.get("interm_venit") or None, p.get("interm_ore")])
        for j, (_, code) in enumerate(fcols):     # indicatorii proprii, dupa coloanele fixe
            ws1.cell(ws1.max_row, 28 + j, formulas.evaluate(code, formulas.house_values(p)))
        fill = {"Performant": "D1FAE5", "Rar": "FEE2E2"}.get(p["perf"])
        for col in range(1, 28):
            c = ws1.cell(i + 1, col)
            c.border = thin
            if fill:
                c.fill = PatternFill("solid", fgColor=fill)

    ws2 = wb.create_sheet("Detalii interventii")
    sheet(ws2, ["Proprietate", "Proprietate (grupata)", "Client", "Data", "Tip menaj", "Facturare",
                "De ce la ora", "Ore", "Pret unitar", "Suma (EUR)", "Factura", "Fisier"],
          [24, 22, 24, 12, 14, 10, 13, 7, 11, 11, 13, 28])
    names = {p["key"]: p["name"] for p in props}
    motiv = {"client": "client la ora", "tip": "tip menaj", "pret": "pret < prag"}
    for iv in sorted(interventii, key=lambda x: (x.get("data_iso") or "", x.get("proprietate") or "")):
        d = datetime.fromisoformat(iv["data_iso"]) if iv.get("data_iso") else None
        ws2.append([iv.get("proprietate"), names.get(prop_key(iv.get("proprietate") or "")),
                    iv.get("client"), d,
                    tip_label(iv.get("tip_menaj") or canonical_type(iv.get("tip_serviciu"))),
                    iv.get("tip_facturare"), motiv.get(iv.get("motiv_ora"), ""), iv.get("ore"),
                    iv.get("pret_unitar"), round(float(iv.get("suma") or 0), 2), iv.get("factura"),
                    iv.get("_source")])
        r = ws2.max_row
        ws2.cell(r, 4).number_format = "DD/MM/YYYY"
        for col in range(1, 13):
            ws2.cell(r, col).border = thin

    ws3 = wb.create_sheet("Facturi")
    sheet(ws3, ["Fisier", "Numar", "Client", "Data emiterii", "Total (EUR)", "Interventii",
                "Status", "Metoda", "Avertismente / eroare"],
          [32, 13, 26, 13, 12, 11, 11, 9, 70])
    for f in sorted(files, key=lambda f: f.get("name") or ""):
        ws3.append([f.get("name"), f.get("numar"), f.get("client"), f.get("data_emitere"),
                    f.get("total"), len(f.get("interventii") or []), f.get("status"), f.get("metoda"),
                    f.get("error") or " | ".join(f.get("avertismente") or [])])
        r = ws3.max_row
        for col in range(1, 10):
            ws3.cell(r, col).border = thin
        if f.get("status") in ("warning", "error", "duplicate"):
            ws3.cell(r, 7).fill = PatternFill("solid", fgColor="FEF3C7" if f["status"] != "error" else "FEE2E2")

    wsz = wb.create_sheet("Zone")
    sheet(wsz, ["Zona", "Case", "Interventii", "Venit (EUR)", "% din venit", "Km parcursi (dus-intors)",
                "Ore de drum", "€ / km", "Medie / interventie"], [16, 8, 11, 13, 12, 22, 12, 9, 18])
    for z in an.get("zones") or []:
        wsz.append([z["zona"], z["case"], z["interventii"], z["venit"], z["procent"] / 100, z["km"] or None,
                    z["ore_drum"] or None, z["eur_km"], z["medie"]])
        wsz.cell(wsz.max_row, 5).number_format = "0.0%"
        for col in range(1, 10):
            wsz.cell(wsz.max_row, col).border = thin

    ws4 = wb.create_sheet("Clienti")
    sheet(ws4, ["Client", "Venit (EUR)", "% din venit", "Case", "Interventii", "Medie / interventie",
                "Ore (la ora)", "Ultima interventie"], [32, 13, 12, 8, 11, 18, 13, 17])
    for c in an["clients"]:
        ws4.append([c["client"], c["venit"], c["procent"] / 100, c["case"], c["interventii"], c["medie"],
                    c["ore"] or None, datetime.fromisoformat(c["ultima"]) if c["ultima"] else None])
        r = ws4.max_row
        ws4.cell(r, 3).number_format = "0.0%"
        ws4.cell(r, 8).number_format = "DD/MM/YYYY"
        for col in range(1, 9):
            ws4.cell(r, col).border = thin

    ws5 = wb.create_sheet("Evolutie lunara")
    sheet(ws5, ["Luna", "Venit (EUR)", "vs luna trecuta", "vs anul trecut", "Interventii", "Ore",
                "Case active"], [18, 13, 15, 15, 11, 8, 12])
    for m in an["monthly"]:
        ws5.append([m["eticheta"], m["venit"],
                    m["vs_luna_trecuta"] / 100 if m["vs_luna_trecuta"] is not None else None,
                    m["vs_an_trecut"] / 100 if m["vs_an_trecut"] is not None else None,
                    m["interventii"], m["ore"] or None, m["case_active"]])
        r = ws5.max_row
        ws5.cell(r, 3).number_format = ws5.cell(r, 4).number_format = "+0.0%;-0.0%;0%"
        for col in range(1, 8):
            ws5.cell(r, col).border = thin

    ws6 = wb.create_sheet("Verificare preturi")
    sheet(ws6, ["Tip", "Proprietate", "Client", "Tip menaj", "Data", "Factura", "Pret facturat",
                "Pret obisnuit", "Diferenta", "Unitate", "Fisier"], [13, 24, 24, 14, 12, 13, 13, 13, 11, 8, 30])
    fel = {"abatere": "Abatere", "de confirmat": "De confirmat", "schimbare": "Schimbare tarif"}
    for p in an["prices"]:
        ws6.append([fel[p["fel"]] + (" (verificat)" if p.get("ok") else ""), p["proprietate"], p["client"], tip_label(p["tip"]),
                    datetime.fromisoformat(p["data"]) if p["data"] else None, p["factura"], p["pret"],
                    p["pret_anterior"], p["diferenta"], p["unitate"], p["fisier"]])
        r = ws6.max_row
        ws6.cell(r, 5).number_format = "DD/MM/YYYY"
        for col in range(1, 12):
            ws6.cell(r, col).border = thin
        if p["fel"] == "abatere":
            ws6.cell(r, 1).fill = PatternFill("solid", fgColor="FEE2E2")

    pr = an.get("profit") or {}
    if pr.get("luni"):
        ws7 = wb.create_sheet("Profit lunar")
        sheet(ws7, ["Luna", "Venit", "Cheltuieli firma", "Profit", "Marja %", "Motorina", "Motorina din",
                    "Transferuri personale", "Ramas"], [12, 12, 15, 12, 9, 11, 13, 18, 12])
        for r in pr["luni"]:
            fuel = r["motorina_reala"] if r["motorina_din"] == "bonuri" else r["motorina_estimata"]
            ws7.append([r["luna"], r["venit"], r["cheltuieli"], r["profit"], r["marja"], round(fuel, 2),
                        r["motorina_din"], r["personal"], r["ramas"]])
            row = ws7.max_row
            for col in range(1, 10):
                ws7.cell(row, col).border = thin
            ws7.cell(row, 4).font = Font(bold=True, color="047857" if r["profit"] >= 0 else "B91C1C")
        ws8 = wb.create_sheet("Cheltuieli")
        sheet(ws8, ["Data", "Categorie", "Suma", "Nota", "Lunar"], [12, 26, 11, 40, 8])
        for e in sorted(costs.expand(settings.get("cheltuieli") or [], costs.today_month()), key=lambda x: x["aparitie"]):
            ws8.append([datetime.fromisoformat(e["aparitie"]), e["categorie"], e["suma"], e.get("nota") or "",
                        "da" if e.get("lunar") else ""])
            ws8.cell(ws8.max_row, 1).number_format = "DD/MM/YYYY"

    # perioada exportata, vizibila in titlul primei foi
    ws1.title = (f"{an['client_sel'][:14]} {an['period']['label']}" if an.get("client_sel")
                 else f"Clasament {an['period']['label']}")[:31].replace("/", "-")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------- rute
def _public_file(f, light=False):
    """light=True: fara lista de interventii (dashboard-ul o cere la fiecare cateva secunde)."""
    heavy = ("interventii", "randuri", "antet", "av_factura")
    out = {k: v for k, v in f.items() if not k.startswith("_") and not (light and k in heavy)}
    if light:
        out["n_interventii"] = len(f.get("interventii") or [])
        out["verificat"] = is_verified(f)
    return out


_STATUS_ORDER = {"processing": 0, "done": 1, "warning": 1, "duplicate": 1, "error": 1, "new": 2}


def _file_sort_key(f):
    # in lucru sus, apoi cele citite (cele mai recente primele), apoi cele care asteapta
    return (_STATUS_ORDER.get(f.get("status"), 3), -(f.get("processed_at") or 0), f.get("name") or "")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/state")
def api_state():
    """
    Cerut de dashboard la cateva secunde. Lista de facturi se trimite doar cand s-a schimbat
    (files_v) — pe serverul gratuit (0,1 CPU) asta conteaza mult.
    """
    client_v = request.args.get("files_v")
    with state_lock:
        sig = hash((tuple((fid, f.get("status"), f.get("processed_at"), f.get("name"), len(f.get("avertismente") or []))
                          for fid, f in state["files"].items()),
                    json.dumps(settings.get("verificate") or {}, sort_keys=True), state["scan"].get("active")))
        files_v = f"{len(state['files'])}-{sig}"
        send_files = client_v != files_v
        files = sorted((_public_file(f, light=True) for f in state["files"].values()),
                       key=_file_sort_key) if send_files else None
        counts = {
            "files": len(state["files"]),
            "analyzed": sum(1 for f in state["files"].values() if f.get("status") in ("done", "warning", "duplicate")),
        }
        warn_files = [f for f in state["files"].values() if f.get("status") in ("warning", "error")]
        interventii = state["interventii"]
        s = {
            "files_v": files_v,
            "now": time.time(),
            "last_scan": state["last_scan"],
            "monitoring": state["monitoring"],
            "log": state["log"][-50:],
            "scan": dict(state["scan"]),
        }
    if send_files:
        s["files"] = files
    s["metrics"] = {**counts,
                    "warnings": sum(1 for f in warn_files if not is_verified(f)),
                    "interventii": len(interventii),
                    "revenue": round(sum(float(iv.get("suma") or 0) for iv in interventii), 2)}
    s["data_key"] = data_key()  # dashboard-ul cere /api/analytics doar cand se schimba
    s["backup"] = {"ok": bool(cache["loaded"] and cache["enabled"]), "msg": cache.get("msg") or ""}
    return jsonify(s)


@app.route("/api/analytics")
def api_analytics():
    val = compute_analytics(request.args.get("period", "all"), request.args.get("client"))
    # formulele si ordinea coloanelor se iau mereu din setarile curente (rezultatul de mai sus e tinut in memorie)
    return jsonify({**val, "formule": settings.get("formule") or [], "coloane": settings.get("coloane") or {},
                    "intrebari": settings.get("intrebari") or []})


@app.route("/api/factura/<file_id>")
def api_factura(file_id):
    with state_lock:
        f = state["files"].get(file_id)
        return (jsonify(_public_file(f)), 200) if f else (jsonify({"error": "not found"}), 404)


@app.route("/api/scan", methods=["POST"])
def api_scan():
    threading.Thread(target=scan_folder, kwargs={"retry_errors": True}, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/monitor", methods=["POST"])
def api_monitor():
    data = request.json or {}
    with state_lock:
        state["monitoring"] = bool(data.get("active", False))
        active = state["monitoring"]
    add_log("Monitorizare " + ("activata" if active else "oprita"))
    return jsonify({"monitoring": active})


@app.route("/api/export")
def api_export():
    buf = build_excel(request.args.get("sort", "interventii"), request.args.get("period", "all"),
                      request.args.get("client"))
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"analiza-curatenie-{datetime.now():%Y-%m-%d}.xlsx",
    )


# ---------------------------------------------------------------- setari & unire proprietati
_save_event = threading.Event()
_saver_started = False
_saver_lock = threading.Lock()


def _saver_loop():
    """
    UN singur fir pentru toate salvarile (backup + orele din Sheet), grupate la cateva secunde.
    Inainte fiecare modificare pornea un fir nou cu propria conexiune Google -> memoria crestea
    pana la limita de 512 MB cand se introduceau multe ore una dupa alta.
    """
    import gc
    while True:
        _save_event.wait()
        time.sleep(3)                 # grupam modificarile facute una dupa alta
        _save_event.clear()
        try:
            if settings.get("ore_nescrise") and settings.get("coord_sheet"):
                flush_hours()         # scrie orele si salveaza backup-ul
            else:
                save_cache()
        except Exception as e:  # noqa: BLE001
            add_log(f"⚠ Salvare: {e}")
        gc.collect()


def _save_async():
    global _saver_started
    with _saver_lock:
        if not _saver_started:
            _saver_started = True
            threading.Thread(target=_saver_loop, daemon=True).start()
    _save_event.set()


def merge_suggestions(props, limit=20):
    """Perechi de proprietati cu nume foarte asemanatoare (ex. 'Reppaser' / 'Reppasage')."""
    import difflib
    ignored = set(settings.get("ignorate") or [])
    items = [(p["key"], p["key"].replace(" ", ""), p) for p in props]
    out = []
    for i in range(len(items)):
        k1, a, p1 = items[i]
        for j in range(i + 1, len(items)):
            k2, b, p2 = items[j]
            if len(a) < 4 or len(b) < 4 or "|".join(sorted((k1, k2))) in ignored:
                continue
            if re.sub(r"\d", "", a) == re.sub(r"\d", "", b):
                continue  # difera doar prin cifre: 'UGF 56' / 'UGF 57' = case diferite
            sm = difflib.SequenceMatcher(None, a, b)
            if sm.real_quick_ratio() < 0.8 or sm.quick_ratio() < 0.8:
                continue  # filtre rapide, ca sa nu incarcam procesorul
            r = sm.ratio()
            if r >= 0.8:
                # unim in proprietatea cu mai multe interventii
                big, small = (p1, p2) if p1["interventii"] >= p2["interventii"] else (p2, p1)
                out.append({"din": small["key"], "din_nume": small["name"], "din_n": small["interventii"],
                            "in": big["key"], "in_nume": big["name"], "in_n": big["interventii"],
                            "scor": round(r, 2)})
    return sorted(out, key=lambda s: -s["scor"])[:limit]


def _groups_preview(interventii):
    """Pentru fiecare grup: ce clienti de pe facturi au intrat in el (ca sa se vada ce prinde regula)."""
    prev = {g["nume"]: {} for g in settings.get("grupuri_clienti") or []}
    for iv in group_clients(interventii):
        if iv.get("client_factura") is not None and iv.get("client") in prev:
            d = prev[iv["client"]]
            d[iv["client_factura"]] = d.get(iv["client_factura"], 0) + 1
    return [{"nume": n, "membri": sorted(d, key=str.casefold), "interventii": sum(d.values())}
            for n, d in prev.items()]


def _settings_payload():
    with state_lock:
        interventii = split_ivs(list(state["interventii"]))[0]
    props = aggregate_props(interventii)
    alias = settings["alias"]
    names = {p["key"]: p["name"] for p in props}
    # numele original (din facturi) al fiecarei variante, pentru lista de uniri
    variant_name = {base_prop_key(v): v for p in props for v in p["variante"]}
    return {
        "clienti_la_ora": settings["clienti_la_ora"],
        "tipuri_la_ora": settings["tipuri_la_ora"],
        "prag_forfait": settings["prag_forfait"],
        "cuvinte_produse": settings.get("cuvinte_produse") or [],
        "grupuri_clienti": settings.get("grupuri_clienti") or [],
        "grupuri_preview": _groups_preview(interventii),
        "separate": settings.get("separate") or [],
        "unite": [{"din": k, "din_nume": variant_name.get(k, k), "in": v,
                   "in_nume": names.get(prop_key(v), v)} for k, v in sorted(alias.items())],
        "proprietati": sorted(({"key": p["key"], "name": p["name"], "variante": p["variante"]}
                               for p in props), key=lambda p: p["name"].casefold()),
        "sugestii": merge_suggestions(props),
        "clienti": sorted({iv.get("client") for iv in interventii if iv.get("client")}),
        "baza": settings.get("baza"),
        "coord_sheet": settings.get("coord_sheet"),
        "geo_sync": settings.get("geo_sync"),
        "n_coords": len(settings.get("coords") or []),
        "n_ore": sum(1 for h in (settings.get("ore_estimate") or {}).values() if h),
        "formule": settings.get("formule") or [], "formule_vars": formulas.VARS,
        "ore_status": dict(hours_status),
    }


def run_geo_async(force=False):
    """Recalculeaza distantele in fundal (daca ruleaza o scanare, se fac oricum la finalul ei)."""
    def job():
        if not scan_lock.acquire(blocking=False):
            add_log("Distantele se recalculeaza la finalul scanarii in curs.")
            return
        try:
            sync_geo(force_routes=force)
            save_cache()
        except Exception as e:  # noqa: BLE001
            add_log(f"⚠ Distante: {e}")
        finally:
            set_stage(None)
            scan_lock.release()
    threading.Thread(target=job, daemon=True).start()


@app.route("/api/geo", methods=["POST"])
def api_geo():
    """Baza si Sheet-ul cu coordonate (salvate in backup, nu in cod)."""
    data = request.json or {}
    if "baza" in data:
        if not geo.parse_latlng(data["baza"]):
            return jsonify({"ok": False, "error": "Coordonatele bazei nu sunt valide (ex. 41.5955, 9.2746)"}), 400
        settings["baza"] = data["baza"].strip()
    if "coord_sheet" in data:
        settings["coord_sheet"] = sheet_id_from(data["coord_sheet"])
    cache["dirty"] = True
    add_log("Setari distante salvate — recalculez")
    run_geo_async(force=bool(data.get("force")))
    return jsonify(_settings_payload())


@app.route("/api/geo/link", methods=["POST"])
def api_geo_link():
    """Leaga manual o casa din facturi de un nume din Sheet (nume gol = anuleaza)."""
    data = request.json or {}
    key, nume = data.get("key"), (data.get("nume") or "").strip()
    if not key:
        return jsonify({"ok": False, "error": "Lipseste casa"}), 400
    if nume:
        settings.setdefault("coord_alias", {})[key] = nume
    else:
        (settings.get("coord_alias") or {}).pop(key, None)
    settings["geo_v"] = settings.get("geo_v", 0) + 1   # schimba cheia datelor -> recalcul
    cache["dirty"] = True
    _save_async()
    return jsonify({"ok": True})


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    if request.method == "POST":
        data = request.json or {}

        def clean_list(v):
            items = v.splitlines() if isinstance(v, str) else (v or [])
            return [s.strip() for s in items if s and s.strip()]
        if "clienti_la_ora" in data:
            settings["clienti_la_ora"] = clean_list(data["clienti_la_ora"])
        if "tipuri_la_ora" in data:
            settings["tipuri_la_ora"] = clean_list(data["tipuri_la_ora"])
        if "cuvinte_produse" in data:
            settings["cuvinte_produse"] = clean_list(data["cuvinte_produse"])
        if "grupuri_clienti" in data:
            # text: o linie pe grup -> "Nume grup: potrivire1, potrivire2"
            grupuri = []
            for line in clean_list(data["grupuri_clienti"]):
                if ":" not in line:
                    return jsonify({"ok": False, "error": f"Lipsește „:” în linia „{line}” "
                                    "(ex. Conciergerie X: Dupont, Martin)"}), 400
                nume, rest = line.split(":", 1)
                pot = [w.strip() for w in re.split(r"[,;]", rest) if w.strip()]
                if nume.strip() and pot:
                    grupuri.append({"nume": nume.strip(), "potriviri": pot})
            settings["grupuri_clienti"] = grupuri
        if "prag_forfait" in data:
            try:
                settings["prag_forfait"] = max(0.0, float(str(data["prag_forfait"]).replace(",", ".")))
            except ValueError:
                return jsonify({"ok": False, "error": "Prag invalid"}), 400
        recompute_all()
        add_log(f"Reguli actualizate: clienti la ora = {', '.join(settings['clienti_la_ora']) or '—'} · "
                f"tipuri la ora = {', '.join(settings['tipuri_la_ora']) or '—'} · "
                f"prag forfait = {settings['prag_forfait']:g} € · produse = "
                f"{', '.join(settings.get('cuvinte_produse') or []) or '—'} · grupuri clienti = "
                f"{len(settings.get('grupuri_clienti') or [])} — recalculat fara recitirea facturilor")
        _save_async()
    return jsonify(_settings_payload())


# ---------------------------------------------------------------- incasari (matrice client x luna)
LUNI_SCURT = ["Ian", "Feb", "Mar", "Apr", "Mai", "Iun", "Iul", "Aug", "Sep", "Oct", "Noi", "Dec"]


def _invoice_month(f, emis):
    """Luna lucrarilor (cea mai frecventa luna din datele menajelor); altfel luna emiterii."""
    cnt = {}
    for iv in f.get("interventii") or []:
        d = iv.get("data_iso") or ""
        if len(d) >= 7:
            cnt[d[:7]] = cnt.get(d[:7], 0) + 1
    if cnt:
        return max(cnt, key=lambda m: (cnt[m], m))
    return emis.strftime("%Y-%m") if emis else None


def invoice_ledger(today=None):
    """O intrare pe factura valida (fara duplicate/erori), cu scadenta, suma si cat s-a incasat."""
    today = today or date.today()
    rules = client_group_rules()
    pay = settings.get("incasari") or {}
    with state_lock:
        files = [dict(f) for f in state["files"].values() if f.get("status") in ("done", "warning")]
    out = []
    for f in files:
        total = float(f.get("total") or 0)
        if total <= 0:
            continue
        emis, _ = sane_invoice_date(parse_ddmmyyyy(f.get("data_emitere") or ""))
        due = parse_ddmmyyyy((f.get("antet") or {}).get("scadenta") or "")
        if emis and (not due or abs((due - emis).days) > 120):
            due = emis + timedelta(days=30)        # scadenta lipsa / tastata gresit
        luna = _invoice_month(f, emis)
        if not luna:
            continue
        client = (f.get("client") or "Necunoscut").strip()
        grup = group_of(client, norm_text(f.get("folder") or ""), rules) if rules else None
        p = pay.get(f["id"]) or {}
        inc = min(float(p.get("suma") or 0), total)
        rest = round(total - inc, 2)
        if rest <= 0.01:
            stare = "ok"
        elif due and due < today:
            stare = "late" if inc <= 0 else "part"
        else:
            stare = "part" if inc > 0 else "wait"
        out.append({
            "id": f["id"], "numar": f.get("numar"), "fisier": f.get("name"), "client": grup or client,
            "client_factura": client, "grup": bool(grup), "luna": luna, "total": round(total, 2),
            "incasat": round(inc, 2), "rest": max(rest, 0), "stare": stare,
            "emisa": emis.isoformat() if emis else None, "scadenta": due.isoformat() if due else None,
            "zile_intarziere": (today - due).days if due and due < today and rest > 0.01 else 0,
            "data_incasare": p.get("data"), "nota": p.get("nota") or "",
        })
    return out


def payments_matrix(an=None):
    today = date.today()
    led = invoice_ledger(today)
    ani = sorted({x["luna"][:4] for x in led}, reverse=True)
    an = an if an in ani else (ani[0] if ani else str(today.year))
    rows_led = [x for x in led if x["luna"][:4] == an]
    luni = sorted({x["luna"] for x in rows_led})
    clients = {}
    for x in rows_led:
        c = clients.setdefault(norm_text(x["client"]), {"client": x["client"], "grup": x["grup"], "celule": {},
                                                        "de_incasat": 0.0, "restant": 0.0})
        cell = c["celule"].setdefault(x["luna"], {"total": 0.0, "incasat": 0.0, "facturi": []})
        cell["total"] += x["total"]
        cell["incasat"] += x["incasat"]
        cell["facturi"].append(x)
        c["de_incasat"] += x["rest"]
        if x["stare"] in ("late", "part") and x["zile_intarziere"] > 0:
            c["restant"] += x["rest"]
    for c in clients.values():
        for cell in c["celule"].values():
            fs = sorted(cell["facturi"], key=lambda x: (x["emisa"] or "", x["numar"] or ""))
            cell["facturi"] = fs
            cell["total"], cell["incasat"] = round(cell["total"], 2), round(cell["incasat"], 2)
            cell["rest"] = round(cell["total"] - cell["incasat"], 2)
            late = [x for x in fs if x["zile_intarziere"] > 0]
            if cell["rest"] <= 0.01:
                cell["stare"] = "ok"            # verde: tot incasat
            elif cell["incasat"] > 0:
                cell["stare"] = "part"          # galben: o parte incasata
            elif late:
                cell["stare"] = "late"          # rosu: trecut de scadenta, nimic incasat
            else:
                cell["stare"] = "wait"          # alb: inca nescadent
            cell["zile"] = max((x["zile_intarziere"] for x in late), default=0)
            nesc = [x["scadenta"] for x in fs if x["rest"] > 0.01 and x["scadenta"]]
            cell["scadenta"] = min(nesc) if nesc else None
            cell["in_curs"] = any(x["luna"] == today.strftime("%Y-%m") for x in fs)
        c["de_incasat"], c["restant"] = round(c["de_incasat"], 2), round(c["restant"], 2)
    # ordine stabila (dupa cat s-a facturat in an), ca randurile sa nu sara cand marchezi o plata
    for c in clients.values():
        c["facturat"] = round(sum(cell["total"] for cell in c["celule"].values()), 2)
    rows = sorted(clients.values(), key=lambda c: (-c["facturat"], c["client"].casefold()))
    tot = []
    for m in luni:
        t = sum(x["total"] for x in rows_led if x["luna"] == m)
        i = sum(x["incasat"] for x in rows_led if x["luna"] == m)
        tot.append({"luna": m, "facturat": round(t, 2), "incasat": round(i, 2),
                    "procent": round(i / t * 100) if t else 0})
    all_rest = round(sum(x["rest"] for x in led), 2)
    all_late = round(sum(x["rest"] for x in led if x["zile_intarziere"] > 0), 2)
    return {
        "an": an, "ani": ani,
        "luni": [{"key": m, "label": LUNI_SCURT[int(m[5:]) - 1]} for m in luni],
        "clienti": rows, "totaluri": tot,
        "sumar": {"de_incasat": all_rest, "restant": all_late,
                  "facturi_restante": sum(1 for x in led if x["zile_intarziere"] > 0),
                  "de_incasat_an": round(sum(c["de_incasat"] for c in rows), 2)},
        "azi": today.isoformat(),
    }


@app.route("/api/incasari")
def api_incasari():
    if request.args.get("sumar"):          # doar pentru patratul "De incasat" (tab-ul inchis)
        led = invoice_ledger()
        return jsonify({"sumar": {"de_incasat": round(sum(x["rest"] for x in led), 2),
                                  "restant": round(sum(x["rest"] for x in led if x["zile_intarziere"] > 0), 2),
                                  "facturi_restante": sum(1 for x in led if x["zile_intarziere"] > 0)}, "doar_sumar": True})
    return jsonify(payments_matrix(request.args.get("an")))


@app.route("/api/incasari", methods=["POST"])
def api_incasari_set():
    """
    Marcheaza facturi incasate. {ids:[...]} = integral; {ids:[id], suma: 200} = partial;
    {ids:[...], undo:true} = anuleaza. Optional data (YYYY-MM-DD) si nota.
    """
    data = request.json or {}
    ids = [i for i in (data.get("ids") or []) if isinstance(i, str)]
    pana_la = str(data.get("pana_la") or "")
    if pana_la and not data.get("undo"):
        # pornire rapida: tot ce e neincasat pana la luna aleasa (inclusiv) devine incasat
        if not re.fullmatch(r"\d{4}-\d{2}", pana_la):
            return jsonify({"ok": False, "error": "Luna nu e validă"}), 400
        ids = [x["id"] for x in invoice_ledger() if x["luna"] <= pana_la and x["rest"] > 0.01]
        if not ids:
            return jsonify({"ok": True, "ids": []})
    if not ids:
        return jsonify({"ok": False, "error": "Nicio factură aleasă"}), 400
    pay = settings.setdefault("incasari", {})
    with state_lock:
        totals = {fid: float(state["files"][fid].get("total") or 0) for fid in ids if fid in state["files"]}
    if data.get("undo"):
        for fid in ids:
            pay.pop(fid, None)
    else:
        when = (data.get("data") or date.today().isoformat())[:10]
        try:
            date.fromisoformat(when)
        except ValueError:
            return jsonify({"ok": False, "error": "Data nu e validă"}), 400
        suma = data.get("suma")
        if suma is not None and len(ids) == 1:
            try:
                suma = float(str(suma).replace(",", ".").replace("€", "").replace(" ", ""))
            except ValueError:
                return jsonify({"ok": False, "error": "Suma nu e validă"}), 400
            if suma <= 0:
                pay.pop(ids[0], None)
            else:
                pay[ids[0]] = {"suma": round(min(suma, totals.get(ids[0], suma)), 2), "data": when,
                               "nota": (data.get("nota") or "").strip()[:200]}
        else:
            for fid in ids:
                if fid in totals:
                    pay[fid] = {"suma": round(totals[fid], 2), "data": when,
                                "nota": (data.get("nota") or "").strip()[:200]}
    cache["dirty"] = True
    _save_async()
    if pana_la:
        add_log(f"Încasări: {len(ids)} facturi până la {pana_la} marcate încasate")
    return jsonify({"ok": True, "ids": ids})


# ---------------------------------------------------------------- cheltuieli: matrice categorii x luni
def expenses_year(an=None):
    today_m = costs.today_month()
    entries = settings.get("cheltuieli") or []
    with state_lock:
        ivs_raw = state["interventii"]
    ivs_all = analysis_ivs(ivs_raw)[0]
    ani = sorted({e["data"][:4] for e in entries if e.get("data")} | {iv["data_iso"][:4] for iv in ivs_all if iv.get("data_iso")}
                 | {today_m[:4]}, reverse=True)
    an = an if an in ani else today_m[:4]
    used = [iv["data_iso"][:7] for iv in ivs_all if (iv.get("data_iso") or "")[:4] == an]
    used += [m for m in (e["luna"] for e in costs.expand(entries, today_m)) if m[:4] == an]
    first = min(used) if used else f"{an}-01"
    months = [m for m in analytics.months_range(first, f"{an}-12") if m <= today_m]
    rows = monthly_profit(months, ivs_all)
    occ = [e for e in costs.expand(entries, today_m) if e["luna"][:4] == an]
    cats = settings.get("categorii_cheltuieli") or costs.CATEGORII_IMPLICITE
    names = [c["nume"] for c in cats] + sorted({e["categorie"] for e in occ} - {c["nume"] for c in cats})
    est = {r["luna"]: r["motorina_estimata"] for r in rows}
    grid = []
    for name in names:
        cells = {}
        for e in occ:
            if e["categorie"] == name:
                c = cells.setdefault(e["luna"], {"total": 0.0, "intrari": []})
                c["total"] = round(c["total"] + float(e.get("suma") or 0), 2)
                c["intrari"].append({k: e.get(k) for k in ("id", "aparitie", "suma", "nota", "lunar", "serie", "luna")})
        grid.append({"categorie": name, "tip": costs.cat_type(name, cats), "motorina": costs.is_fuel(name),
                     "celule": cells, "total": round(sum(c["total"] for c in cells.values()), 2),
                     "estimat": {m: round(est.get(m, 0), 2) for m in months} if costs.is_fuel(name) else None})
    return {"an": an, "ani": ani, "luni": [{"key": m, "label": LUNI_SCURT[int(m[5:]) - 1]} for m in months],
            "categorii": grid, "lunar": rows,
            "total": {k: round(sum(r[k] for r in rows), 2) for k in ("venit", "cheltuieli", "profit", "personal", "ramas")},
            "consum": settings.get("consum"), "pret_motorina": settings.get("pret_motorina"), "azi": date.today().isoformat(),
            "tipuri": {c["nume"]: c.get("tip", "firma") for c in cats}}


@app.route("/api/cheltuieli")
def api_cheltuieli():
    return jsonify(expenses_year(request.args.get("an")))


def _cost_changed():
    settings["cost_v"] = settings.get("cost_v", 0) + 1
    cache["dirty"] = True
    _save_async()


@app.route("/api/cheltuieli", methods=["POST"])
def api_cheltuieli_add():
    """Adauga o cheltuiala (sau o pune inapoi dupa 'Anuleaza', cu acelasi id)."""
    import uuid
    d = request.json or {}
    try:
        suma = float(str(d.get("suma") or "").replace(",", ".").replace("€", "").replace(" ", ""))
    except ValueError:
        return jsonify({"ok": False, "error": "Suma nu e validă"}), 400
    data = str(d.get("data") or "")[:10]
    try:
        date.fromisoformat(data)
    except ValueError:
        return jsonify({"ok": False, "error": "Data nu e validă"}), 400
    cat = (d.get("categorie") or "").strip()
    if not cat or suma <= 0:
        return jsonify({"ok": False, "error": "Alege categoria și scrie o sumă mai mare ca 0"}), 400
    entries = settings.setdefault("cheltuieli", [])
    eid = str(d.get("id") or "")[:20] or uuid.uuid4().hex[:10]
    if any(e["id"] == eid for e in entries):
        eid = uuid.uuid4().hex[:10]
    e = {"id": eid, "data": data, "categorie": cat[:60], "suma": round(suma, 2),
         "nota": (d.get("nota") or "").strip()[:200], "lunar": bool(d.get("lunar"))}
    if d.get("pana_la"):
        e["pana_la"] = str(d["pana_la"])[:7]
    entries.append(e)
    _cost_changed()
    return jsonify({"ok": True, "entry": e})


@app.route("/api/cheltuieli/sterge", methods=["POST"])
def api_cheltuieli_del():
    """mod 'tot' = sterge cheltuiala (sau toata seria lunara); 'opreste' = seria se opreste inainte de luna data."""
    d = request.json or {}
    entries = settings.setdefault("cheltuieli", [])
    e = next((x for x in entries if x["id"] == d.get("id")), None)
    if not e:
        return jsonify({"ok": False, "error": "Cheltuiala nu mai există"}), 404
    old = dict(e)
    if d.get("mod") == "opreste" and e.get("lunar") and re.fullmatch(r"\d{4}-\d{2}", str(d.get("luna") or "")):
        stop = costs._add_months(d["luna"], -1)
        if stop < e["data"][:7]:
            entries.remove(e)
        else:
            e["pana_la"] = stop
    else:
        entries.remove(e)
    _cost_changed()
    return jsonify({"ok": True, "vechi": old})


@app.route("/api/cheltuieli/restaureaza", methods=["POST"])
def api_cheltuieli_restore():
    """'Anuleaza' dupa stergere: pune cheltuiala exact cum era."""
    old = (request.json or {}).get("vechi") or {}
    entries = settings.setdefault("cheltuieli", [])
    if not old.get("id"):
        return jsonify({"ok": False, "error": "Nimic de restaurat"}), 400
    entries[:] = [x for x in entries if x["id"] != old["id"]] + [{k: old[k] for k in
                  ("id", "data", "categorie", "suma", "nota", "lunar", "pana_la") if k in old}]
    _cost_changed()
    return jsonify({"ok": True})


@app.route("/api/cheltuieli/setari", methods=["POST"])
def api_cheltuieli_settings():
    d = request.json or {}
    for k in ("consum", "pret_motorina"):
        if k in d:
            try:
                settings[k] = max(0.0, float(str(d[k]).replace(",", ".")))
            except ValueError:
                return jsonify({"ok": False, "error": "Consum / preț invalid"}), 400
    if "categorii" in d:
        cats = []
        for line in str(d["categorii"]).splitlines():
            line = line.strip()
            if not line:
                continue
            personal = line.lower().endswith("(personal)")
            nume = line[: -len("(personal)")].strip() if personal else line
            if nume and all(c["nume"] != nume for c in cats):
                cats.append({"nume": nume[:60], "tip": "personal" if personal else "firma"})
        if not cats:
            return jsonify({"ok": False, "error": "Lista de categorii e goală"}), 400
        settings["categorii_cheltuieli"] = cats
    _cost_changed()
    return jsonify({"ok": True})


# ---------------------------------------------------------------- ore estimate -> Google Sheet
_hours_lock = threading.Lock()
hours_status = {"ok": None, "msg": ""}


def _col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


ORE_TAB = "Ore"   # foaia separata din acelasi document, doar cu nume + ore


_sheets_local = threading.local()


def _sheets_service():
    """Conexiunea la Sheets se creeaza o data pe fir si se refoloseste (nu la fiecare scriere)."""
    svc = getattr(_sheets_local, "service", None)
    if svc is None:
        info = json.loads(GOOGLE_CREDENTIALS_JSON)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        http = google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=40))
        svc = _sheets_local.service = build("sheets", "v4", http=http, cache_discovery=False)
    return svc


def write_hours_to_sheet(items):
    """
    items: [(nume casa, ore)]. Scrie in foaia separata "Ore" (nume | ore) din documentul cu coordonate;
    o creeaza la final daca nu exista. Foaia cu coordonate si coduri nu e atinsa.
    """
    sid = sheet_id_from(settings.get("coord_sheet"))
    if not sid:
        raise RuntimeError("Sheet-ul cu coordonate nu e configurat")
    svc = _sheets_service()
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets.properties(title)").execute()
    titles = [x["properties"]["title"] for x in meta.get("sheets", [])]
    tab = next((t for t in titles if t.strip().casefold() == ORE_TAB.casefold()), None)
    if not tab:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
            {"addSheet": {"properties": {"title": ORE_TAB}}}]}).execute()   # se adauga la final
        tab = ORE_TAB
        svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{tab}'!A1:B1", valueInputOption="RAW",
                                           body={"values": [["nume", "ore"]]}).execute()
    t = "'" + tab.replace("'", "''") + "'"
    vals = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"{t}!A1:B").execute().get("values", [])
    rows = {}
    for r, row in enumerate(vals[1:], start=2):
        if row and row[0].strip():
            rows.setdefault(norm_text(row[0]), r)
    data, new = [], []
    for nume, ore in items:
        r = rows.get(norm_text(nume))
        if r:
            data.append({"range": f"{t}!B{r}", "values": [[ore if ore else ""]]})
        elif ore:
            new.append([nume, ore])
    if data:
        svc.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={"valueInputOption": "RAW", "data": data}).execute()
    if new:
        svc.spreadsheets().values().append(spreadsheetId=sid, range=f"{t}!A:B", valueInputOption="RAW",
                                           insertDataOption="INSERT_ROWS", body={"values": new}).execute()
    return []


def read_hours_tab(sid):
    """
    Orele din foaia "Ore": prin Sheets API, altfel prin linkul public.
    Intoarce ({nume: ore}, {numele tuturor randurilor}) sau None daca foaia nu poate fi citita.
    """
    rows = None
    try:
        rows = _sheets_service().spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{ORE_TAB}'!A1:B").execute().get("values", [])
    except Exception:  # noqa: BLE001
        try:
            url = (f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet="
                   + urllib.parse.quote(ORE_TAB))
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30) as r:
                body = r.read().decode("utf-8-sig")
            if not body.lstrip().startswith("<"):
                import csv as _csv
                rows = list(_csv.reader(io.StringIO(body)))
        except Exception:  # noqa: BLE001
            rows = None
    if rows is None:
        return None
    out, names = {}, set()
    if rows and rows[0] and rows[0][0].strip().casefold() in geo.NAME_COLS:
        for row in rows[1:]:
            if not row or not row[0].strip():
                continue
            names.add(row[0].strip())
            if len(row) >= 2:
                try:
                    v = float(str(row[1]).replace(",", ".").strip())
                except ValueError:
                    continue
                if 0 < v <= 48:
                    out[row[0].strip()] = round(v, 2)
    return out, names


def _house_display_names():
    with state_lock:
        ivs = state["interventii"]
    return {p["key"]: p["name"] for p in aggregate_props(split_ivs(ivs)[0])}


def flush_hours():
    """Scrie in Sheet orele modificate in aplicatie (si reincearca ce nu s-a putut scrie)."""
    with _hours_lock:
        pending = list(settings.get("ore_nescrise") or [])
        if not pending:
            return
        ore = settings.get("ore_estimate") or {}
        items, no_row = [], []
        names = _house_display_names()
        for k in pending:
            # numele casei din facturi (nu cel din foaia cu coordonate): case diferite din aceeasi
            # rezidenta au aceleasi coordonate, dar ore si tarife diferite -> fiecare pe randul ei
            nume = names.get(k)
            (items.append((nume, ore.get(k))) if nume else no_row.append(k))
        try:
            missing = write_hours_to_sheet(items) if items else []
            settings["ore_nescrise"] = []        # scrise (sau fara rand in Sheet: raman doar in backup)
            n = len(items) - len(missing)
            extra = len(no_row) + len(missing)
            hours_status.update(ok=True, msg=f"✓ Scris în foaia „{ORE_TAB}” ({n})" + (f" · {extra} doar în backup" if extra else ""))
            add_log(f"Ore estimate: {n} scrise în Sheet" + (f", {extra} doar în backup" if extra else ""))
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            if "has not been used" in msg or "disabled" in msg or "accessNotConfigured" in msg:
                m = re.search(r"project[ =](\d+)", msg)
                msg = ("Google Sheets API nu e activat în proiectul service account-ului"
                       + (f" (proiectul nr. {m.group(1)}: activează-l aici → https://console.developers.google.com/apis/api/"
                          f"sheets.googleapis.com/overview?project={m.group(1)})" if m else ""))
            elif "PERMISSION_DENIED" in msg or "403" in msg or "permission" in msg.lower():
                msg = "service account-ul nu are drept de Editor pe Sheet"
            hours_status.update(ok=False, msg=f"Salvat în backup; nescris încă în Sheet ({msg}) — se reîncearcă")
            add_log(f"⚠ Ore estimate: {msg}")
        cache["dirty"] = True
        save_cache()


MAX_INTREBARI = 60


@app.route("/api/intrebari", methods=["POST"])
def api_intrebari():
    """Intrebarile pastrate (lista intreaga). Doar date simple, cu limite de marime."""
    lst = (request.json or {}).get("lista")
    if not isinstance(lst, list) or len(lst) > MAX_INTREBARI or not all(isinstance(q, dict) for q in lst):
        return jsonify({"ok": False, "error": f"Format greșit (cel mult {MAX_INTREBARI} întrebări)"}), 400
    for q in lst:
        if len(json.dumps(q, ensure_ascii=False)) > 12000:
            return jsonify({"ok": False, "error": "O întrebare e prea mare (concluzia sau istoricul)"}), 400
    if lst != settings.get("intrebari"):
        settings["intrebari"] = lst
        cache["dirty"] = True
        _save_async()
    return jsonify({"ok": True, "n": len(lst)})


COL_COLORS = {"verde", "albastru", "galben", "mov", "roz", "gri", "scala", "scala_inv"}


@app.route("/api/coloane", methods=["POST"])
def api_coloane():
    """Ordinea coloanelor (Proprietati / Clienti) si indicatorii ascunsi. Doar liste scurte de nume."""
    d = request.json or {}
    cur = dict(settings.get("coloane") or {})
    for k in ("props", "clienti", "ascunse"):
        if k in d:
            v = d.get(k)
            if not isinstance(v, list):
                return jsonify({"ok": False, "error": "Format greșit"}), 400
            cur[k] = [str(x)[:60] for x in v if isinstance(x, (str, int, float))][:80]
    if "culori" in d:          # culoarea aleasa pentru fiecare coloana: {cheie coloana: culoare}
        v = d.get("culori")
        if not isinstance(v, dict):
            return jsonify({"ok": False, "error": "Format greșit"}), 400
        cur["culori"] = {str(k)[:60]: str(c) for k, c in list(v.items())[:80] if str(c) in COL_COLORS}
    if "nume" in d:            # nume date de tine coloanelor: {cheie: "Nume nou"}
        v = d.get("nume")
        if not isinstance(v, dict):
            return jsonify({"ok": False, "error": "Format greșit"}), 400
        cur["nume"] = {str(k)[:60]: str(n).strip()[:40] for k, n in list(v.items())[:80] if str(n).strip()}
    if "latimi" in d:          # latimea fiecarei coloane, in pixeli
        v = d.get("latimi")
        if not isinstance(v, dict):
            return jsonify({"ok": False, "error": "Format greșit"}), 400
        cur["latimi"] = {str(k)[:60]: max(50, min(420, int(w))) for k, w in list(v.items())[:80]
                         if isinstance(w, (int, float))}
    if cur != settings.get("coloane"):
        settings["coloane"] = cur
        cache["dirty"] = True
        _save_async()
    return jsonify({"ok": True, "coloane": cur})


@app.route("/api/formule", methods=["POST"])
def api_formule():
    """Salveaza indicatorii proprii ('Nume = expresie' pe fiecare rand). Cu greseli: nu salveaza nimic."""
    text = (request.json or {}).get("text", "")
    lst, errors = formulas.parse_lines(text)
    if errors:
        return jsonify({"ok": False, "error": "; ".join(errors), "erori": errors}), 400
    settings["formule"] = lst
    cache["dirty"] = True
    add_log(f"Indicatori proprii salvați: {', '.join(f['nume'] for f in lst) or '—'}")
    _save_async()
    return jsonify({"ok": True, "formule": lst})


@app.route("/api/backup/versiuni")
def api_backup_versions():
    """Versiunile backup-ului pastrate de Google Drive (cele mai noi primele), ca sa recuperezi regulile."""
    if not cache.get("file_id"):
        return jsonify({"ok": False, "error": "Backup-ul nu e găsit încă în Drive"}), 400
    try:
        res = get_drive_service().revisions().list(
            fileId=cache["file_id"], fields="revisions(id, modifiedTime, size)", pageSize=200).execute(num_retries=DRIVE_RETRIES)
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": f"Nu pot citi versiunile din Drive: {e}"}), 502
    revs = sorted(res.get("revisions") or [], key=lambda r: r.get("modifiedTime") or "", reverse=True)
    return jsonify({"ok": True, "versiuni": [{"id": r["id"], "data": r.get("modifiedTime"), "kb": round(int(r.get("size") or 0) / 1024)}
                                             for r in revs[:60]]})


@app.route("/api/backup/versiuni", methods=["POST"])
def api_backup_version_restore():
    """Recupereaza SETARILE dintr-o versiune din Drive (doar ce lipseste acum; facturile nu se ating)."""
    rid = str((request.json or {}).get("id") or "")
    if not cache.get("file_id") or not rid:
        return jsonify({"ok": False, "error": "Alege o versiune"}), 400
    try:
        req = get_drive_service().revisions().get_media(fileId=cache["file_id"], revisionId=rid)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk(num_retries=DRIVE_RETRIES)
        data = json.loads(buf.getvalue().decode("utf-8-sig"))
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": f"Nu pot citi versiunea: {e}"}), 502
    return _restore_settings_from(data)


@app.route("/api/restaureaza-setari", methods=["POST"])
def api_restore_settings():
    """Incarci un cache-facturi.json vechi (din istoricul Drive) -> se recupereaza doar SETARILE."""
    f = request.files.get("fisier")
    try:
        raw = (f.read() if f else request.get_data()).decode("utf-8-sig")
        data = json.loads(raw)
    except Exception:  # noqa: BLE001
        return jsonify({"ok": False, "error": "Fișierul nu e un backup valid (cache-facturi.json)"}), 400
    return _restore_settings_from(data)


def _restore_settings_from(data):
    old = data.get("setari") if isinstance(data, dict) else None
    if not isinstance(old, dict):
        return jsonify({"ok": False, "error": "Backup-ul nu conține setări"}), 400
    got = merge_old_settings(old)
    if not got:
        return jsonify({"ok": True, "recuperat": [], "msg": "Nimic nou de recuperat din fișierul ăsta"})
    for k in ("geo_v", "cost_v", "ore_v", "folder_v"):
        settings[k] = settings.get(k, 0) + 1
    recompute_all()
    add_log(f"Setări recuperate din backup-ul vechi ({data.get('saved_at', '?')}): {', '.join(got)}")
    _save_async()
    if settings.get("coord_sheet") and "coord_sheet" in got:
        run_geo_async()
    return jsonify({"ok": True, "recuperat": got, "salvat_la": data.get("saved_at")})


@app.route("/api/ore/rescrie", methods=["POST"])
def api_ore_rescrie():
    """Rescrie in foaia "Ore" toate orele salvate in aplicatie (ex. dupa ce ai golit foaia)."""
    ore = settings.get("ore_estimate") or {}
    if not ore:
        return jsonify({"ok": False, "error": "Nu există ore salvate în aplicație"}), 400
    settings["ore_nescrise"] = sorted(k for k, h in ore.items() if h)
    flush_hours()
    return jsonify({"ok": True, "n": len(settings.get("ore_estimate") or {}), "status": dict(hours_status)})


@app.route("/api/ore", methods=["POST"])
def api_ore():
    d = request.json or {}
    key = (d.get("key") or "").strip()
    if not key:
        return jsonify({"ok": False, "error": "Lipsește casa"}), 400
    raw = str(d.get("ore") if d.get("ore") is not None else "").replace(",", ".").strip()
    ore = settings.setdefault("ore_estimate", {})
    if raw == "":
        ore.pop(key, None)
    else:
        try:
            v = float(raw)
        except ValueError:
            return jsonify({"ok": False, "error": "Număr de ore invalid"}), 400
        if not 0 < v <= 48:
            return jsonify({"ok": False, "error": "Orele trebuie să fie între 0 și 48"}), 400
        ore[key] = round(v, 2)
    settings["ore_nescrise"] = sorted(set(settings.get("ore_nescrise") or []) | {key})
    settings["ore_v"] = settings.get("ore_v", 0) + 1
    cache["dirty"] = True
    if settings.get("coord_sheet"):
        hours_status.update(ok=None, msg="Salvat · se scrie în foaia „Ore” în câteva secunde")
    else:
        settings["ore_nescrise"] = [k for k in settings["ore_nescrise"] if k != key]
        hours_status.update(ok=True, msg="Salvat în backup (Sheet-ul cu coordonate nu e configurat)")
    _save_async()     # orele introduse una dupa alta se scriu impreuna, intr-o singura cerere
    return jsonify({"ok": True, "status": dict(hours_status), "ore": ore.get(key)})


# ---------------------------------------------------------------- Google Calendar
def service_account_email():
    try:
        return json.loads(GOOGLE_CREDENTIALS_JSON).get("client_email") if GOOGLE_CREDENTIALS_JSON else None
    except ValueError:
        return None


def read_calendar():
    """Citeste evenimentele (doar data + titlu) din calendarul partajat cu service account-ul."""
    cid = (settings.get("calendar_id") or "").strip()
    if not cid:
        return
    set_stage("citesc Google Calendar")
    try:
        info = json.loads(GOOGLE_CREDENTIALS_JSON)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/calendar.readonly"])
        http = google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=60))
        svc = build("calendar", "v3", http=http, cache_discovery=False)
        start = settings.get("calendar_de_la") or (date.today() - timedelta(days=550)).isoformat()
        tmin = f"{start}T00:00:00Z"
        tmax = f"{(date.today() + timedelta(days=1)).isoformat()}T00:00:00Z"
        events, token = [], None
        while True:
            r = svc.events().list(calendarId=cid, timeMin=tmin, timeMax=tmax, singleEvents=True,
                                  orderBy="startTime", maxResults=2500, pageToken=token,
                                  fields="nextPageToken,items(summary,start,status)").execute()
            for ev in r.get("items", []):
                d = calendar_check.event_day(ev)
                if d and ev.get("status") != "cancelled" and ev.get("summary"):
                    events.append((d, ev["summary"].strip()[:200]))
            token = r.get("nextPageToken")
            if not token or len(events) > 20000:
                break
        with state_lock:
            state["calendar"] = {"events": events, "citit": datetime.now().strftime("%d/%m/%Y %H:%M"), "eroare": None}
        add_log(f"Google Calendar: {len(events)} evenimente citite")
    except Exception as e:  # noqa: BLE001
        msg = str(e)
        if "accessNotConfigured" in msg or "has not been used" in msg or "disabled" in msg:
            m = re.search(r"project[ =](\d+)", msg)
            msg = ("Google Calendar API nu e activat în proiectul service account-ului"
                   + (f" (proiectul nr. {m.group(1)}: https://console.developers.google.com/apis/api/"
                      f"calendar-json.googleapis.com/overview?project={m.group(1)})" if m else "") + ".")
        elif "notFound" in msg or "Not Found" in msg or "404" in msg:
            msg = "Calendarul nu e găsit: verifică ID-ul și că l-ai partajat cu email-ul service account-ului."
        elif "forbidden" in msg.lower() or "403" in msg:
            msg = "Acces refuzat: partajează calendarul cu service account-ul („Vezi toate detaliile evenimentelor”)."
        with state_lock:
            state["calendar"]["eroare"] = msg
        add_log(f"⚠ Google Calendar: {msg}")
    finally:
        set_stage(None)


def calendar_report():
    if not settings.get("calendar_id"):
        return {"configurat": False, "calendar_id": "", "de_la_setat": settings.get("calendar_de_la"), "citit": None,
                "eroare": None, "service_account": service_account_email(), "de_verificat": 0, "evenimente": 0,
                "nefacturat": [], "data_diferita": [], "fara_eveniment": [], "de_facturat": [], "neidentificate": [],
                "potrivite": 0, "facturat_pana_la": None, "de_la": None}
    with state_lock:
        cal = dict(state["calendar"])
        ivs_raw = state["interventii"]
    ivs = split_ivs(ivs_raw)[0]
    props = aggregate_props(ivs)
    names = {p["key"]: list(p["variante"]) + [p["name"]] for p in props}
    for c in settings.get("coords") or []:            # numele din Sheet ajuta si ele la potrivire
        names.setdefault(prop_key(c["nume"]), []).append(c["nume"])
    for k, nume in (settings.get("coord_alias") or {}).items():
        names.setdefault(k, []).append(nume)
    display = {p["key"]: p["name"] for p in props}
    inv = {}
    for iv in ivs:
        if iv.get("data_iso"):
            k = prop_key(re.sub(r"\s+", " ", (iv.get("proprietate") or "").strip()) or "Necunoscut")
            inv.setdefault((k, iv["data_iso"]), {"factura": iv.get("factura"), "client": iv.get("client"),
                                                  "fisier": iv.get("_source")})
    start = None
    if settings.get("calendar_de_la"):
        try:
            start = date.fromisoformat(settings["calendar_de_la"])
        except ValueError:
            start = None
    out = calendar_check.compare(cal.get("events") or [], inv, calendar_check.build_index(names), start,
                                 frozenset(settings.get("calendar_ok") or []))
    for lst in ("nefacturat", "data_diferita", "fara_eveniment", "de_facturat"):
        for x in out[lst]:
            x["nume"] = display.get(x["cheie"], x["cheie"])
    out.update({"configurat": bool(settings.get("calendar_id")), "calendar_id": settings.get("calendar_id"),
                "de_la_setat": settings.get("calendar_de_la"), "citit": cal.get("citit"), "eroare": cal.get("eroare"),
                "service_account": service_account_email()})
    out["de_verificat"] = sum(1 for k in ("nefacturat", "data_diferita") for x in out[k] if not x["ok"])
    return out


@app.route("/api/calendar")
def api_calendar():
    return jsonify(calendar_report())


@app.route("/api/calendar", methods=["POST"])
def api_calendar_set():
    d = request.json or {}
    if "calendar_id" in d:
        settings["calendar_id"] = str(d["calendar_id"] or "").strip()[:200]
    if "de_la" in d:
        v = str(d["de_la"] or "").strip()[:10]
        if v:
            try:
                date.fromisoformat(v)
            except ValueError:
                return jsonify({"ok": False, "error": "Data de început nu e validă"}), 400
        settings["calendar_de_la"] = v
    cache["dirty"] = True
    _save_async()
    if settings.get("calendar_id"):
        read_calendar()
    else:
        with state_lock:
            state["calendar"] = {"events": [], "citit": None, "eroare": None}
    return jsonify(calendar_report())


@app.route("/api/calendar/ok", methods=["POST"])
def api_calendar_ok():
    d = request.json or {}
    ids, cur = set(d.get("ids") or []), set(settings.get("calendar_ok") or [])
    settings["calendar_ok"] = sorted(cur - ids if d.get("undo") else cur | ids)
    cache["dirty"] = True
    _save_async()
    return jsonify({"ok": True})


@app.route("/api/exclude", methods=["POST"])
def api_exclude():
    """Scoate o proprietate din lista (ex. un produs facturat), sau o pune inapoi cu undo=true."""
    data = request.json or {}
    key, undo = (data.get("key") or "").strip(), bool(data.get("undo"))
    if not key:
        return jsonify({"ok": False, "error": "Lipseste proprietatea"}), 400
    cur, keep = set(settings.get("excluse") or []), set(settings.get("pastrate") or [])
    if undo:
        # "Pune inapoi": ramane in lista si daca numele contine un cuvant de produs
        settings["excluse"] = sorted(cur - {key})
        settings["pastrate"] = sorted(keep | {key})
    else:
        settings["excluse"] = sorted(cur | {key})
        settings["pastrate"] = sorted(keep - {key})
        # daca fusese legat gresit de o casa din Sheet (tab-ul Zone), legatura se sterge
        if (settings.get("coord_alias") or {}).pop(key, None) is not None:
            settings["geo_v"] = settings.get("geo_v", 0) + 1
    cache["dirty"] = True
    add_log(f"{'Pus inapoi in lista' if undo else 'Scos din lista'}: '{data.get('name') or key}'")
    _save_async()
    return jsonify({"ok": True})


@app.route("/api/merge", methods=["POST"])
def api_merge():
    """Uneste proprietatea `din` in proprietatea `in` (ex. 'Blue' in 'Villa Blue')."""
    data = request.json or {}
    src, dst = prop_key(data.get("din") or ""), prop_key(data.get("in") or "")
    if not src or not dst or src == dst:
        return jsonify({"ok": False, "error": "Alegeti doua proprietati diferite"}), 400
    settings["alias"][src] = dst
    cache["dirty"] = True
    add_log(f"Proprietati unite: '{data.get('din')}' → '{data.get('in')}'")
    _save_async()
    return jsonify(_settings_payload())


def warn_signature(f):
    """Amprenta problemelor unei facturi: daca apar probleme noi, 'verificat' nu se mai aplica."""
    import hashlib
    txt = (f.get("error") or "") + "|" + "|".join(f.get("avertismente") or [])
    return hashlib.md5(txt.encode("utf-8")).hexdigest()[:12]


def is_verified(f):
    return (f.get("status") in ("warning", "error")
            and (settings.get("verificate") or {}).get(f.get("id")) == warn_signature(f))


@app.route("/api/verify-files", methods=["POST"])
def api_verify_files():
    """'Am verificat' pe facturi (sau anulare cu undo=true)."""
    data = request.json or {}
    ids, undo = data.get("ids") or [], bool(data.get("undo"))
    ver = settings.setdefault("verificate", {})
    with state_lock:
        for fid in ids:
            f = state["files"].get(fid)
            if not f:
                continue
            if undo:
                ver.pop(fid, None)
            elif f.get("status") in ("warning", "error"):
                ver[fid] = warn_signature(f)
    cache["dirty"] = True
    _save_async()
    return jsonify({"ok": True, "n": len(ids)})


@app.route("/api/prices-ok", methods=["POST"])
def api_prices_ok():
    """'E corect' pe semnalari de pret (sau anulare cu undo=true)."""
    data = request.json or {}
    ids, undo = set(data.get("ids") or []), bool(data.get("undo"))
    cur = set(settings.get("preturi_ok") or [])
    settings["preturi_ok"] = sorted(cur - ids if undo else cur | ids)
    cache["dirty"] = True
    _save_async()
    return jsonify({"ok": True, "n": len(ids)})


@app.route("/api/ignore-suggestion", methods=["POST"])
def api_ignore_suggestion():
    """'Nu e dubla' — sugestia nu mai apare."""
    data = request.json or {}
    pair = "|".join(sorted((data.get("din") or "", data.get("in") or "")))
    settings.setdefault("ignorate", [])
    if pair not in settings["ignorate"]:
        settings["ignorate"].append(pair)
        cache["dirty"] = True
        _save_async()
    return jsonify(_settings_payload())


@app.route("/api/backup/reincarca", methods=["POST"])
def api_backup_reload():
    """Reincearca acum citirea backup-ului (reguli + facturi) din Drive."""
    if cache["loaded"] and cache["enabled"]:
        return jsonify({"ok": True, "msg": "Backup-ul e deja încărcat"})
    threading.Thread(target=scan_folder, daemon=True).start()
    return jsonify({"ok": True, "msg": "Caut backup-ul în Drive…"})


@app.route("/api/separa", methods=["POST"])
def api_separa():
    """Separa un nume unit automat ('Appartement Marina' de 'Marina'); undo=true il uneste la loc."""
    d = request.json or {}
    n = norm_text(d.get("nume") or "")
    if not n:
        return jsonify({"ok": False, "error": "Lipsește numele"}), 400
    cur = [x for x in (settings.get("separate") or []) if x != n]
    if not d.get("undo"):
        cur.append(n)
    settings["separate"] = sorted(cur)
    cache["dirty"] = True
    add_log(("Unit la loc: '" if d.get("undo") else "Separat (altă casă): '") + (d.get("nume") or "") + "'")
    _save_async()
    return jsonify({**_settings_payload(), "ok": True})


@app.route("/api/unmerge", methods=["POST"])
def api_unmerge():
    key = (request.json or {}).get("din")
    if settings["alias"].pop(key, None) is not None:
        cache["dirty"] = True
        add_log(f"Unire anulata: '{key}'")
        _save_async()
    return jsonify(_settings_payload())


@app.route("/healthz")
def healthz():
    return "ok"


@app.route("/api/debug")
def api_debug():
    """Diagnostic: la ce pas e scanarea si unde anume asteapta codul (stiva firului de executie)."""
    import sys
    import traceback
    with state_lock:
        sc = {k: v for k, v in state["scan"].items()}
    out = {
        "scanare_activa": sc.get("active"),
        "pas_curent": sc.get("stage"),
        "secunde_la_acest_pas": round(time.time() - (sc.get("stage_since") or time.time())),
        "drive_folder_id": FOLDER_ID[:6] + "…" if FOLDER_ID else None,
        "cheie_google_setata": bool(GOOGLE_CREDENTIALS_JSON),
        "backup": {"activ": cache["enabled"], "incarcat": cache["loaded"]},
        "ultimele_mesaje": state["log"][-15:],
    }
    frame = sys._current_frames().get(sc.get("thread")) if sc.get("active") else None
    if frame is not None:
        out["unde_asteapta"] = [line.strip() for line in traceback.format_stack(frame)[-8:]]
    return Response(json.dumps(out, ensure_ascii=False, indent=2), mimetype="application/json")


@app.route("/api/net")
def api_net():
    """Test de retea la cerere: fara IPv4 fortat se vede si daca IPv6 merge."""
    hosts = ["oauth2.googleapis.com", "www.googleapis.com"]
    return Response(json.dumps({"ipv4_fortat": FORCE_IPV4, "rezultate": [net_check(h) for h in hosts]},
                               ensure_ascii=False, indent=2), mimetype="application/json")


@app.route("/api/test")
def api_test():
    return jsonify({
        "ai_key": bool(LLM_API_KEY),
        "ai_model": LLM_MODEL,
        "folder_id": bool(FOLDER_ID),
        "files_count": len(state["files"]),
        "monitor_thread": _bg_started,
        "backup": {"activ": cache["enabled"], "fisier": CACHE_FILE_NAME,
                   "file_id": cache["file_id"], "nesalvat": cache["dirty"]},
    })


@app.route("/api/test-ai")
def api_test_ai():
    req = urllib.request.Request(f"{LLM_BASE_URL}/models",
                                 headers={"Authorization": f"Bearer {LLM_API_KEY}"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            ids = [m.get("id") for m in json.loads(r.read()).get("data", [])]
            return jsonify({"ok": True, "model_configurat": LLM_MODEL,
                            "model_disponibil": LLM_MODEL in ids, "modele": ids})
    except urllib.error.HTTPError as e:
        return jsonify({"ok": False, "status": e.code, "body": e.read().decode()[:500]}), e.code


@app.route("/api/reset", methods=["POST"])
def api_reset():
    """Sterge tot (inclusiv backup-ul) -> urmatoarea scanare reciteste toate facturile."""
    if not scan_lock.acquire(timeout=5):
        return jsonify({"ok": False, "error": "Scanare in curs, incercati dupa ce se termina"}), 409
    try:
        with state_lock:
            state.update({"files": {}, "interventii": [], "last_scan": None,
                          "monitoring": MONITOR_ON, "log": []})
            state["scan"].update({"done": 0, "total": 0, "eta": None})
        save_cache(force=True, allow_empty=True)
    finally:
        scan_lock.release()
    add_log("Reset complet (inclusiv backup-ul). Apasati Scanare pentru a reciti facturile.")
    return jsonify({"ok": True})


# pornim firele de fundal la import (gunicorn importa modulul, nu ruleaza __main__)
if os.environ.get("DISABLE_BACKGROUND") != "1":
    start_background()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
