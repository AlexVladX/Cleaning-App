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
import urllib.error
from datetime import datetime, date

from flask import Flask, render_template, jsonify, send_file, request, Response
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import google_auth_httplib2
import httplib2
import openpyxl
import analytics
import geo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from invoice_parser import (
    parse_invoice, compute_interventions, layout_text, parse_ddmmyyyy, ParseError,
    norm_text, canonical_type, DEFAULT_RULES, TIP_LABEL,
)

app = Flask(__name__)

# ---------------------------------------------------------------- configurare
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")
POLL_INTERVAL = int(os.environ.get("POLL_INTERVAL", "60"))

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

cache = {"file_id": None, "loaded": False, "dirty": False, "enabled": False}

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
def base_prop_key(name):
    """Cheia automata (fara unirile manuale): 'Villa Blue' -> 'blue'."""
    words = norm_text(name).split()
    while len(words) > 1 and words[0] in GENERIC_PROP_WORDS:
        words = words[1:]
    return " ".join(words) or norm_text(name)


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


state = {
    "files": {},
    "interventii": [],
    "last_scan": None,
    "monitoring": False,
    "log": [],
    "scan": {"active": False, "phase": None, "done": 0, "total": 0, "eta": None},
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
            ).execute()
            for f in results.get("files", []):
                if f["mimeType"] == "application/pdf":
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
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def file_signature(f):
    """Amprenta continutului PDF: se schimba doar daca se schimba fisierul (nu la redenumire)."""
    return f.get("md5Checksum") or f.get("modifiedTime") or ""


# ---------------------------------------------------------------- backup in Drive
CACHE_KEYS = ("id", "name", "sig", "status", "error", "avertismente", "interventii",
              "numar", "client", "data_emitere", "total", "metoda", "processed_at",
              "randuri", "antet", "av_factura")
RULE_KEYS = ("prag_forfait", "tipuri_la_ora", "clienti_la_ora")
PERSIST_KEYS = ("alias", "ignorate", "verificate", "preturi_ok", "excluse", "cuvinte_produse", "pastrate", "baza", "coord_sheet", "coords",
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


def find_cache_file():
    if CACHE_FILE_ID:
        return CACHE_FILE_ID
    name = CACHE_FILE_NAME.replace("'", "\\'")
    svc = get_drive_service()
    set_stage(f"caut fisierul de backup '{CACHE_FILE_NAME}'")
    res = svc.files().list(
        q=f"name = '{name}' and '{FOLDER_ID}' in parents and trashed = false",
        fields="files(id, name, capabilities/canEdit)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    found = res.get("files", [])
    if not found:
        return None
    if not found[0].get("capabilities", {}).get("canEdit"):
        add_log(f"⚠ Backup: {CACHE_FILE_NAME} exista, dar service account-ul nu e Editor pe el")
        return None
    return found[0]["id"]


def load_cache():
    """Incarca scanarile salvate. Fara backup -> aplicatia merge ca inainte (doar in memorie)."""
    cache["loaded"] = True
    try:
        fid = find_cache_file()
    except TimeoutError:
        cache["loaded"] = False  # problema de retea, nu de backup: oprim scanarea, reincercam data viitoare
        raise
    except Exception as e:
        add_log(f"⚠ Backup indisponibil ({e}) — continui fara backup")
        return
    if not fid:
        add_log(f"Backup dezactivat: nu gasesc '{CACHE_FILE_NAME}' (editabil) in folderul facturilor")
        return
    cache.update({"file_id": fid, "enabled": True})
    try:
        set_stage("descarc backup-ul")
        raw = download_pdf(fid).read().decode("utf-8").strip()
        data = json.loads(raw) if raw else {}
    except Exception as e:
        add_log(f"⚠ Backup ilizibil ({e}) — il refac de la zero")
        cache["dirty"] = True
        return
    apply_saved_settings(data.get("setari"))  # setarile raman valabile si cand se schimba parserul
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
    add_log(f"Backup incarcat: {len(files)} facturi (salvat {data.get('saved_at', '?')})")


def save_cache(force=False):
    if not cache["enabled"] or not (cache["dirty"] or force):
        return
    with state_lock:
        files = {fid: {k: f.get(k) for k in CACHE_KEYS}
                 for fid, f in state["files"].items()
                 if f.get("status") in ("done", "warning", "duplicate", "error")}
        cache["dirty"] = False
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
        ).execute()
    except Exception as e:
        cache["dirty"] = True
        add_log(f"⚠ Nu am putut salva backup-ul: {e}")


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
    for i, f in enumerate(new_files):
        used_ai = process_file(f, i + 1, n)
        elapsed = time.time() - t0
        eta = round(elapsed / (i + 1) * (n - i - 1))
        with state_lock:
            state["scan"].update({"done": i + 1, "total": n, "eta": eta})
        if (i + 1) % 10 == 0 and i + 1 < n:
            add_log(f"— Progres {i + 1}/{n} · mai sunt {n - i - 1} · "
                    f"{elapsed / (i + 1):.1f}s per factura · ~{max(1, round(eta / 60))} min ramase")
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
        return
    if not scan_lock.acquire(blocking=False):
        with state_lock:
            stage, since = state["scan"].get("stage"), state["scan"].get("stage_since") or time.time()
        add_log(f"Scanare deja in curs — pasul curent: {stage} ({time.time() - since:.0f}s)")
        return
    scan_id = time.time()
    with state_lock:
        state["scan"] = {"active": True, "phase": "listare", "done": 0, "total": 0,
                         "eta": None, "found": 0, "current": None, "id": scan_id,
                         "stage": "pornire", "stage_since": time.time(),
                         "thread": threading.get_ident()}
    threading.Thread(target=_watchdog, args=(scan_id,), daemon=True).start()
    add_log("Pornesc scanarea...")
    try:
        if not cache["loaded"]:
            load_cache()
        files = list_pdf_files(FOLDER_ID)
        with state_lock:
            state["scan"]["phase"] = "procesare"
            known = {fid: dict(f) for fid, f in state["files"].items()}  # copie, nu referinte
        in_drive = {f["id"] for f in files}

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
            for f in renamed:
                state["files"][f["id"]]["name"] = f["name"]
                for iv in state["files"][f["id"]].get("interventii") or []:
                    iv["_source"] = f["name"]
            for f in todo:
                prev = state["files"].get(f["id"], {})
                state["files"][f["id"]] = {
                    **prev, "id": f["id"], "name": f["name"], "status": "new",
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
        save_cache()
    except Exception as e:
        add_log(f"Eroare scanare: {e}")
        _drive_local.service = None  # la urmatoarea scanare ne reconectam de la zero
        save_cache()
    finally:
        with state_lock:
            state["scan"]["active"] = False
        scan_lock.release()


def monitor_loop():
    while True:
        time.sleep(POLL_INTERVAL)
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
        t = threading.Timer(5.0, scan_folder, kwargs={"retry_errors": True})
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
                                   "variante": {}, "tipuri": {}})
        p["interventii"] += 1
        p["variante"][raw] = p["variante"].get(raw, 0) + 1
        tip = iv.get("tip_menaj") or canonical_type(iv.get("tip_serviciu"))
        p["tipuri"][tip] = p["tipuri"].get(tip, 0) + 1
        suma = float(iv.get("suma") or 0)
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
        p["client"] = ", ".join(sorted(p.pop("clients")))
        p["ore"] = round(p["ore"], 2)
        p["total"] = round(p["total"], 2)
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
        data = get_drive_service().files().export_media(fileId=sheet_id, mimeType="text/csv").execute()
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
    settings["coords"] = coords
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
    cache["dirty"] = True


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
               tuple(settings.get("cuvinte_produse") or []), tuple(sorted(settings.get("pastrate") or []))))
    g = hash((settings.get("baza"), len(settings.get("coords") or []), len(settings.get("distante") or {}),
              tuple(sorted((settings.get("coord_alias") or {}).items())), settings.get("geo_sync"), settings.get("geo_v")))
    return f"{id(ivs)}-{len(ivs)}-{alias}-{ok}-{g}"


def compute_analytics(period_key="all"):
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
    ivs_all, ex_all = split_ivs(ivs_raw)
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
        },
        "excluse": excluse,
        "props": props,
        "clients": clients, "top3": top3,
        "monthly": analytics.monthly_series(ivs_all, chart_months, prop_key),
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


def build_excel(sort="interventii", period_key="all"):
    with state_lock:
        files = [dict(f) for f in state["files"].values()]
    an = compute_analytics(period_key)
    interventii = analytics.filter_ivs(split_ivs(state["interventii"])[0], an["period"])
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
                "Performanta", "Apare in facturi ca"],
          [6, 26, 24, 11, 10, 13, 11, 11, 8, 13, 17, 15, 16, 15, 9, 12, 14, 12, 13, 36])
    for i, p in enumerate(props, 1):
        ultima = datetime.fromisoformat(p["ultima"]).strftime("%d/%m/%Y") if p.get("ultima") else ""
        t = p["tipuri"]
        ws1.append([i, p["name"], p["client"], p["interventii"], t.get("standard") or None,
                    t.get("intermediaire") or None, t.get("ouverture") or None, p["ore"] or None,
                    p["pret_ora"], p["total"], p["medie"], p.get("frecventa_luna"), ultima,
                    p.get("zile_de_la_ultima"), "da" if p.get("inactiva") else "",
                    p.get("distanta_km"), p.get("timp_min"), p.get("zona") or "", p["perf"],
                    ", ".join(p["variante"]) if len(p["variante"]) > 1 else ""])
        fill = {"Performant": "D1FAE5", "Rar": "FEE2E2"}.get(p["perf"])
        for col in range(1, 21):
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

    # perioada exportata, vizibila in titlul primei foi
    ws1.title = f"Clasament {an['period']['label']}"[:31].replace("/", "-")

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
    with state_lock:
        interventii = split_ivs(list(state["interventii"]))[0]
        files = sorted((_public_file(f, light=True) for f in state["files"].values()),
                       key=_file_sort_key)
        s = {
            "files": files,
            "now": time.time(),
            "last_scan": state["last_scan"],
            "monitoring": state["monitoring"],
            "log": state["log"][-50:],
            "scan": dict(state["scan"]),
        }
    s["metrics"] = {
        "files": len(files),
        "analyzed": sum(1 for f in files if f["status"] in ("done", "warning", "duplicate")),
        "warnings": sum(1 for f in files if f["status"] in ("warning", "error") and not f.get("verificat")),
        "interventii": len(interventii),
        "revenue": round(sum(float(iv.get("suma") or 0) for iv in interventii), 2),
    }
    s["data_key"] = data_key()  # dashboard-ul cere /api/analytics doar cand se schimba
    return jsonify(s)


@app.route("/api/analytics")
def api_analytics():
    return jsonify(compute_analytics(request.args.get("period", "all")))


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
    buf = build_excel(request.args.get("sort", "interventii"), request.args.get("period", "all"))
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"analiza-curatenie-{datetime.now():%Y-%m-%d}.xlsx",
    )


# ---------------------------------------------------------------- setari & unire proprietati
def _save_async():
    threading.Thread(target=save_cache, daemon=True).start()


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
        if "prag_forfait" in data:
            try:
                settings["prag_forfait"] = max(0.0, float(str(data["prag_forfait"]).replace(",", ".")))
            except ValueError:
                return jsonify({"ok": False, "error": "Prag invalid"}), 400
        recompute_all()
        add_log(f"Reguli actualizate: clienti la ora = {', '.join(settings['clienti_la_ora']) or '—'} · "
                f"tipuri la ora = {', '.join(settings['tipuri_la_ora']) or '—'} · "
                f"prag forfait = {settings['prag_forfait']:g} € · produse = "
                f"{', '.join(settings.get('cuvinte_produse') or []) or '—'} — recalculat fara recitirea facturilor")
        _save_async()
    return jsonify(_settings_payload())


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
                          "monitoring": False, "log": []})
            state["scan"].update({"done": 0, "total": 0, "eta": None})
        save_cache(force=True)
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
