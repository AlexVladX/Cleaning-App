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
import socket
import threading
import time
import urllib.request
import urllib.error
from datetime import datetime

from flask import Flask, render_template, jsonify, send_file, request, Response
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import google_auth_httplib2
import httplib2
import openpyxl
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
PARSER_VERSION = "2026-10-11b"

cache = {"file_id": None, "loaded": False, "dirty": False, "enabled": False}

# Reguli de calcul + unificari de proprietati. Se editeaza din dashboard si se salveaza in backup.
settings = {
    **DEFAULT_RULES,
    "clienti_la_ora": [c.strip() for c in os.environ.get(
        "CLIENTI_LA_ORA", "Aquila Vacances;Esther").split(";") if c.strip()],
    "alias": {},   # cheie proprietate -> cheie proprietate canonica (unire manuala)
}

# Cuvinte generice ignorate cand comparam numele proprietatilor ("Villa Blue" = "Blue")
GENERIC_PROP_WORDS = {"villa", "maison", "appartement", "appart", "appt", "apt", "app",
                      "chalet", "residence", "gite", "house", "casa", "studio", "logement",
                      "menage", "la", "le", "les", "l"}


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
              "numar", "client", "data_emitere", "total", "metoda", "_dup_key", "processed_at",
              "randuri", "antet", "av_factura")
RULE_KEYS = ("prag_forfait", "tipuri_la_ora", "clienti_la_ora")


def current_rules():
    return {k: settings[k] for k in RULE_KEYS}


def recompute_all():
    """
    Recalculeaza interventiile tuturor facturilor din randurile salvate, cu regulile curente.
    Nu descarca si nu reciteste niciun PDF — dureaza o fractiune de secunda.
    """
    rules = current_rules()
    with state_lock:
        active = [f for f in state["files"].values()
                  if f.get("status") in ("done", "warning") and f.get("randuri")]
        by_nr = {}
        for f in active:
            if f.get("numar"):
                by_nr.setdefault(f["numar"], []).append(f)
        for f in active:
            ivs, calc_w = compute_interventions(f["randuri"], f.get("antet") or {}, rules)
            for iv in ivs:
                iv["_source"], iv["_file_id"] = f["name"], f["id"]
            same = [o for o in by_nr.get(f.get("numar"), []) if o is not f]
            nr_w = [f"Numarul {f['numar']} e folosit si de {o['name']} (client {o.get('client')})" for o in same]
            f["interventii"] = ivs
            f["avertismente"] = calc_w + list(f.get("av_factura") or []) + nr_w
            f["status"] = "warning" if f["avertismente"] else "done"
        state["interventii"] = [iv for f in state["files"].values()
                                if f.get("status") in ("done", "warning")
                                for iv in f.get("interventii") or []]
    cache["dirty"] = True


def apply_saved_settings(saved):
    if not isinstance(saved, dict):
        return
    for k in RULE_KEYS + ("alias",):
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
            if entry.get("_dup_key") is not None:
                entry["_dup_key"] = tuple(entry["_dup_key"])
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
                          "setari": {k: settings[k] for k in RULE_KEYS + ("alias",)},
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
def _dup_key(result):
    a, t = result["antet"], result["totaluri"]
    return ((a.get("numar") or "").upper(), (a.get("client") or "").casefold(), t.get("total"))


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
            dup = next((f for f in others if f.get("_dup_key") == _dup_key(result)), None)
            same_nr = [f for f in others if a.get("numar") and f.get("numar") == a.get("numar")]

            entry = state["files"][file_id]
            entry.update({
                "numar": a.get("numar"), "client": a.get("client"),
                "data_emitere": a.get("data_emitere"), "total": t.get("total"),
                "metoda": result["metoda"], "_dup_key": _dup_key(result),
                # date brute pastrate in backup -> regulile se pot schimba fara recitirea PDF-ului
                "randuri": result["randuri"], "antet": a,
                "av_factura": result.get("avertismente_factura", []),
            })
            if dup:
                entry["status"] = "duplicate"
                entry["avertismente"] = [f"Duplicat al fisierului {dup['name']} — ignorat"]
                entry["interventii"] = []
                state["interventii"] = [iv for iv in state["interventii"] if iv.get("_file_id") != file_id]
            else:
                for f in same_nr:
                    warnings.append(f"Numarul {a['numar']} e folosit si de {f['name']} (client {f.get('client')})")
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


def build_excel():
    with state_lock:
        interventii = list(state["interventii"])
        files = [dict(f) for f in state["files"].values()]
    props = aggregate_props(interventii)

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
                "Ouverture", "Ore (la ora)", "€/ora", "Venit total (EUR)", "Medie / interventie",
                "Ultima interventie", "Performanta", "Apare in facturi ca"],
          [6, 26, 24, 11, 10, 13, 11, 11, 8, 15, 17, 16, 13, 36])
    for i, p in enumerate(props, 1):
        ultima = datetime.fromisoformat(p["ultima"]).strftime("%d/%m/%Y") if p["ultima"] else ""
        t = p["tipuri"]
        ws1.append([i, p["name"], p["client"], p["interventii"], t.get("standard") or None,
                    t.get("intermediaire") or None, t.get("ouverture") or None, p["ore"] or None,
                    p["pret_ora"], p["total"], p["medie"], ultima, p["perf"],
                    ", ".join(p["variante"]) if len(p["variante"]) > 1 else ""])
        fill = {"Performant": "D1FAE5", "Rar": "FEE2E2"}.get(p["perf"])
        for col in range(1, 15):
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
        interventii = list(state["interventii"])
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
        "warnings": sum(1 for f in files if f["status"] in ("warning", "error")),
        "interventii": len(interventii),
        "revenue": round(sum(float(iv.get("suma") or 0) for iv in interventii), 2),
    }
    s["props"] = aggregate_props(interventii)
    return jsonify(s)


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
    buf = build_excel()
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"analiza-curatenie-{datetime.now():%Y-%m-%d}.xlsx",
    )


# ---------------------------------------------------------------- setari & unire proprietati
def _save_async():
    threading.Thread(target=save_cache, daemon=True).start()


def _settings_payload():
    with state_lock:
        interventii = list(state["interventii"])
    props = aggregate_props(interventii)
    alias = settings["alias"]
    names = {p["key"]: p["name"] for p in props}
    return {
        "clienti_la_ora": settings["clienti_la_ora"],
        "tipuri_la_ora": settings["tipuri_la_ora"],
        "prag_forfait": settings["prag_forfait"],
        "unite": [{"din": k, "in": v, "in_nume": names.get(prop_key(v), v)} for k, v in sorted(alias.items())],
        "proprietati": sorted(({"key": p["key"], "name": p["name"], "variante": p["variante"]}
                               for p in props), key=lambda p: p["name"].casefold()),
        "clienti": sorted({iv.get("client") for iv in interventii if iv.get("client")}),
    }


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
        if "prag_forfait" in data:
            try:
                settings["prag_forfait"] = max(0.0, float(str(data["prag_forfait"]).replace(",", ".")))
            except ValueError:
                return jsonify({"ok": False, "error": "Prag invalid"}), 400
        recompute_all()
        add_log(f"Reguli actualizate: clienti la ora = {', '.join(settings['clienti_la_ora']) or '—'} · "
                f"tipuri la ora = {', '.join(settings['tipuri_la_ora']) or '—'} · "
                f"prag forfait = {settings['prag_forfait']:g} € — recalculat fara recitirea facturilor")
        _save_async()
    return jsonify(_settings_payload())


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
