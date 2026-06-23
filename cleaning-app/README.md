# Analizator Facturi Curățenie

Aplicație Python care citește automat facturile PDF din Google Drive, le analizează cu Claude AI și afișează un clasament al proprietăților după frecvența intervențiilor.

---

## Deployment pe Railway (recomandat — gratuit)

### Pasul 1 — Creează cont Service Account Google

1. Mergi la [console.cloud.google.com](https://console.cloud.google.com)
2. Creează un proiect nou (ex: `facturi-curatenie`)
3. Activează **Google Drive API**: APIs & Services → Enable APIs → caută "Google Drive API" → Enable
4. Creează Service Account: APIs & Services → Credentials → Create Credentials → Service Account
   - Nume: `facturi-reader`
   - Role: Viewer
5. Descarcă cheia JSON: click pe service account → Keys → Add Key → JSON
6. **Copiază email-ul service account** (ex: `facturi-reader@proiect.iam.gserviceaccount.com`)

### Pasul 2 — Dă acces la folderul Drive

1. Deschide folderul Drive cu facturile
2. Click dreapta → Share → adaugă email-ul service account de mai sus
3. Permisiune: **Viewer** — e suficient

### Pasul 3 — Deploy pe Railway

1. Mergi la [railway.app](https://railway.app) → New Project → Deploy from GitHub
2. Conectează repo-ul GitHub cu acest cod
3. Adaugă variabilele de mediu (Settings → Variables):

```
ANTHROPIC_API_KEY=sk-ant-...
GOOGLE_CREDENTIALS_JSON={"type":"service_account","project_id":"..."}  ← tot JSON-ul din fișierul descărcat
DRIVE_FOLDER_ID=1iMypzmKG0AQzTunicr0cwzy3-kKTuBcJ
POLL_INTERVAL=60
```

4. Deploy automat → vei primi un URL de forma `https://app-name.railway.app`
5. Adaugă URL-ul ca bookmark în browser

---

## Rulare locală (pentru testare)

```bash
pip install -r requirements.txt

export ANTHROPIC_API_KEY="sk-ant-..."
export GOOGLE_CREDENTIALS_JSON='{"type":"service_account",...}'
export DRIVE_FOLDER_ID="1iMypzmKG0AQzTunicr0cwzy3-kKTuBcJ"

python app.py
```

Deschide: http://localhost:5000

---

## Cum funcționează

- La pornire scanează automat folderul Drive
- Monitorizare activă = verifică folderul la fiecare 60 secunde
- Fișierele noi sunt detectate și analizate automat
- Dashboard-ul se actualizează live la fiecare 5 secunde
- Export Excel cu un click — clasament + detalii complete

---

## Structură proiect

```
cleaning-app/
├── app.py              # Backend Flask + logica Drive + Claude
├── templates/
│   └── index.html      # Dashboard web
├── requirements.txt    # Dependențe Python
├── Procfile            # Pentru Railway/Render
└── README.md
```
