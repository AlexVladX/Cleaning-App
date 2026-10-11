# Analizator Facturi Curățenie

Aplicație Python care citește automat facturile PDF SimNettoyage din Google Drive, extrage intervențiile pe fiecare proprietate și afișează un clasament (dashboard + export Excel).

## Cum citește facturile

1. **Parser direct (implicit, gratuit, exact).** `invoice_parser.py` citește pozițiile din PDF:
   - antet: client ("Facture pour"), număr factură, data emiterii, scadența
   - tabelul **Détail des prestations**, rând cu rând: casa, tipul serviciului, **toate datele** de curățenie, cantitate, preț unitar, total
   - totalurile (Subtotal, Ajustements, Total), cu verificare că rândurile dau Subtotalul
2. **AI doar ca rezervă.** Dacă o factură nu are șablonul SimNettoyage, iar `GROQ_API_KEY` e setat, AI-ul copiază doar datele brute. Calculele se fac tot în cod, iar factura e marcată „de verificat”.

### Reguli de calcul (în cod, nu în AI)
- **Forfait** (preț unitar ≥ `FORFAIT_MIN_PRICE`, implicit 40 €): fiecare dată = o intervenție, la prețul unitar.
- **La oră** (preț unitar < 40 €): orele se împart egal pe date (10 h, 5 date, 30 €/h → 5 × 2 h × 60 €).
- Datele din tabel sunt în format MM/DD/YYYY și se convertesc automat în DD/MM/YYYY.

### Avertismente automate
- cantitate ≠ număr de date, sau `cantitate × preț ≠ total`
- suma rândurilor ≠ Subtotal, ajustări nerepartizate
- același număr de factură la doi clienți diferiți
- duplicat (același număr, client și total), care este ignorat automat

---

## Variabile de mediu (Render → Environment)

| Variabilă | Obligatoriu | Descriere |
|---|---|---|
| `GOOGLE_CREDENTIALS_JSON` | da | tot JSON-ul cheii Service Account |
| `DRIVE_FOLDER_ID` | da | ID-ul folderului cu facturi |
| `POLL_INTERVAL` | nu | secunde între verificări (implicit 60) |
| `APP_PASSWORD` | recomandat | parola pentru dashboard (orice utilizator) |
| `GROQ_API_KEY` | nu | cheie gratuită de la console.groq.com, doar pentru rezervă AI |
| `LLM_MODEL` | nu | implicit `openai/gpt-oss-120b` |
| `LLM_BASE_URL` / `LLM_API_KEY` | nu | alt furnizor compatibil OpenAI (ex. Gemini) |
| `FORFAIT_MIN_PRICE` | nu | pragul forfait / oră (implicit 40) |

---

## Deployment gratuit pe Render (fără card)

Service Account-ul Google și share-ul pe folder rămân cele de pe Railway. Pentru unul nou:
1. [console.cloud.google.com](https://console.cloud.google.com) → activează **Google Drive API**
2. Credentials → Service Account → Keys → Add Key → JSON
3. Share pe folderul Drive către email-ul service account-ului, cu permisiune **Viewer**

Deploy:
1. Pune în repo-ul GitHub fișierul `render.yaml` (în rădăcină, lângă folderul `cleaning-app/`).
2. [dashboard.render.com](https://dashboard.render.com) → cont gratuit cu GitHub → **New → Blueprint** → alege repo-ul.
3. Completează valorile cerute: `GOOGLE_CREDENTIALS_JSON`, `DRIVE_FOLDER_ID`, `APP_PASSWORD`, iar `GROQ_API_KEY` dacă vrei rezerva AI.
4. Primești un URL `https://facturi-curatenie.onrender.com`.

**Cum se comportă planul gratuit:** după 15 minute fără vizite, serverul adoarme. La prima deschidere pornește în ~1 minut și rescanează automat tot folderul Drive, deci datele sunt complete. Cât timp ai dashboard-ul deschis, serverul rămâne treaz și monitorizarea merge normal.

---

## Rulare locală

```bash
pip install -r requirements.txt
export GOOGLE_CREDENTIALS_JSON='{"type":"service_account",...}'
export DRIVE_FOLDER_ID="..."
python app.py          # http://localhost:5000
```

Testare parser pe PDF-uri locale, fără Drive:

```bash
python invoice_parser.py Facture_NET-26-001.pdf Facture_NET-26-010.pdf
```

## Rute utile
- `/api/test`: configurația activă
- `/api/test-ai`: verifică cheia AI și dacă modelul există
- `/api/factura/<file_id>`: tot ce s-a extras dintr-o factură

## Indicatori (dashboard pe tab-uri)

Toți indicatorii se calculează din facturi și respectă **filtrul de perioadă** de sus: toată perioada, un an, sezonul (iun–sep) sau o lună.

| Tab | Ce arată |
|---|---|
| **Proprietăți** | clasament după încasări, intervenții, frecvență (intervenții/lună), € / intervenție, inactivitate; căutare și filtru „Inactive” |
| **Clienți** | venit, % din venit (concentrare: cât aduc primii 3), case, intervenții, € / intervenție |
| **Evoluție** | grafice lunare (venit, intervenții) și tabel cu comparația față de luna și anul trecut; luna curentă e marcată „în curs” |
| **Calendar** | harta zilelor lucrate, media menajelor pe zi a săptămânii, cele mai aglomerate zile |
| **Verificare prețuri** | prețul fiecărei case, factură după factură: **abatere** (preț diferit izolat → probabil greșeală), **de confirmat** (ultima factură are alt preț), **schimbare de tarif** |
| **Facturi** | lista facturilor cu filtrele De verificat / Erori / Duplicate |

- O casă e **inactivă** când n-a mai avut menaj de mult mai mult timp decât intervalul ei obișnuit (de 3 ori intervalul tipic, cel puțin 45 de zile).
- Verificarea prețurilor compară facturile aceleiași case, pentru același client și același tip de menaj.
- **Export Excel** folosește perioada și sortarea alese și are foile: Clasament, Detalii intervenții, Facturi, Clienți, Evoluție lunară, Verificare prețuri.

## Backup în Google Drive (recomandat)

Rezultatele scanărilor se salvează în fișierul `cache-facturi.json`, din folderul facturilor. La pornire, aplicația îl încarcă și citește doar facturile **noi sau modificate**. Facturile șterse din Drive dispar din aplicație, iar cele redenumite se actualizează fără să fie recitite.

Configurare, o singură dată:
1. Urcă fișierul `cache-facturi.json` (conține doar `{}`) în **folderul cu facturi**.
2. Click dreapta pe fișier → **Share** → email-ul service account-ului → permisiunea **Editor**. Folderul rămâne la **Viewer**.
3. Gata. În jurnal apare „Backup încărcat: N facturi”, iar `/api/test` arată `"backup": {"activ": true}`.

Fără fișier, aplicația funcționează ca înainte (doar în memorie).

- **Reset** golește și backup-ul, iar la următoarea scanare se recitesc toate facturile.
- Când se modifică regulile de calcul din `invoice_parser.py`, schimbă `PARSER_VERSION` din `app.py`. Facturile se recitesc atunci automat, o singură dată.
- Facturile cu eroare se reîncearcă la pornire și la „Scanare manuală”, nu la fiecare verificare automată.

## Limitări
- PDF-urile scanate (imagini) nu au text, deci apar ca eroare.

## Structură
```
cleaning-app/
├── app.py              # Flask, Google Drive, agregări, export Excel
├── invoice_parser.py   # parserul facturilor SimNettoyage
├── analytics.py        # indicatorii: perioade, clienți, evoluție, calendar, prețuri
├── templates/index.html
├── requirements.txt
├── Procfile            # pentru Railway/Heroku (Render folosește render.yaml)
└── README.md
render.yaml             # în rădăcina repo-ului: configurația Render
```
