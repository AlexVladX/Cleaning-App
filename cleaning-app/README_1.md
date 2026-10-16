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
| `BASE_COORDS` | nu | coordonatele bazei, implicit (se pot schimba din aplicație) |

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

Toți indicatorii se calculează din facturi și respectă **filtrul de perioadă** de sus: toată perioada, un an, sezonul (apr–oct), sezonul fierbinte (iul–aug) sau o lună.

| Tab | Ce arată |
|---|---|
| **Proprietăți** | clasament după încasări, intervenții, frecvență (intervenții/lună), € / intervenție, inactivitate; căutare și filtru „Inactive” |
| **Clienți** | venit, % din venit (concentrare: cât aduc primii 3), case, intervenții, € / intervenție |
| **Evoluție** | grafice lunare (venit, intervenții) și tabel cu comparația față de luna și anul trecut; luna curentă e marcată „în curs” |
| **Calendar** | harta zilelor lucrate, media menajelor pe zi a săptămânii, cele mai aglomerate zile |
| **Zone** | distanța pe drum de la bază la fiecare casă, venit pe zone (sub 10 / 10–20 / 20–35 / peste 35 km), € încasați pe km de drum, casele scumpe la drum, casele fără coordonate |
| **Încasări** | matrice clienți × luni, colorată: verde încasat, galben parțial, roșu restant, alb nescadent. Click pe căsuță → facturile lunii, cu „✓ Încasată”, „Parțial…” și „Toată luna încasată”. „Pornire rapidă” marchează dintr-o dată tot ce e plătit până la o lună. Încasările se salvează în backup. |
| **Cheltuieli & profit** | cheltuieli pe categorii × luni (unele lunare, recurente), profit și marjă, „unde se duc banii”. Motorina e estimată din traseele zilnice (baza → casele zilei, de la una la cea mai apropiată → baza); unde treci bonuri reale pe o lună, profitul folosește bonurile. Categoriile marcate „(personal)” (ex. transferuri acasă) nu scad din profitul firmei. În Proprietăți apare „Profit est.” pe casă: încasări minus drumul casei, salarii/produse/mâncare după numărul de menaje și restul cheltuielilor după venit. |
| **Verificare calendar** | compară menajele din Google Calendar cu cele facturate: **nefacturate** (în calendar, pe nicio factură), **dată diferită** (aceeași casă la 1–3 zile), **fără eveniment**. Configurare: activezi Google Calendar API în proiectul service account-ului, partajezi calendarul cu email-ul service account-ului („Vezi toate detaliile evenimentelor”) și lipești ID-ul calendarului în tab. Titlurile evenimentelor nu se salvează, se citesc doar pentru comparație. |
| **Verificare prețuri** | prețul fiecărei case, factură după factură: **abatere** (preț diferit izolat → probabil greșeală), **de confirmat** (ultima factură are alt preț), **schimbare de tarif** |
| **Facturi** | lista facturilor cu filtrele De verificat / Erori / Duplicate |

- **Ore / menaj** (Proprietăți): orele estimate pentru o persoană la un menaj forfait, editabile direct în tabel. Se salvează în backup și în coloana `ore` din Sheet-ul cu coordonate (o creează dacă lipsește). Din ele se calculează **Ore lucru** în perioadă și **€ / oră**. Pentru scriere în Sheet: activează **Google Sheets API** în proiectul service account-ului și partajează Sheet-ul cu service account-ul ca **Editor**. Aplicația scrie doar în coloana `ore`.
- **Coloane** (Proprietăți): alegi ce coloane vezi (Selectează tot / Deselectează tot / Implicit); alegerea rămâne salvată în browser.
- **Produse și alte rânduri care nu sunt case**: butonul ✕ din Proprietăți le scoate din listă (se pot pune înapoi din „Scoase din listă”). Rândurile care conțin cuvintele din **Reguli → Produse** (implicit „produit” și „fourniture”) se scot automat, și din facturile viitoare. Nu intră în venit, intervenții sau clasament. Siguranță: ✕ cere confirmare (cu avertisment dacă rândul pare o casă) și poate fi anulat imediat; cuvintele-cheie nu scot niciodată o casă din Sheet-ul cu coordonate sau una pusă înapoi manual.
- **Grupuri de clienți** (Reguli): un rând pe grup, `Nume grup: cuvânt, cuvânt`. O factură intră în grup dacă numele clientului sau folderul ei din Drive conține unul din cuvinte (ex. persoane fizice aduse de același colaborator). În Clienți apare un singur rând, cu lista clienților facturați; pe factură clientul rămâne neschimbat. Regula „Clienți facturați doar la oră” se aplică tot după numele de pe factură.
- O casă e **inactivă** când n-a mai avut menaj de mult mai mult timp decât intervalul ei obișnuit (de 3 ori intervalul tipic, cel puțin 45 de zile).
- Verificarea prețurilor compară facturile aceleiași case, pentru același client și același tip de menaj.
- **Export Excel** folosește perioada și sortarea alese și are foile: Clasament (cu distanță, timp și zonă), Detalii intervenții, Facturi, Zone, Clienți, Evoluție lunară, Verificare prețuri.

## Distanțe și zone

Coordonatele caselor vin din Google Sheet-ul tău. Aplicația păstrează **doar** coloanele `nume`, `lat` și `lng`. Orice altă coloană (ex. `cod`) e ignorată imediat și nu se salvează nicăieri.

Configurare, o singură dată:
1. Share pe Sheet către email-ul service account-ului, cu permisiunea **Viewer**. Fără share, aplicația încearcă linkul public.
2. În aplicație: **Reguli de calcul → Distanțe**. Completează coordonatele bazei și linkul Sheet-ului, apoi apasă **Salvează și calculează distanțele**.

Cum funcționează:
- Traseele pe drum (km și minute) vin gratuit de la serviciul public OSRM, într-o singură cerere pentru până la 90 de case. Rezultatele se salvează în backup, deci se recalculează doar pentru casele noi sau mutate.
- Dacă serviciul nu răspunde, distanța se estimează din linia dreaptă × 1,35 și e marcată „estimat”. La următoarea citire se reîncearcă traseul real.
- Sheet-ul se recitește la fiecare scanare. O casă din facturi cu alt nume decât în Sheet se leagă o singură dată din tab-ul **Zone**.
- **€ pe km** = media încasată pe intervenție / (2 × distanța), adică la un drum dus-întors.
- Linkul Sheet-ului se salvează în backup, nu în cod.

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
├── geo.py              # coordonate din Sheet, trasee OSRM, zone
├── costs.py            # cheltuieli, motorina estimata, profit
├── calendar_check.py   # comparatia Google Calendar ↔ facturi
├── templates/index.html
├── requirements.txt
├── Procfile            # pentru Railway/Heroku (Render folosește render.yaml)
└── README.md
render.yaml             # în rădăcina repo-ului: configurația Render
```
