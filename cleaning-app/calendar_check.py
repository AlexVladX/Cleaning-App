"""
Verificarea calendarului cu interventii fata de facturi.

  - fiecare eveniment din Google Calendar e legat de casa (sau casele) al caror nume apare in titlu
    (ex. "Ménage Villa Emma 45 – Ouverture" -> Emma 45); potrivirea se face pe cuvinte intregi,
    deci "Emma 4" nu se confunda cu "Emma 45"
  - apoi se compara (casa, zi) din calendar cu (casa, zi) din facturi:
      nefacturat     = in calendar, dar nu pe nicio factura (bani pierduti!)
      data diferita  = aceeasi casa, la 1-3 zile distanta (aproape sigur o data tastata gresit)
      fara eveniment = pe factura, dar nu in calendar
  - evenimentele de dupa ultima zi facturata sunt doar "de facturat", nu greseli
  - titlurile evenimentelor raman doar in memorie (nu se salveaza in backup)
"""
import re
import unicodedata
from datetime import date, timedelta

MAX_GAP = 3   # zile: peste atat nu mai e "data diferita", ci doua probleme separate


def norm_tokens(text):
    t = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode().lower()
    return re.findall(r"[a-z0-9]+", t)


def event_day(ev):
    st = ev.get("start") or {}
    d = st.get("date") or (st.get("dateTime") or "")[:10]
    try:
        return date.fromisoformat(d)
    except ValueError:
        return None


def build_index(names_by_key):
    """names_by_key: {cheie casa: [nume posibile]} -> lista (tokeni, cheie), cele mai lungi primele."""
    idx = {}
    for key, names in names_by_key.items():
        for n in list(names) + [key]:
            toks = tuple(norm_tokens(n))
            if toks and len("".join(toks)) >= 3:
                idx.setdefault(toks, key)
    return sorted(idx.items(), key=lambda x: (-len(x[0]), -len("".join(x[0]))))


def houses_in_title(title, index):
    toks = norm_tokens(title)
    used = [False] * len(toks)
    found = []
    for pat, key in index:
        n = len(pat)
        for i in range(len(toks) - n + 1):
            if tuple(toks[i:i + n]) == pat and not any(used[i:i + n]):
                for j in range(i, i + n):
                    used[j] = True
                if key not in found:
                    found.append(key)
    return found


def compare(events, inv, index, start=None, ok=frozenset()):
    """
    events: [(date, titlu)]; inv: {(cheie, iso_date): {"factura","client","nume"}}.
    Intoarce dict cu listele de probleme + ce nu s-a putut lega de nicio casa.
    """
    cal, unknown = {}, []
    for d, title in events:
        if start and d < start:
            continue
        keys = houses_in_title(title, index)
        if not keys:
            unknown.append({"data": d.isoformat(), "titlu": title})
        for k in keys:
            cal.setdefault((k, d.isoformat()), title)
    inv_days = sorted(d for _, d in inv)
    billed_until = inv_days[-1] if inv_days else None
    first_cal = min((d for _, d in cal), default=None)
    lo = max(x for x in (start.isoformat() if start else None, first_cal, inv_days[0] if inv_days else None) if x) \
        if (first_cal or inv_days) else None

    def in_window(d):
        return lo and d >= lo and (not billed_until or d <= billed_until)

    cal_left = {k for k in cal if in_window(k[1])}
    pending = sorted(({"cheie": k, "data": d, "titlu": cal[(k, d)]} for k, d in cal
                      if billed_until and d > billed_until), key=lambda x: x["data"])
    inv_left = {k for k in inv if in_window(k[1])}
    both = cal_left & inv_left
    cal_left -= both
    inv_left -= both

    shifted = []
    for k, d in sorted(cal_left):
        dd = date.fromisoformat(d)
        best = None
        for g in range(1, MAX_GAP + 1):
            for cand in ((dd - timedelta(days=g)).isoformat(), (dd + timedelta(days=g)).isoformat()):
                if (k, cand) in inv_left:
                    best = cand
                    break
            if best:
                break
        if best:
            inv_left.discard((k, best))
            shifted.append({"cheie": k, "data_calendar": d, "data_factura": best, "titlu": cal[(k, d)], **inv[(k, best)]})
    shifted_cal = {(x["cheie"], x["data_calendar"]) for x in shifted}
    unbilled = [{"cheie": k, "data": d, "titlu": cal[(k, d)]} for k, d in sorted(cal_left) if (k, d) not in shifted_cal]
    no_event = [{"cheie": k, "data": d, **inv[(k, d)]} for k, d in sorted(inv_left)]

    def tag(items, kind, dkey):
        for x in items:
            x["id"] = f"{kind}|{x['cheie']}|{x[dkey]}"
            x["ok"] = x["id"] in ok
        return items
    return {
        "nefacturat": tag(unbilled, "nf", "data"),
        "data_diferita": tag(shifted, "dd", "data_calendar"),
        "fara_eveniment": tag(no_event, "fe", "data"),
        "de_facturat": pending,
        "neidentificate": sorted(unknown, key=lambda x: x["data"], reverse=True)[:200],
        "potrivite": len(both),
        "facturat_pana_la": billed_until,
        "de_la": lo,
        "evenimente": len(events),
    }
