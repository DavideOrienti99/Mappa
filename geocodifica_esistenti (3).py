"""
Geocodifica SOLO le sedi esistenti: legge l'Excel e aggiunge Latitudine e Longitudine.
Provider gratuito: Nominatim/OSM.
"""

from pathlib import Path
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

BASE_DIR = Path(__file__).resolve().parent

# ============================================================================
# MODO SICURO: incolla qui il percorso completo del tuo file e sei a posto.
# Lascia None per usare la ricerca automatica piu' sotto.
# ============================================================================
FILE_INPUT = r"C:\Users\david\GitHub\Mappa\SEMPLICE\SEDI_ATTUALI.xlsx"

# Ricerca automatica (usata solo se FILE_INPUT = None).
# Sono in ordine di priorita': "attual" viene scelto prima di "input", ecc.
KEYWORDS = ["attual", "esistenti", "geocodif", "input"]
ESCLUDI = ["nuov", "2000", "esistenti_geo", "unificate", "output"]

OUT_XLSX = BASE_DIR / "SEDI_ESISTENTI_GEO.xlsx"
OUT_CSV = BASE_DIR / "SEDI_ESISTENTI_GEO.csv"


def trova_file_input() -> Path:
    if FILE_INPUT:
        p = Path(FILE_INPUT)
        if not p.exists():
            raise FileNotFoundError(f"Il percorso indicato in FILE_INPUT non esiste:\n  {p}")
        print(f"--> Uso il file indicato in FILE_INPUT: {p}\n")
        return p

    tutti = [f for f in BASE_DIR.rglob("*.xlsx") if not f.name.startswith("~$")]
    print("File .xlsx trovati nelle cartelle vicine:")
    for f in tutti:
        print(f"  - {f}")

    candidati = []
    for f in tutti:
        nome = f.name.lower()
        if any(e in nome for e in ESCLUDI):
            continue
        rank = next((i for i, k in enumerate(KEYWORDS) if k in nome), None)
        if rank is not None:
            candidati.append((rank, len(f.parts), f))

    if not candidati:
        raise FileNotFoundError(
            "Nessun file riconosciuto come 'sedi esistenti'.\n"
            "Apri lo script e incolla il percorso completo nella variabile FILE_INPUT."
        )
    candidati.sort(key=lambda t: (t[0], t[1]))   # prima per priorita' keyword, poi per vicinanza
    scelto = candidati[0][2]
    print(f"\n--> Uso questo file: {scelto}\n")
    return scelto


geolocator = Nominatim(user_agent="sedi-esistenti-geo/1.0")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=2, error_wait_seconds=2)


def indirizzo_completo(r):
    parti = [r.get("Indirizzo"), r.get("CAP"), r.get("Comune"), r.get("Provincia"), r.get("Regione"), "Italia"]
    return ", ".join(str(p).strip() for p in parti if pd.notna(p) and str(p).strip())


def solo_comune(r):
    parti = [r.get("Comune"), r.get("Provincia"), r.get("Regione"), "Italia"]
    return ", ".join(str(p).strip() for p in parti if pd.notna(p) and str(p).strip())


INPUT = trova_file_input()
df = pd.read_excel(INPUT, dtype=str)
df.columns = df.columns.str.strip()

cache, lat, lon = {}, [], []
for i, r in df.iterrows():
    q = indirizzo_completo(r)
    if q not in cache:
        loc = geocode(q, country_codes="it")
        if loc is None:
            loc = geocode(solo_comune(r), country_codes="it")
        cache[q] = (loc.latitude, loc.longitude) if loc else (None, None)
    la, lo = cache[q]
    lat.append(la)
    lon.append(lo)
    print(f"{i + 1}/{len(df)}  {q}  ->  {la}, {lo}")

df["Latitudine"] = lat
df["Longitudine"] = lon

df.to_excel(OUT_XLSX, index=False)
df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

print(f"\nFatto. Righe con coordinate: {df['Latitudine'].notna().sum()}/{len(df)}")
print(f"Excel: {OUT_XLSX}")
print(f"CSV:   {OUT_CSV}")
