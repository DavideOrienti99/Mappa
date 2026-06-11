"""
Geocodifica SOLO le sedi esistenti: legge l'Excel e aggiunge Latitudine e Longitudine.
Provider gratuito: OpenStreetMap / Nominatim.
"""

from pathlib import Path
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# Il file viene cercato nella stessa cartella di questo script.
BASE_DIR = Path(__file__).resolve().parent
INPUT = BASE_DIR / "SEDI ATTUALI.xlsx"          # <-- rinomina qui se il tuo file ha un altro nome
OUT_XLSX = BASE_DIR / "SEDI_ESISTENTI_GEO.xlsx"
OUT_CSV = BASE_DIR / "SEDI_ESISTENTI_GEO.csv"

# 1 richiesta/secondo (policy Nominatim) + qualche retry automatico
geolocator = Nominatim(user_agent="sedi-esistenti-geo/1.0")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=2, error_wait_seconds=2)


def indirizzo_completo(r):
    parti = [r.get("Indirizzo"), r.get("CAP"), r.get("Comune"), r.get("Provincia"), r.get("Regione"), "Italia"]
    return ", ".join(str(p).strip() for p in parti if pd.notna(p) and str(p).strip())


def solo_comune(r):
    parti = [r.get("Comune"), r.get("Provincia"), r.get("Regione"), "Italia"]
    return ", ".join(str(p).strip() for p in parti if pd.notna(p) and str(p).strip())


df = pd.read_excel(INPUT, dtype=str)
df.columns = df.columns.str.strip()

cache, lat, lon = {}, [], []
for i, r in df.iterrows():
    q = indirizzo_completo(r)
    if q not in cache:
        loc = geocode(q, country_codes="it")
        if loc is None:                       # se non trova il civico, prova col solo comune
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
