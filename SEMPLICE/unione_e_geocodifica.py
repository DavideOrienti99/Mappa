"""
Unione di SEDI ATTUALI + SEDI NUOVE e geocodifica per Power BI (mappa Azure).

Cosa fa:
1. Legge i due file Excel di input.
2. Aggiunge la colonna 'tipo punto' (Esistenti / Nuovi) per distinguere l'origine.
3. Unisce tutto in un'unica tabella.
4. Per ogni riga calcola Latitudine e Longitudine in base a Indirizzo + Comune + Provincia.
5. Salva un unico file Excel e CSV pronti per la mappa Azure di Power BI.

Provider di geocodifica selezionabile con la variabile PROVIDER:
- "nominatim" : OpenStreetMap, gratuito (consigliato fino a poche centinaia di righe)
- "google"    : Google Geocoding API (richiede GOOGLE_MAPS_API_KEY) - preciso, a pagamento
- "azure"     : Azure Maps (richiede AZURE_MAPS_KEY) - stesso ecosistema della mappa Power BI

Per ~2000+ righe si consiglia "google" o "azure": piu' veloci e con termini d'uso
adatti al volume. Nominatim pubblico e' limitato a circa 1 richiesta/secondo.
"""

import logging
import os
import time
from pathlib import Path

import pandas as pd
from geopy.exc import GeocoderServiceError, GeocoderTimedOut, GeopyError


# ----------------------------- CONFIGURAZIONE -----------------------------

# Rinomina qui i tuoi file se hanno un nome diverso.
INPUT_SEDI_ATTUALI = Path("SEDI_ATTUALI.xlsx")
INPUT_SEDI_NUOVE = Path("SEDI_NUOVE_2000.xlsx")

OUTPUT_EXCEL = Path("SEDI_UNIFICATE_GEO.xlsx")
OUTPUT_CSV = Path("SEDI_UNIFICATE_GEO.csv")

# Provider: "nominatim" | "google" | "azure"
PROVIDER = "nominatim"

# Etichette colonna 'tipo punto'
ETICHETTA_ESISTENTI = "Esistenti"
ETICHETTA_NUOVI = "Nuovi"

# Colonne chiave comuni ai due file
INDIRIZZO = "Indirizzo"
COMUNE = "Comune"
CAP = "CAP"
PROVINCIA = "Provincia"
REGIONE = "Regione"

# Pausa fra richieste (Nominatim vuole >= 1s; Google/Azure tollerano molto meno)
DELAY = {"nominatim": 1.1, "google": 0.05, "azure": 0.05}
MAX_RETRIES = 3

USER_AGENT = "sedi-unificate-powerbi/1.0"


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


# ----------------------------- UTILITY -----------------------------

def clean_value(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def to_float(value: object):
    try:
        v = float(str(value).replace(",", "."))
        return v
    except (ValueError, TypeError):
        return None


def build_full_address(row: pd.Series) -> str:
    return (
        f"{clean_value(row.get(INDIRIZZO))}, {clean_value(row.get(CAP))} "
        f"{clean_value(row.get(COMUNE))}, {clean_value(row.get(PROVINCIA))}, "
        f"{clean_value(row.get(REGIONE))}, Italia"
    )


def build_fallback_address(row: pd.Series) -> str:
    return (
        f"{clean_value(row.get(INDIRIZZO))}, {clean_value(row.get(COMUNE))}, "
        f"{clean_value(row.get(PROVINCIA))}, Italia"
    )


def build_city_fallback_address(row: pd.Series) -> str:
    return (
        f"{clean_value(row.get(COMUNE))}, {clean_value(row.get(PROVINCIA))}, "
        f"{clean_value(row.get(REGIONE))}, Italia"
    )


# ----------------------------- GEOCODER -----------------------------

def make_geolocator(provider: str):
    if provider == "nominatim":
        from geopy.geocoders import Nominatim
        return Nominatim(user_agent=USER_AGENT)
    if provider == "google":
        from geopy.geocoders import GoogleV3
        key = os.getenv("GOOGLE_MAPS_API_KEY")
        if not key:
            raise ValueError("Variabile ambiente GOOGLE_MAPS_API_KEY non impostata.")
        return GoogleV3(api_key=key)
    if provider == "azure":
        from geopy.geocoders import AzureMaps
        key = os.getenv("AZURE_MAPS_KEY")
        if not key:
            raise ValueError("Variabile ambiente AZURE_MAPS_KEY non impostata.")
        return AzureMaps(subscription_key=key)
    raise ValueError(f"Provider non valido: {provider}")


def geocode_once(geolocator, provider: str, query: str):
    if provider == "nominatim":
        return geolocator.geocode(query, country_codes="it", addressdetails=True, timeout=10)
    if provider == "google":
        return geolocator.geocode(query, region="it", timeout=10)
    if provider == "azure":
        return geolocator.geocode(query, timeout=10)
    return None


def geocode_with_retry(geolocator, provider: str, query: str):
    last_error = ""
    delay = DELAY.get(provider, 1.1)
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            location = geocode_once(geolocator, provider, query)
            time.sleep(delay)
            return location, ""
        except (GeocoderTimedOut, GeocoderServiceError) as exc:
            last_error = str(exc)
            logging.warning("Tentativo %s/%s fallito per '%s': %s", attempt, MAX_RETRIES, query, last_error)
            time.sleep(delay * attempt)
        except GeopyError as exc:
            last_error = str(exc)
            logging.error("Errore geocoding per '%s': %s", query, last_error)
            time.sleep(delay)
            break
    return None, last_error


def geocode_address(geolocator, provider: str, row: pd.Series) -> dict:
    full = build_full_address(row)

    if not clean_value(row.get(INDIRIZZO)) or not clean_value(row.get(COMUNE)):
        return {"Indirizzo_Completo": full, "Latitudine": None, "Longitudine": None,
                "Indirizzo_Trovato": "", "Stato_Geocoding": "DATI_INSUFFICIENTI"}

    for query, stato in (
        (full, "TROVATO"),
        (build_fallback_address(row), "TROVATO_FALLBACK"),
        (build_city_fallback_address(row), "TROVATO_COMUNE"),
    ):
        location, error = geocode_with_retry(geolocator, provider, query)
        if location:
            return {"Indirizzo_Completo": full, "Latitudine": location.latitude,
                    "Longitudine": location.longitude, "Indirizzo_Trovato": location.address,
                    "Stato_Geocoding": stato}

    return {"Indirizzo_Completo": full, "Latitudine": None, "Longitudine": None,
            "Indirizzo_Trovato": "", "Stato_Geocoding": "NON_TROVATO"}


# ----------------------------- LETTURA / UNIONE -----------------------------

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]
    return df


def read_input(path: Path, tipo: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"File '{path}' non trovato nella cartella.")
    # legge il primo foglio che contiene almeno Indirizzo + Comune
    sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    for name, df in sheets.items():
        df = normalize_columns(df)
        if INDIRIZZO in df.columns and COMUNE in df.columns:
            df["tipo punto"] = tipo
            logging.info("'%s' -> foglio '%s', %s righe, tipo '%s'", path.name, name, len(df), tipo)
            return df
    raise ValueError(f"In '{path}' nessun foglio ha le colonne {INDIRIZZO} e {COMUNE}.")


def has_existing_coords(row: pd.Series) -> bool:
    if "Latitudine" not in row or "Longitudine" not in row:
        return False
    return to_float(row.get("Latitudine")) is not None and to_float(row.get("Longitudine")) is not None


# ----------------------------- MAIN -----------------------------

def main() -> None:
    df_att = read_input(INPUT_SEDI_ATTUALI, ETICHETTA_ESISTENTI)
    df_new = read_input(INPUT_SEDI_NUOVE, ETICHETTA_NUOVI)

    # unione: tiene tutte le colonne di entrambi, riempie i mancanti con vuoto
    df = pd.concat([df_att, df_new], ignore_index=True, sort=False)
    logging.info("Tabella unificata: %s righe totali", len(df))

    geolocator = make_geolocator(PROVIDER)
    cache: dict[str, dict] = {}
    results = []
    riusate = 0

    for index, row in df.iterrows():
        # 1) se la riga ha gia' coordinate valide (es. SEDI ATTUALI gia' geocodificate), le riuso
        if has_existing_coords(row):
            results.append({
                "Indirizzo_Completo": build_full_address(row),
                "Latitudine": to_float(row.get("Latitudine")),
                "Longitudine": to_float(row.get("Longitudine")),
                "Indirizzo_Trovato": clean_value(row.get("Indirizzo_Trovato")),
                "Stato_Geocoding": clean_value(row.get("Stato_Geocoding")) or "RIUSATO",
            })
            riusate += 1
            continue

        # 2) altrimenti geocodifico (con cache sugli indirizzi identici)
        key = build_full_address(row)
        logging.info("Riga %s/%s: %s", index + 1, len(df), key)
        if key not in cache:
            cache[key] = geocode_address(geolocator, PROVIDER, row)
        else:
            logging.info("Indirizzo gia' elaborato, uso cache.")
        results.append(cache[key])

    geo = pd.DataFrame(results)
    # sovrascrive/crea le colonne geo senza duplicarle
    for col in ["Indirizzo_Completo", "Latitudine", "Longitudine", "Indirizzo_Trovato", "Stato_Geocoding"]:
        df[col] = geo[col].values
    df["Provider_Geocoding"] = PROVIDER

    df.to_excel(OUTPUT_EXCEL, index=False, engine="openpyxl")
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    con_coord = df["Latitudine"].notna().sum()
    print("\nRiepilogo")
    print("---------")
    print(f"Righe totali:           {len(df)}")
    print(f"  di cui Esistenti:     {(df['tipo punto'] == ETICHETTA_ESISTENTI).sum()}")
    print(f"  di cui Nuovi:         {(df['tipo punto'] == ETICHETTA_NUOVI).sum()}")
    print(f"Coordinate riusate:     {riusate}")
    print(f"Righe con coordinate:   {con_coord}")
    print(f"Righe senza coordinate: {len(df) - con_coord}")
    print(f"Provider usato:         {PROVIDER}")
    print(f"File Excel: {OUTPUT_EXCEL}")
    print(f"File CSV:   {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
