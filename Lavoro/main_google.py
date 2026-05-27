import logging
import os
import time
from pathlib import Path

import pandas as pd
from geopy.exc import GeocoderServiceError, GeocoderTimedOut, GeopyError
from geopy.geocoders import GoogleV3


INPUT_FILE = Path("sedi_input.xlsx")
OUTPUT_EXCEL = Path("sedi_geocodificate_google.xlsx")
OUTPUT_CSV = Path("sedi_geocodificate_google.csv")

API_KEY_ENV = "GOOGLE_MAPS_API_KEY"
REQUEST_DELAY_SECONDS = 0.2
MAX_RETRIES = 3

REQUIRED_COLUMNS = [
    "Tipologia",
    "STATO PDV",
    "Rag. Sociale",
    "Codice Agenzia",
    "Codici FR",
    "Indirizzo",
    "Comune",
    "CAP",
    "Provincia",
    "Regione",
    "Sede Operativa SUP",
]


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def clean_value(value: object) -> str:
    """Converte valori Excel vuoti o numerici in stringhe pulite."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]

    return text


def build_full_address(row: pd.Series) -> str:
    """Costruisce la query principale per Google Geocoding."""
    indirizzo = clean_value(row.get("Indirizzo"))
    cap = clean_value(row.get("CAP"))
    comune = clean_value(row.get("Comune"))
    provincia = clean_value(row.get("Provincia"))
    regione = clean_value(row.get("Regione"))

    return f"{indirizzo}, {cap} {comune}, {provincia}, {regione}, Italia"


def build_fallback_address(row: pd.Series) -> str:
    """Costruisce una query semplificata se la prima non trova risultati."""
    indirizzo = clean_value(row.get("Indirizzo"))
    comune = clean_value(row.get("Comune"))
    provincia = clean_value(row.get("Provincia"))

    return f"{indirizzo}, {comune}, {provincia}, Italia"


def build_city_fallback_address(row: pd.Series) -> str:
    """Costruisce una query comunale per ottenere almeno un punto approssimato."""
    comune = clean_value(row.get("Comune"))
    provincia = clean_value(row.get("Provincia"))
    regione = clean_value(row.get("Regione"))

    return f"{comune}, {provincia}, {regione}, Italia"


def validate_columns(df: pd.DataFrame) -> None:
    """Controlla che il file Excel contenga le colonne attese."""
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        missing = ", ".join(missing_columns)
        found = ", ".join(repr(column) for column in df.columns)
        raise ValueError(
            f"Colonne mancanti nel file Excel: {missing}\n"
            f"Colonne trovate nel file Excel: {found}"
        )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalizza nomi colonna ed elimina colonne vuote o senza nome."""
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]
    return df


def load_input_dataframe() -> pd.DataFrame:
    """Legge automaticamente il foglio Excel che contiene l'anagrafica sedi."""
    xls = pd.ExcelFile(INPUT_FILE)
    print("Fogli disponibili:", xls.sheet_names)

    for sheet in xls.sheet_names:
        print(f"\nAnalisi foglio: {sheet}")
        df = pd.read_excel(INPUT_FILE, sheet_name=sheet, dtype=str)
        df = normalize_columns(df)

        missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
        if missing_columns:
            print("Foglio ignorato: colonne richieste mancanti.")
            print("Colonne trovate:")
            for col in df.columns:
                print(repr(col))
            continue

        print(f"Foglio usato: {sheet}")
        print("Colonne trovate:")
        for col in df.columns:
            print(repr(col))
        print("Prime 5 righe:")
        print(df.head(5).to_string())

        return df

    raise ValueError(
        "Nessun foglio contiene la tabella anagrafica sedi. Il file sembra "
        "contenere solo una pivot o un riepilogo. Esporta/incolla il foglio "
        "con le colonne complete e riprova."
    )


def strip_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rimuove spazi iniziali/finali dai valori testuali delle colonne principali."""
    for column in REQUIRED_COLUMNS:
        df[column] = df[column].map(lambda value: value.strip() if isinstance(value, str) else value)

    return df


def geocode_with_retry(geolocator: GoogleV3, query: str) -> tuple[object, str]:
    """Esegue una richiesta a Google Geocoding con retry."""
    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            location = geolocator.geocode(
                query,
                exactly_one=True,
                region="it",
                timeout=10,
            )
            time.sleep(REQUEST_DELAY_SECONDS)
            return location, ""
        except (GeocoderTimedOut, GeocoderServiceError) as exc:
            last_error = str(exc)
            logging.warning(
                "Tentativo %s/%s fallito per '%s': %s",
                attempt,
                MAX_RETRIES,
                query,
                last_error,
            )
            time.sleep(REQUEST_DELAY_SECONDS * attempt)
        except GeopyError as exc:
            last_error = str(exc)
            logging.error("Errore geocoding per '%s': %s", query, last_error)
            time.sleep(REQUEST_DELAY_SECONDS)
            break

    return None, last_error


def format_status(status: str, precisione: str) -> str:
    """Rende esplicito che il risultato arriva dal provider Google."""
    return f"GOOGLE_{status}_{precisione}"


def geocode_address(geolocator: GoogleV3, row: pd.Series) -> dict:
    """Geocodifica una sede usando Google e fallback progressivi."""
    full_address = build_full_address(row)
    fallback_address = build_fallback_address(row)
    city_fallback_address = build_city_fallback_address(row)

    if not clean_value(row.get("Indirizzo")) or not clean_value(row.get("Comune")):
        return {
            "Indirizzo_Completo": full_address,
            "Latitudine": None,
            "Longitudine": None,
            "Indirizzo_Trovato": "",
            "Stato_Geocoding": "DATI_INSUFFICIENTI",
            "Provider_Geocoding": "Google",
            "Precisione_Geocoding": "NESSUNA",
        }

    location, error = geocode_with_retry(geolocator, full_address)
    if location:
        return {
            "Indirizzo_Completo": full_address,
            "Latitudine": location.latitude,
            "Longitudine": location.longitude,
            "Indirizzo_Trovato": location.address,
            "Stato_Geocoding": format_status("TROVATO", "INDIRIZZO"),
            "Provider_Geocoding": "Google",
            "Precisione_Geocoding": "INDIRIZZO",
        }

    location, fallback_error = geocode_with_retry(geolocator, fallback_address)
    if location:
        return {
            "Indirizzo_Completo": full_address,
            "Latitudine": location.latitude,
            "Longitudine": location.longitude,
            "Indirizzo_Trovato": location.address,
            "Stato_Geocoding": format_status("TROVATO", "FALLBACK"),
            "Provider_Geocoding": "Google",
            "Precisione_Geocoding": "FALLBACK",
        }

    location, city_error = geocode_with_retry(geolocator, city_fallback_address)
    if location:
        return {
            "Indirizzo_Completo": full_address,
            "Latitudine": location.latitude,
            "Longitudine": location.longitude,
            "Indirizzo_Trovato": location.address,
            "Stato_Geocoding": format_status("TROVATO", "COMUNE"),
            "Provider_Geocoding": "Google",
            "Precisione_Geocoding": "COMUNE",
        }

    status = "NON_TROVATO"
    if error or fallback_error or city_error:
        status = f"ERRORE: {city_error or fallback_error or error}"

    return {
        "Indirizzo_Completo": full_address,
        "Latitudine": None,
        "Longitudine": None,
        "Indirizzo_Trovato": "",
        "Stato_Geocoding": status,
        "Provider_Geocoding": "Google",
        "Precisione_Geocoding": "NESSUNA",
    }


def main() -> None:
    """Legge l'Excel, geocodifica con Google e salva output Excel/CSV per Power BI."""
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"File '{INPUT_FILE}' non trovato. Metti il file Excel nella cartella "
            "del progetto oppure rinominalo con questo nome."
        )

    api_key = os.getenv(API_KEY_ENV)
    if not api_key:
        raise ValueError(
            f"Variabile ambiente '{API_KEY_ENV}' non impostata. Crea una chiave "
            "Google Maps Geocoding API e impostala prima di eseguire lo script."
        )

    logging.info("Lettura file input: %s", INPUT_FILE)
    df = load_input_dataframe()
    validate_columns(df)
    df = strip_text_columns(df)

    geolocator = GoogleV3(api_key=api_key)
    cache = {}
    results = []

    total_rows = len(df)
    logging.info("Righe da elaborare: %s", total_rows)

    for index, row in df.iterrows():
        full_address = build_full_address(row)
        logging.info("Elaborazione riga %s/%s: %s", index + 1, total_rows, full_address)

        if full_address not in cache:
            cache[full_address] = geocode_address(geolocator, row)
        else:
            logging.info("Indirizzo gia' elaborato, uso risultato in cache.")

        results.append(cache[full_address])

    result_df = pd.concat([df, pd.DataFrame(results)], axis=1)

    result_df.to_excel(OUTPUT_EXCEL, index=False, engine="openpyxl")
    result_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    found_count = result_df["Latitudine"].notna().sum()
    not_found_count = result_df["Latitudine"].isna().sum()
    error_count = result_df["Stato_Geocoding"].astype(str).str.startswith("ERRORE").sum()

    print("\nRiepilogo geocoding Google")
    print("--------------------------")
    print(f"Totale righe: {total_rows}")
    print(f"Indirizzi con coordinate: {found_count}")
    print(f"Indirizzi senza coordinate: {not_found_count}")
    print(f"Errori: {error_count}")
    print(f"File Excel generato: {OUTPUT_EXCEL}")
    print(f"File CSV generato: {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
