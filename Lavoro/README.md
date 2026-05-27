# Geocodifica sedi operative per Power BI

Questo progetto legge un file Excel con le sedi operative, geocodifica gli indirizzi tramite OpenStreetMap/Nominatim e genera un nuovo Excel pronto per Power BI.

## File del progetto

- `main.py`: script principale
- `requirements.txt`: dipendenze Python
- `sedi_input.xlsx`: file Excel di input da inserire nella cartella del progetto
- `sedi_geocodificate.xlsx`: file Excel generato
- `sedi_geocodificate.csv`: file CSV generato per Power BI

## Installazione

Apri il terminale in questa cartella ed esegui:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Se VS Code chiede quale interprete Python usare, seleziona quello dentro `.venv`.

## Preparazione del file Excel

Metti il file Excel nella stessa cartella di `main.py` e rinominalo:

```text
sedi_input.xlsx
```

Il file deve contenere queste colonne:

- `Tipologia`
- `STATO PDV`
- `Rag. Sociale`
- `Codice Agenzia`
- `Codici FR`
- `Indirizzo`
- `Comune`
- `CAP`
- `Provincia`
- `Regione`
- `Sede Operativa SUP`

## Esecuzione

Con l'ambiente virtuale attivo, lancia:

```powershell
python main.py
```

Lo script crea:

- `sedi_geocodificate.xlsx`
- `sedi_geocodificate.csv`

## Test con Google Geocoding

Per confrontare Nominatim/OpenStreetMap con una soluzione non open source, puoi usare lo script:

```powershell
main_google.py
```

Serve una chiave API Google Maps con Geocoding API abilitata. Imposta la chiave nella variabile ambiente:

```powershell
$env:GOOGLE_MAPS_API_KEY="LA_TUA_CHIAVE_API"
python main_google.py
```

Lo script genera file separati:

- `sedi_geocodificate_google.xlsx`
- `sedi_geocodificate_google.csv`

In questo modo puoi confrontare quanti indirizzi vengono trovati con Google rispetto a `main.py`.

## Come funziona la geocodifica

Per ogni riga viene costruito un indirizzo nel formato:

```text
{Indirizzo}, {CAP} {Comune}, {Provincia}, {Regione}, Italia
```

Esempio:

```text
Corso Torino 119, 10040 Rivarolo Canavese, TO, Piemonte, Italia
```

Se Nominatim non trova risultati, viene provato un secondo formato:

```text
{Indirizzo}, {Comune}, {Provincia}, Italia
```

Se anche questo tentativo non trova risultati, viene provato un fallback comunale:

```text
{Comune}, {Provincia}, {Regione}, Italia
```

Questo permette di ottenere coordinate approssimate per Power BI anche quando il civico o la via non vengono riconosciuti.

Lo script rispetta le policy di Nominatim usando una sola richiesta al secondo circa, un User-Agent personalizzato, retry sugli errori temporanei e una cache per evitare richieste duplicate sullo stesso indirizzo.

## Colonne aggiunte

Il file finale conserva tutte le colonne originali e aggiunge:

- `Indirizzo_Completo`
- `Latitudine`
- `Longitudine`
- `Indirizzo_Trovato`
- `Stato_Geocoding`

I possibili stati principali sono:

- `TROVATO`
- `TROVATO_FALLBACK`
- `TROVATO_COMUNE`
- `NON_TROVATO`
- `DATI_INSUFFICIENTI`
- `ERRORE: ...`

## Importazione in Power BI

In Power BI Desktop:

1. Vai su **Home > Recupera dati**.
2. Seleziona **Excel** e importa `sedi_geocodificate.xlsx`, oppure seleziona **Testo/CSV** e importa `sedi_geocodificate.csv`.
3. Carica la tabella.
4. Per una visualizzazione mappa usa:
   - `Latitudine` nel campo latitudine
   - `Longitudine` nel campo longitudine
   - `Sede Operativa SUP`, `Comune` o `Codice Agenzia` come dettaglio/tooltip

Se Power BI non riconosce subito le coordinate, imposta la categoria dati di `Latitudine` su **Latitudine** e quella di `Longitudine` su **Longitudine**.

## Note importanti

Nominatim e' un servizio pubblico gratuito: per circa 250 sedi lo script e' adatto, ma non aumentare la frequenza delle richieste. In caso di molti indirizzi non trovati, controlla la qualita' dei campi `Indirizzo`, `Comune`, `CAP` e `Provincia`.
