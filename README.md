# EasyMARC

Strumento Python per estrarre dati da file catalografici **ISO 2709** (UNIMARC / MARC21) e salvarli in un foglio di calcolo Excel (`.xlsx`), configurabile tramite un semplice file JSON.

- Python 3.8+
- Dipendenza unica: `openpyxl`
- Cross-platform: Linux, Windows, macOS
- Parser ISO 2709 built-in (nessuna libreria MARC esterna)

---

## Funzionalita principali

- Legge file `.iso` in formato ISO 2709 senza dipendenze esterne
- Estrae solo i campi che ti interessano, configurati in `config.json`
- Supporta **formati condizionali**: usa un formato diverso a seconda dei sottocampi effettivamente presenti nel record
- Supporta **valori costanti** (`"constant"`): inserisce un valore fisso uguale per tutti i record
- Supporta **estrazione dal Leader** (`"source": "leader"`): legge un carattere dalla posizione indicata e lo decodifica con una mappa testuale
- Supporta **slicing di sottocampi** (`"slice"`): estrae una sottostringa dal valore (es. l'anno da un campo data codificata)
- Supporta **filtro per valore di sottocampo** (`"filter"`): per tag ripetuti, seleziona solo le occorrenze che corrispondono a un valore specifico (es. solo la biblioteca principale)
- Gestisce **campi ripetuti** (es. piu autori, piu soggetti) unendoli con separatore configurabile
- Genera un Excel con header colorato, larghezze automatiche e riga di intestazione bloccata
- Registra ogni esecuzione in un file di log con timestamp
- Launcher `run.sh` (Linux/macOS) e `run.bat` (Windows) che creano il virtual environment automaticamente al primo avvio

---

## Installazione

**Prerequisito:** Python 3.8 o superiore installato e disponibile nel PATH.

### Metodo consigliato — launcher automatico

I launcher `run.sh` / `run.bat` creano il virtual environment e installano le dipendenze al primo avvio, senza nessun intervento manuale.

```bash
# Linux / macOS — rendi eseguibile una volta sola
chmod +x run.sh
```

### Metodo manuale

```bash
python3 -m venv .venv
source .venv/bin/activate          # Linux/macOS
# oppure: .venv\Scripts\activate   # Windows

pip install -r requirements.txt
```

---

## Uso

### Sintassi

```
./run.sh <file.iso> [--config config.json] [--output risultato.xlsx]
```

| Argomento | Obbligatorio | Default | Descrizione |
|-----------|:---:|---------|-------------|
| `file.iso` | Si | — | File UNIMARC/MARC21 in formato ISO 2709 |
| `--config` | No | `config.json` (stessa cartella dello script) | File JSON con la lista dei campi da estrarre |
| `--output` | No | Stessa cartella del `.iso`, stessa nome con estensione `.xlsx` | Percorso del file Excel di output |

### Esempi

**1. Uso base — config e output predefiniti**
```bash
./run.sh /dati/catalogo.iso
# Genera: /dati/catalogo.xlsx
```

**2. Output in una cartella diversa**
```bash
./run.sh /dati/catalogo.iso --output /export/catalogo_2026.xlsx
```

**3. Configurazione personalizzata**
```bash
./run.sh /dati/catalogo.iso --config config_soggetti.json
```

**4. Tutti gli argomenti espliciti**
```bash
./run.sh /dati/catalogo.iso --config config_custom.json --output /export/out.xlsx
```

**5. Su Windows**
```bat
run.bat "C:\dati\catalogo.iso" --output "C:\export\catalogo.xlsx"
```

### Output a console

```
INFO  Log: logs/easy_marc_20260327_143022.log
INFO  Configurazione: 21 colonne da config.json
INFO  Leggo: /dati/catalogo.iso
INFO  Record trovati: 1234
INFO  Scrivo: /dati/catalogo.xlsx
INFO  ✓ 1234 record elaborati → /dati/catalogo.xlsx
```

---

## Configurazione (config.json)

Il file `config.json` descrive le colonne da estrarre nell'Excel. Ha un'unica chiave `columns`, che contiene una lista di oggetti — uno per colonna.

### Schema

```json
{
  "columns": [
    {
      "tag":       "200",
      "label":     "Titolo",
      "formats": [
        { "subfields": ["a", "e"], "format": "{a} : {e}" },
        { "subfields": ["a"],      "format": "{a}" }
      ],
      "separator": " | "
    }
  ]
}
```

### Campi dell'oggetto colonna

Ogni oggetto nella lista `columns` puo avere queste proprieta:

| Campo | Obbligatorio | Default | Descrizione |
|-------|:---:|---------|-------------|
| `label` | Si | — | Intestazione della colonna nell'Excel |
| `tag` | Condizionale | — | Codice tag UNIMARC/MARC21 (es. `"001"`, `"200"`). Obbligatorio per colonne da tag; assente per `constant` e `source` |
| `formats` | No | Concatenazione automatica | Lista ordinata di formati da provare (vedi sotto) |
| `separator` | No | `" \| "` | Separatore tra piu occorrenze dello stesso tag |
| `join` | No | — | Dentro un formato: unisce i valori ripetuti di un sottocampo con il separatore indicato (vedi sotto) |
| `filter` | No | — | Filtra le occorrenze del tag per valore di un sottocampo (vedi sotto) |
| `constant` | — | — | Inserisce un valore fisso, uguale per tutti i record. Alternativo a `tag` |
| `source` | — | — | `"leader"`: legge un carattere dal Leader del record. Usare insieme a `offset` e `map` |
| `offset` | No | — | Posizione (0-based) nel Leader da leggere. Usato solo con `"source": "leader"` |
| `map` | No | — | Dizionario di decodifica per il valore letto dal Leader (es. `"a"` → `"TESTO A STAMPA"`) |

### Formati condizionali

Ogni elemento di `formats` definisce un template da usare **solo se tutti i sottocampi indicati sono presenti** nel campo. I formati vengono provati nell'ordine: viene usato il primo che corrisponde.

```json
"formats": [
  { "subfields": ["a", "e"], "format": "{a} : {e}" },
  { "subfields": ["a"],      "format": "{a}" }
]
```

Per un record con `$a Fondamenti di Python $e Guida pratica`:
- Il primo formato corrisponde → output: `Fondamenti di Python : Guida pratica`

Per un record con solo `$a Introduzione alla logica`:
- Il primo formato non corrisponde (manca `$e`)
- Il secondo corrisponde → output: `Introduzione alla logica`

Se nessun formato corrisponde, la cella rimane vuota.

> **Nota:** I tag `001`–`009` sono campi di controllo (senza indicatori ne sottocampi): il loro valore viene estratto direttamente, indipendentemente da `formats`.

#### Join di sottocampi ripetuti

Per default, se un sottocampo compare più volte nello stesso campo (es. più `$e` nel tag 200), viene usato solo il primo valore. Aggiungendo `"join"` al formato è possibile raccogliere **tutti** i valori e unirli con un separatore personalizzato.

```json
{
  "tag": "200",
  "label": "Titolo",
  "formats": [
    { "subfields": ["a", "e"], "format": "{a} : {e}", "join": {"e": " : "} },
    { "subfields": ["a"],      "format": "{a}" }
  ]
}
```

`"join": {"e": " : "}` significa: unisci tutte le occorrenze di `$e` con il separatore ` : ` invece di usarne solo una.

Esempio: `$a Storia dell'arte $e Pittura : Scultura`:
- Senza `join`: `Storia dell'arte : Pittura`
- Con `join: {"e": " : "}`: `Storia dell'arte : Pittura : Scultura`

#### Slicing (sottostringa)

Aggiungendo `"slice": [inizio, fine]` a un formato, il risultato viene tagliato a quella sottostringa (stessa sintassi di Python). Utile per estrarre porzioni da campi a posizione fissa come il tag `100` di UNIMARC, dove la data di pubblicazione inizia al byte 9.

```json
{ "subfields": ["a"], "format": "{a}", "slice": [9, 13] }
```

Esempio: `$a 1984    19840000` → slice [9,13] → `1984`

### Filtro occorrenze (`filter`)

Quando un tag compare piu volte nello stesso record (es. tag `957` con esemplari di biblioteche diverse), `filter` seleziona solo le occorrenze dove un dato sottocampo ha un valore specifico.

```json
{
  "tag": "957",
  "label": "managementId",
  "filter": { "subfield": "3", "value": "ABR0ME" },
  "formats": [
    { "subfields": ["3", "c", "d"], "format": "{3}-{c}-{d}" }
  ]
}
```

In questo esempio vengono estratti solo gli esemplari di `ABR0ME`, ignorando tutte le altre biblioteche presenti nel campo `957`.

### Valori costanti (`constant`)

Per colonne che devono contenere lo stesso valore su ogni riga (es. l'identificativo dell'ente conservatore), si usa `constant` al posto di `tag`:

```json
{ "constant": "IT-CH0020", "label": "conservativeId" }
```

### Lettura dal Leader (`source: leader`)

Il Leader e il blocco iniziale di 24 caratteri di ogni record MARC. Per decodificare la posizione 6 (tipo di risorsa) con una mappa testuale:

```json
{
  "source": "leader",
  "offset": 6,
  "label": "typeOfResource",
  "map": {
    "a": "TESTO A STAMPA",
    "b": "TESTO MANOSCRITTO",
    "e": "MATERIALE CARTOGRAFICO A STAMPA",
    "g": "PROIEZIONI"
  }
}
```

Se il carattere letto non e presente nella `map`, viene restituito il carattere grezzo.

### Campi ripetuti

Se un tag compare piu volte nello stesso record (es. piu autori in tag `700`), i valori vengono uniti con il `separator`. Esempio con `"separator": " | "`:

```
Rossi, Mario | Verdi, Anna | Bianchi, Luigi
```

### File di configurazione inclusi

Il repository include due file di configurazione pronti all'uso:

| File | Cliente / Uso |
|------|--------------|
| `config-EasyCat-Chieti-periodici.json` | EasyCat, Chieti — periodici. Usa tag `957` con `filter` per biblioteca `ABR0ME`, `conservativeId = IT-CH0020` |
| `config-SBNcloud-Foligno-Loreti.json` | SBNcloud, Foligno — biblioteca Loreti. Usa tag `950` per collocazioni, `conservativeId = IT-PG0035`, `join` su `$e` e `$a` del tag 200 |

---

### Configurazione di default (SBN/UNIMARC)

Il file `config.json` fornito e precompilato per cataloghi SBN. Contiene 21 colonne:

| # | Label | Sorgente | Formato / Note |
|---|-------|----------|----------------|
| 1 | logicalId | Tag 001 | Identificativo SBN (valore diretto) |
| 2 | Collocazioni | Tag 957 | `{3}-{c}-{d}` con fallback — tutte le biblioteche (ripetibile) |
| 3 | managementId | Tag 957 | `{3}-{c}-{d}` — solo biblioteca con `$3 = ABR0ME` |
| 4 | conservativeId | Costante | `IT-CH0020` su ogni riga |
| 5 | Titolo | Tag 200 | `{a} : {e}` oppure `{a}` |
| 6 | Luogo | Tag 210 | `{a}` |
| 7 | Editore | Tag 210 | `{c}` |
| 8 | Lingua | Tag 101 | `{a}` |
| 9 | Paese | Tag 102 | `{a}` |
| 10 | Tipo | Tag 110 | `{a}` |
| 11 | typeOfResource | Leader pos. 6 | Decodificato con mappa (es. `a` → `TESTO A STAMPA`) |
| 12 | title | Tag 200 | `{a} : {e}` oppure `{a}` |
| 13 | dateIssued_start | Tag 100 | `$a` slice [9:13] — anno di inizio |
| 14 | dateIssued_end | Tag 100 | `$a` slice [13:17] — anno di fine |
| 15 | form | Costante | `n.d.` su ogni riga |
| 16 | extent | Tag 215 | `{a} : {c} ; {d}` con fallback |
| 17 | Soggetto | Tag 610 | `{a}` (ripetibile con ` \| `) |
| 18 | Autore principale | Tag 700 | `{a}, {b}` oppure `{a}` (ripetibile) |
| 19 | Altro responsabile | Tag 702 | `{a}, {b}` oppure `{a}` (ripetibile) |
| 20 | Origine catalogazione | Tag 801 | `{a}/{b} ({c})` (ripetibile) |
| 21 | Collocazione | Tag 852 | `{a} - {j}` con fallback (ripetibile) |

---

## Output Excel

Il file Excel generato contiene un unico foglio denominato **UNIMARC**.

- **Riga 1 — Intestazioni:** sfondo blu scuro (`#1F497D`), testo bianco in grassetto, bloccata durante lo scorrimento
- **Righe 2 in poi — Dati:** un record per riga, font Calibri 10pt
- **Larghezza colonne:** calcolata automaticamente dal contenuto (massimo 62 caratteri)
- **Intestazione bloccata:** la riga 1 rimane visibile anche scorrendo verso il basso

---

## Formato ISO 2709

ISO 2709 e lo standard internazionale (ISO/IEC 2709) per lo scambio di record bibliografici. E il formato di trasporto usato da tutti i principali sistemi bibliotecari (SBN, OPAC, ILS) per esportare e importare cataloghi.

Ogni file `.iso` contiene una sequenza di record. Ogni record e strutturato in tre parti:
- **Leader** (24 byte): metadati sul record (lunghezza, tipo, posizione dei dati)
- **Directory**: indice dei campi presenti (tag, lunghezza, posizione)
- **Campi dati**: valori veri e propri, con indicatori e sottocampi separati da caratteri di controllo

EasyMARC implementa il parser internamente, senza librerie esterne, ed e compatibile con file UNIMARC, MARC21 e le varianti SBN.

---

## Log

Ogni esecuzione produce un file di log in `logs/` con nome `easy_marc_YYYYMMDD_HHMMSS.log`.

```
logs/
└── easy_marc_20260327_143022.log
```

I messaggi di errore piu comuni:

| Messaggio | Causa |
|-----------|-------|
| `File non trovato: <file.iso>` | Il percorso del file ISO e errato |
| `File di configurazione non trovato: <config>` | Il `config.json` non esiste nel percorso indicato |
| `Errore nel file di configurazione: ...` | Il JSON non e valido o manca la chiave `columns` |
| `Record troppo corto, saltato` | Record corrotto nel file ISO (warning, elaborazione continua) |

---

## Struttura del progetto

```
EasyMARC/
├── easy_marc.py       # Script principale
├── config.json        # Configurazione campi da estrarre
├── requirements.txt   # Dipendenze Python (openpyxl>=3.1)
├── run.sh             # Launcher Linux/macOS
├── run.bat            # Launcher Windows
├── logs/              # File di log (creata automaticamente)
│   └── easy_marc_YYYYMMDD_HHMMSS.log
└── README.md          # Questo file
```
