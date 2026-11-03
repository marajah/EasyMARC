# EasyMARC

Script Python per estrarre dati da file catalografici **ISO 2709** (UNIMARC/MARC21) e salvarli in un foglio Excel (`.xlsx`), configurabile tramite JSON.

## Quick Start

```bash
# Linux/macOS — crea il venv e installa le dipendenze al primo avvio
chmod +x run.sh
./run.sh <file.iso> [--config config.json] [--output out.xlsx]

# Windows
run.bat "C:\dati\catalogo.iso"
```

## Ambiente

- **Python:** 3.8+
- **Dipendenza principale:** `openpyxl`
- **Venv:** `.venv/` (escluso da git)

## Struttura

```
EasyMARC/
├── easy_marc.py                        # Script principale (unico file)
├── config-EasyCat-Chieti-periodici.json   # Config per EasyCat Chieti (periodici)
├── config-SBNcloud-Foligno-Loreti.json    # Config per SBNcloud Foligno (biblioteca Loreti)
├── requirements.txt
├── run.sh / run.bat                    # Launcher cross-platform
└── logs/                               # Log automatici (easy_marc_YYYYMMDD_HHMMSS.log)
```

## Architettura di easy_marc.py

| Sezione | Funzione | Cosa fa |
|---------|----------|---------|
| Parser | `parse_iso2709()` | Legge il file `.iso` byte per byte, estrae leader, directory e campi |
| Parser | `_parse_field()` | Decodifica indicatori e sottocampi (US = `0x1F`) di un singolo campo dati |
| Parser | `_sanitize()` | Rimuove escape ISO 2022 e caratteri di controllo non validi in Excel |
| Estrazione | `apply_column()` | Applica una colonna della config a un record: gestisce `constant`, `source: leader`, tag normali, filtro e ripetizioni |
| Estrazione | `_apply_formats()` | Prova i formati condizionali nell'ordine, usa il primo con tutti i sottocampi presenti |
| Estrazione | `_apply_auto()` | Fallback: concatena tutti i sottocampi come `$x valore` |
| Output | `write_excel()` | Scrive il file `.xlsx` con header colorato (blu `#1F497D`), larghezze auto, riga bloccata |
| CLI | `main()` | Argparse: `iso_file`, `--config`, `--output` |

## Schema config.json

```json
{
  "columns": [
    {
      "tag": "200",
      "label": "Titolo",
      "formats": [
        { "subfields": ["a", "e"], "format": "{a} : {e}", "join": {"e": " : "} },
        { "subfields": ["a"],      "format": "{a}" }
      ],
      "separator": " | "
    }
  ]
}
```

### Chiavi per ogni colonna

| Chiave | Descrizione |
|--------|-------------|
| `tag` | Tag UNIMARC/MARC21 (es. `"200"`) |
| `label` | Intestazione colonna nell'Excel |
| `formats` | Lista di formati condizionali (vedi sotto) |
| `separator` | Separatore tra occorrenze ripetute (default `" \| "`) |
| `filter` | Seleziona solo occorrenze dove `subfield` == `value` (es. solo una biblioteca) |
| `constant` | Valore fisso uguale per tutti i record (alternativo a `tag`) |
| `source: "leader"` | Legge un carattere dal Leader alla posizione `offset`, decodificato con `map` |

### Formati condizionali

Ogni formato in `formats` ha:
- `subfields`: lista dei sottocampi richiesti (tutti devono essere presenti)
- `format`: template con placeholder `{a}`, `{b}`, ecc.
- `join` *(opzionale)*: per sottocampi ripetuti, unisce tutti i valori col separatore indicato invece di usare solo il primo. Es. `{"e": " : "}` → tutti i `$e` uniti da ` : `
- `slice` *(opzionale)*: `[inizio, fine]` sulla stringa risultante (es. anno da campo 100)

### Fogli aggiuntivi (`sheets`, opzionale)

Chiave top-level accanto a `columns`. Ogni voce crea un foglio con **una riga per ogni occorrenza** di `tag$subfield` (es. tutti i `950$e` di un periodico). Se la chiave manca, l'output è identico a prima (solo foglio `UNIMARC`).

| Chiave | Default | Descrizione |
|--------|---------|-------------|
| `name` | `<tag><subfield>` | Nome del foglio (max 31 caratteri) |
| `tag`, `subfield` | — (obbligatori) | Sottocampo da espandere |
| `id_tag` | `"001"` | Campo di controllo usato come id del record |
| `key_regex` | `(\\d{4})\\s*$` (anno finale) | Regex sul valore a spazi compressi; il gruppo 1 è la chiave. Alias storico: `year_regex`. Es. `^\\S+ (\\d+)` = inventario |
| `key_zfill` | `4` | Se la chiave è numerica, zeri a sinistra fino a N cifre (mai troncata): `34` → `0034` |
| `combo_separator` | `"_"` | Separatore tra id e chiave |
| `pad_to` | — | Completa il valore con spazi finali fino a N caratteri (mai troncato; warning se più lungo) |
| `filter` | — | Come nelle colonne: `{"subfield": "3", "value": "..."}` |
| `labels` | `logicalId` / `managementId` / `sheet_key` | Intestazioni (`id`, `value`, `combo`) |

Colonne: id del record, valore con spazi compressi (`SV 33 5 1927`), `id_chiave` (`CFI0350986_0034` con key_regex inventario, `CFI0350986_1928` col default anno). Se la chiave non si trova, la cella combo resta vuota e viene loggato un warning.

## File di configurazione presenti

| File | Cliente / Uso |
|------|--------------|
| `config-EasyCat-Chieti-periodici.json` | EasyCat, Chieti — periodici. Usa tag `957` con `filter` per biblioteca `ABR0ME`, `conservativeId = IT-CH0020` |
| `config-SBNcloud-Savona-Letimbro.json` | SBNcloud, Savona — Biblioteca Barrili (periodici). Copia della config Imperia + foglio `managementId` da `950$e` (valore a 15 caratteri, chiave `logicalId_inventario` a 4 cifre) |
| `config-SBNcloud-Foligno-Loreti.json` | SBNcloud, Foligno — biblioteca Loreti. Usa tag `950` per collocazioni, `conservativeId = IT-PG0035`, `join` su sottocampi `$e` e `$a` del 200 |

## Cronologia modifiche rilevanti

- **Commit `223e5a5`** — Aggiunto `"join"` nei format spec per unire sottocampi ripetuti (es. più `$e` nel tag 200). I file di config sono stati rinominati con il nome cliente/istituto. Prima il primo sottocampo ripetuto veniva silenziosamente scartato.
- **Commit `7e87a5c`** — Commit iniziale: parser ISO 2709 built-in, estrazione Excel, formati condizionali, `constant`, `source: leader`, `filter`, `slice`, launcher cross-platform.

## Git

- **Branch principale:** `main`
- **Remote:** da configurare (repo indipendente da XMLData)
- Questo repo è nested dentro `scripts/` del progetto XMLData — tenerli separati
