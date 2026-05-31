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

## File di configurazione presenti

| File | Cliente / Uso |
|------|--------------|
| `config-EasyCat-Chieti-periodici.json` | EasyCat, Chieti — periodici. Usa tag `957` con `filter` per biblioteca `ABR0ME`, `conservativeId = IT-CH0020` |
| `config-SBNcloud-Foligno-Loreti.json` | SBNcloud, Foligno — biblioteca Loreti. Usa tag `950` per collocazioni, `conservativeId = IT-PG0035`, `join` su sottocampi `$e` e `$a` del 200 |

## Cronologia modifiche rilevanti

- **Commit `223e5a5`** — Aggiunto `"join"` nei format spec per unire sottocampi ripetuti (es. più `$e` nel tag 200). I file di config sono stati rinominati con il nome cliente/istituto. Prima il primo sottocampo ripetuto veniva silenziosamente scartato.
- **Commit `7e87a5c`** — Commit iniziale: parser ISO 2709 built-in, estrazione Excel, formati condizionali, `constant`, `source: leader`, `filter`, `slice`, launcher cross-platform.

## Git

- **Branch principale:** `main`
- **Remote:** da configurare (repo indipendente da XMLData)
- Questo repo è nested dentro `scripts/` del progetto XMLData — tenerli separati
