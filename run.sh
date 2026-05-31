#!/usr/bin/env bash
# Avvia easy_marc.py nel venv locale.
# Uso: ./run.sh <file.iso> [--config config.json] [--output out.xlsx]

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$SCRIPT_DIR/.venv"

# Crea il venv se non esiste
if [ ! -d "$VENV" ]; then
    echo "Creo l'ambiente virtuale..."
    python3 -m venv "$VENV"
    "$VENV/bin/pip" install --quiet -r "$SCRIPT_DIR/requirements.txt"
    echo "Ambiente pronto."
fi

"$VENV/bin/python" "$SCRIPT_DIR/easy_marc.py" "$@"
