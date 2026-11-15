#!/usr/bin/env bash
set -e

# Compila Flutter (usar ruta del script)
"$(dirname "$0")/build_flutter.sh"

# Prepara el backend Python
cd "$(dirname "$0")/../backend"
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

echo "✅ Backend deps installed (local)"
