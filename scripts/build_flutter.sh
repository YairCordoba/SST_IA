#!/usr/bin/env bash
set -e
# Cambiar al directorio del proyecto Flutter (un nivel arriba de scripts)
cd "$(dirname "$0")/../flutter"
flutter clean
flutter pub get
flutter build web --release
echo "✅ Flutter web compiled → flutter/build/web"