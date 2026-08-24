#!/usr/bin/env bash
set -e

echo "🚀 Iniciando compilación de la app de escritorio de Windows (.exe) mediante Docker..."

# Asegurar permisos de ejecución y directorio dist
mkdir -p dist

# Ejecutar docker compose build y run
docker compose -f docker-compose.windows.yml build
docker compose -f docker-compose.windows.yml run --rm windows_builder

echo "=========================================================="
echo "✔ ¡Proceso completado exitosamente!"
echo "📂 El ejecutable listo para Windows está en: ./dist/Titulador.exe"
echo "=========================================================="
