#!/usr/bin/env bash
set -e

# Build docker image if needed
echo "🐳 Construyendo/Verificando imagen Docker 'titulador:latest'..."
docker build -t titulador:latest .

if [ "$1" == "--gui" ]; then
    echo "🖥️  Iniciando Titulador GUI en contenedor Docker con reenvío X11..."
    xhost +local:docker 2>/dev/null || true
    docker run -it --rm \
      -e DISPLAY="${DISPLAY}" \
      -v /tmp/.X11-unix:/tmp/.X11-unix \
      -v "$(pwd)/ejemplos:/app/ejemplos" \
      -v "$(pwd)/output:/app/output" \
      titulador:latest python app_gui.py
else
    echo "🚀 Ejecutando Titulador CLI en Docker..."
    docker run -it --rm \
      -v "$(pwd)/ejemplos:/app/ejemplos" \
      -v "$(pwd)/output:/app/output" \
      titulador:latest python main.py "$@"
fi
