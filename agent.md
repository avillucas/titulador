# Documentación Operativa y Guía de Agente: Titulador Desktop & Docker

Este documento sirva como guía operativa para desarrolladores y agentes de IA encargados del mantenimiento, compilación y despliegue de la aplicación **Titulador**.

---

## 1. Arquitectura General

La aplicación **Titulador** está diseñada con una arquitectura modular desacoplada:

- **`src/excel_parser.py`**: Procesa libros de trabajo Excel (`.xlsx`), valida la estructura de actas de examen de titulación e extrae la nómina de alumnos.
- **`src/pptx_generator.py`**: Motor central de generación de documentos `.pptx` (formato A5) basado en plantillas personalizadas.
- **`src/cli.py` / `main.py`**: Interfaz de línea de comandos (CLI) construida sobre `typer` y `rich`.
- **`src/gui_app.py` / `app_gui.py`**: Interfaz gráfica de escritorio (GUI) moderna construida con `customtkinter`.

---

## 2. Configuración e Integración con Docker

El proyecto cuenta con dos entornos de Docker aislados para distintos objetivos:

### 2.1. Entorno de Ejecución Local / Servidor (Linux CLI & GUI)
- **`Dockerfile`**: Imagen ligera basada en `python:3.11-slim` que incluye LibreOffice e impresoras CUPS.
- **`docker-compose.yml`**:
  - Servicio `titulador`: Ejecuta procesos batch en modo CLI.
  - Servicio `titulador-gui`: Permite lanzar la interfaz de escritorio GUI mediante mapeo de socket X11 (`/tmp/.X11-unix`).

#### Ejecutar GUI en Docker:
```bash
./docker-run.sh --gui
```

#### Ejecutar CLI en Docker:
```bash
./docker-run.sh batch --acta ejemplos/acta.xlsx --output-dir output/
```

---

## 3. Pipeline de Compilación Cruzada para Windows (`.exe`)

El ejecutable para Windows se compila desde Linux sin requerir una máquina virtual Windows dedicada, utilizando un contenedor con Wine + Python 3.11 oficial + PyInstaller.

### 3.1. Archivos involucrados
- **`Dockerfile.windows`**: Entorno basado en Ubuntu 22.04, instala Wine 64-bit y desempaqueta el entorno Nuget de Python 3.11 (que incluye soporte nativo Tcl/Tk y `tkinter` necesario para CustomTkinter).
- **`docker-compose.windows.yml`**: Define la orquestación del build y monta la carpeta `./dist` para extraer el `.exe`.
- **`windows_build.sh`**: Script principal que dispara el proceso completo.

### 3.2. Ejecutar la compilación localmente
```bash
./windows_build.sh
```
El archivo resultante estará disponible en `./dist/Titulador.exe`.

---

## 4. Comandos SSH para Compilación y Ejecución Remota

A continuación se detallan los comandos SSH para operar la aplicación desde un servidor remoto:

### 4.1. Disparar Compilación de Windows vía SSH
Para ejecutar el build del ejecutable `.exe` en el servidor remoto:
```bash
ssh usuario@servidor "cd /ruta/a/titulador && ./windows_build.sh"
```

### 4.2. Descargar el Ejecutable `.exe` Generado vía SCP
Una vez finalizada la compilación en el servidor remoto, se puede copiar el ejecutable a la máquina local:
```bash
scp usuario@servidor:/ruta/a/titulador/dist/Titulador.exe ./dist/Titulador.exe
```

### 4.3. Lanzar la Interfaz GUI vía SSH con X11 Forwarding
Para ver y controlar la GUI ejecutándose en el servidor desde la máquina local:
```bash
ssh -X usuario@servidor "cd /ruta/a/titulador && ./docker-run.sh --gui"
```

### 4.4. Ejecutar Proceso Batch vía SSH
Para procesar un acta de examen en segundo plano vía SSH:
```bash
ssh usuario@servidor "cd /ruta/a/titulador && docker compose run --rm titulador batch --acta ejemplos/acta.xlsx"
```

---

## 5. Matriz de Archivos y Responsabilidades

| Archivo | Descripción |
| :--- | :--- |
| `app_gui.py` | Lanzador principal de la GUI desktop. |
| `main.py` | Lanzador principal de la CLI. |
| `windows_build.sh` | Orchestrador de build para Windows `.exe`. |
| `docker-run.sh` | Orchestrador de ejecución Docker (CLI/GUI). |
| `Dockerfile.windows` | Definición de compilación cruzada Wine + Python 3.11. |
| `Dockerfile` | Definición del contenedor de ejecución Linux. |
| `docker-compose.windows.yml` | Composición de servicios de compilación. |
| `docker-compose.yml` | Composición de servicios de ejecución. |
