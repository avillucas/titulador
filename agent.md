# Documentación Operativa y Guía de Agente: Titulador Desktop & Docker

Este documento sirve como guía operativa para desarrolladores y agentes de IA encargados del mantenimiento, compilación y despliegue de la aplicación **Titulador**.

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
- **`Dockerfile`**: Imagen ligera basada en `python:3.11-slim` que incluye LibreOffice, paquetes de Tkinter (`python3-tk`, `tk`, `tcl`) y fuentes de sistema.
  > **Nota de Configuración**: Se utiliza `CMD ["python", "main.py", "batch"]` en lugar de `ENTRYPOINT` para permitir a Docker y Docker Compose sobrescribir libremente el comando al invocar `python app_gui.py` sin conflictos con el parser CLI de `main.py`.
- **`docker-compose.yml`**:
  - Servicio `titulador`: Ejecuta procesos batch en modo CLI (`python main.py batch`).
  - Servicio `titulador-gui`: Permite lanzar la interfaz de escritorio GUI (`python app_gui.py`) mediante mapeo de socket X11 (`/tmp/.X11-unix`) y la variable `DISPLAY`.

#### Ejecutar GUI en Docker:
```bash
./docker-run.sh --gui
# O directamente vía docker compose:
xhost +local:docker 2>/dev/null || true
docker compose run --rm titulador-gui
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
| `Dockerfile` | Definición del contenedor de ejecución Linux (CLI/GUI). |
| `docker-compose.windows.yml` | Composición de servicios de compilación. |
| `docker-compose.yml` | Composición de servicios de ejecución (titulador / titulador-gui). |

---

## 6. Normativa e Instructivo Técnico (Circular 04-2020)

El proceso de confección de certificados cumple estrictamente con el **ANEXO INSTRUCTIVO TÉCNICO** de la **Circular N° 04-2020** de la Dirección de Formación Profesional (DGCyE):

1. **Nombre y Apellido (Frente)**: Debe completarse en **MAYÚSCULAS** (fuente Arial 14 negrita).
2. **D.U. / Documento (Frente)**: Transcrito separando las unidades por puntos (ej. `35.140.353`), fuente Arial 12.
3. **Nombre del Trayecto / Curso (Frente)**: Extraído de la columna *"Nombre del Trayecto"* del Catálogo JSON, respetando mayúsculas, minúsculas, puntos y abreviaturas oficiales (fuente Arial 12).
4. **Cantidad de Horas (Frente)**:
   - Se consigna en **Horas Reloj** por defecto.
   - **Excepción Reglamentaria (Sección 4.c Anexo)**: Para las certificaciones de *Gasista de 3ra Categoría*, *Gasista de 2da Categoría*, *Montador Electricista* y *Electricista Instalador*, se consigna obligatoriamente en **Horas Cátedra** en cumplimiento de las Res. 3993/11 y 2265/01.
5. **Certificación DE (Frente)**: Extraído de la columna *"Certificación"* del Catálogo (fuente Arial 12).
6. **Módulos (Anverso)**: Formato `Denominación del Módulo. (CÓDIGO)` (fuente Arial 10, un módulo por renglón). Renglones excedentes testados con línea de puntos (`---`).
7. **Fecha de Egreso y N° de Egresado (Anverso)**: Fecha en formato texto completo (ej. `14 de Julio de 2025`) y filtrado estricto para procesar únicamente alumnos aprobados con número de egresado asignado.

