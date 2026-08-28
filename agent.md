# Documentación Operativa y Guía de Agente: Titulador Desktop & CLI

Este documento sirve como guía operativa para desarrolladores y agentes de IA encargados del mantenimiento y ejecución de la aplicación **Titulador**.

### Finalidad Principal del Proyecto
La finalidad principal de la aplicación **Titulador** es poner a disposición un medio de uso libre para la generación automatizada de los certificados oficiales de **Certificación de Trayectos Formativos y Cursos de Formación Profesional** de los distintos CFP de la Provincia de Buenos Aires, dando cumplimiento estricto a las especificaciones y normativas dictadas en el anexo técnico de la **Circular N° 04-2020** (`documentacion/Circular 4-2020 Certificacion de trayectos formativos y cursos de Formacion Profesional.pdf`).


---


## 1. Arquitectura General

La aplicación **Titulador** está diseñada con una arquitectura modular desacoplada:

- **`src/excel_parser.py`**: Procesa libros de trabajo Excel (`.xlsx`), valida la estructura de actas de examen de titulación y extrae la nómina de alumnos aprobados (omitinedo aquellos sin Número de Egresado).
- **`src/catalog.py`**: Carga y gestiona el catálogo oficial `CATALOGO_2025_FP_agrupado.json`, realizando el emparejamiento automático por especialidad y extrayendo código, horas, resolución y lista de módulos.
- **`src/pptx_generator.py`**: Motor central de generación de documentos `.pptx` (formato A5) basado en plantillas PowerPoint.
- **`src/exporter.py`**: Gestiona la conversión de PPTX a PDF mediante LibreOffice (`convert_pptx_to_pdf`) e impresiones.
- **`src/cli.py` / `main.py`**: Interfaz de línea de comandos (CLI) construida sobre `typer` y `rich`.
- **`src/gui_app.py` / `app_gui.py`**: Interfaz gráfica de escritorio (GUI) moderna construida con `customtkinter`.

---

## 2. Configuración e Integración con Docker (Linux / Servidores)

- **`Dockerfile`**: Imagen basada en `python:3.11-slim` que incluye LibreOffice, paquetes de Tkinter (`python3-tk`, `tk`, `tcl`) y fuentes de sistema.
- **`docker-compose.yml`**:
  - Servicio `titulador`: Ejecuta procesos batch en modo CLI (`python main.py batch`).
  - Servicio `titulador-gui`: Permite lanzar la interfaz de escritorio GUI (`python app_gui.py`) mediante mapeo de socket X11 (`/tmp/.X11-unix`) y la variable `DISPLAY`.

#### Ejecutar GUI en Docker:
```bash
./docker-run.sh --gui
```

#### Ejecutar CLI en Docker:
```bash
./docker-run.sh batch --excel ejemplos/Acta\ de\ examen.xlsx --outdir output/
```

---

## 3. Guía de Ejecución Nativa en Windows

Para simplificar el uso y evitar fallos de librerías congeladas de PyInstaller/Wine, la aplicación se ejecuta directamente con Python nativo en Windows:

1. **Instalar Python 3.11+**:
   - Asegurarse de marcar `[X] Add python.exe to PATH`.
2. **Instalar LibreOffice**:
   - Necesario para que `convert_pptx_to_pdf` pueda generar los PDFs automáticamente.
3. **Instalar dependencias**:
   ```cmd
   pip install -r requirements.txt
   ```
4. **Lanzar la GUI**:
   ```cmd
   python app_gui.py
   ```

---

## 4. Matriz de Archivos del Proyecto

| Archivo | Descripción |
| :--- | :--- |
| `app_gui.py` | Lanzador principal de la GUI desktop. |
| `main.py` | Lanzador principal de la CLI. |
| `src/excel_parser.py` | Parser de actas de examen en Excel. |
| `src/catalog.py` | Gestor del catálogo oficial JSON de FP. |
| `src/pptx_generator.py` | Motor de llenado y formateo de plantillas PPTX. |
| `src/exporter.py` | Conversor de PPTX a PDF vía LibreOffice. |
| `src/gui_app.py` | Componentes y vistas de la interfaz CustomTkinter. |
| `src/cli.py` | Comandos CLI (batch, select, form). |
| `docker-run.sh` | Orchestrador de ejecución Docker (CLI/GUI). |
| `Dockerfile` | Definición del contenedor de ejecución Linux (CLI/GUI). |
| `docker-compose.yml` | Composición de servicios de ejecución. |

---

## 5. Normativa e Instructivo Técnico (Circular 04-2020)

El proceso de confección de certificados cumple estrictamente con el **ANEXO INSTRUCTIVO TÉCNICO** de la **Circular N° 04-2020** de la Dirección General de Cultura y Educación (DGCyE):

- **Nombre y Apellido**: MAYÚSCULAS, Arial 14 Negrita.
- **Documento**: Puntos de separación por unidades (ej: `35.140.353`).
- **Horas**: Arial 12, posicionado tras `"de "` sin superposición.
- **Título / Certificación**: Arial 12, renglón 1 alineado tras `"CERTIFICADO DE "` desde 114 mm hasta 202 mm (`max_line1 = 32`).
- **Módulos (Dorso)**: Formateados con código entre paréntesis al final y testados con líneas de puntos (`.....`).
