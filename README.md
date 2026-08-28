# Titulador 🎓 - Generador de Certificados para Centros de Formación Profesional (CFP)

**Titulador** es una herramienta de **uso libre** desarrollada en Python con **Interfaz Gráfica de Escritorio (GUI)** e **Interfaz de Línea de Comandos (CLI)** cuyo propósito es poner a disposición de los **Centros de Formación Profesional (CFP)** de la Provincia de Buenos Aires un medio automatizado, sencillo y eficiente para la confección masiva de certificados y títulos de sus egresados.

El sistema genera los certificados en formato **PowerPoint (`.pptx`) editable** y/o **PDF** para hojas de tamaño **A5** (210 mm x 148.5 mm). El formato PPTX permite a los directivos y secretarios realizar correcciones o retoques manuales posteriores en caso de ser necesario antes de la impresión definitiva sobre los formularios preimpresos.

Toda la confección se realiza en estricto cumplimiento de los lineamientos e instructivo técnico fijados por la **Circular N° 04-2020** de la Dirección de Formación Profesional (DGCyE).

---

## 📚 Fuentes y Normativa de Referencia (Archivos Modelo)

La aplicación utiliza como referencia los documentos oficiales almacenados en la carpeta `documentacion/`:

| Documento | Ruta del Archivo | Descripción |
| :--- | :--- | :--- |
| **Normativa de Certificación** | `documentacion/Circular 4-2020 Certificacion de trayectos formativos y cursos de Formacion Profesional.pdf` | Instructivo técnico oficial que establece los tamaños de fuente, alineación, estilos y pautas de testado para el frente y dorso del certificado A5. |
| **Catálogo de Trayectos** | `documentacion/CATALOGO 2025 FP.pdf` | Catálogo Oficial de Trayectos Formativos y Cursos de Formación Profesional (procesado en JSON para búsqueda automática de módulos, resoluciones y horas). |
| **Modelo Base PPTX** | `documentacion/Modelo base.pptx` | Plantilla base por defecto en PowerPoint A5 (Diapositivas 1 y 2) para la generación de certificados. |
| **Acta de Examen Modelo** | `documentacion/Acta de examen.xlsx` | Planilla Excel modelo con el formato estándar de actas de examen final de donde el sistema extrae los datos de los egresados aprobados. |

> ⚠️ **Nota aclaratoria**: Los archivos presentes en las carpetas `documentacion/` and `ejemplos/` son provistos exclusivamente a modo de **ejemplos y modelos de prueba** para verificar el correcto funcionamiento del sistema.

> 🤖 **Aviso / Disclaimer**: Esta aplicación fue creada y desarrollada con la asistencia de tecnologías de Inteligencia Artificial (IA).

---

## 🛠️ Guía de Instalación de Python y Requisitos

### 🪟 En Entornos Windows

1. **Instalar Python (versión 3.11 o superior)**:
   - Descarga el instalador oficial de Python desde [python.org/downloads](https://www.python.org/downloads/).
   - Al ejecutar el instalador, **marcar obligatoriamente la casilla**:
     `[X] Add python.exe to PATH` (Agregar Python al PATH del sistema).
   - Asegurarse de mantener activada la opción `tcl/tk and IDLE` durante la instalación (necesaria para la interfaz gráfica).
   - Verificar la instalación abriendo una terminal de **CMD** o **PowerShell** y escribiendo:
     ```cmd
     python --version
     ```

2. **Instalar LibreOffice (para exportación a PDF)**:
   - Descarga e instala LibreOffice desde [libreoffice.org](https://es.libreoffice.org/descarga/libreoffice/).
   - Permite a la aplicación convertir los certificados `.pptx` a `.pdf` de forma automática.

3. **Instalar las Dependencias del Proyecto**:
   - Abre la terminal de **CMD** o **PowerShell** en la carpeta del proyecto y ejecuta:
     ```cmd
     pip install -r requirements.txt
     ```

---

### 🐧 En Entornos Linux (Ubuntu / Debian)

1. **Instalar Python, Tkinter y LibreOffice**:
   - Abre una terminal y ejecuta:
     ```bash
     sudo apt update
     sudo apt install -y python3 python3-pip python3-tk libreoffice
     ```

2. **Instalar las Dependencias del Proyecto**:
   - En la carpeta del proyecto, ejecuta:
     ```bash
     pip install -r requirements.txt
     ```

---

## 🚀 Guía de Uso de la Aplicación

El sistema permite operar mediante **Interfaz Gráfica (GUI)** o **Línea de Comandos (CLI)** en ambos sistemas operativos.

---

### 🖥️ 1. Uso por Interfaz Gráfica de Escritorio (GUI)

La interfaz gráfica ofrece un entorno accesible con selectores de archivos y barra de progreso.

#### Lanzar la GUI:
- **En Windows (CMD / PowerShell)**:
  ```cmd
  python app_gui.py
  ```
  *(o también `python main.py gui`)*

- **En Linux**:
  ```bash
  python3 app_gui.py
  ```
  *(o `./docker-run.sh --gui` si utilizas Docker)*

#### Pasos para la Generación en la GUI:
1. **Definir el Archivo de Actas de Entrada**:
   - Haz clic en el botón **"Buscar..."** junto a *Acta Excel* y selecciona la planilla correspondiente (ej. `documentacion/Acta de examen.xlsx`).
2. **Seleccionar el Trayecto Formativo**:
   - En el desplegable *Trayecto Catálogo*, selecciona el curso correspondiente del catálogo o permite que el sistema lo detecte automáticamente desde la especialidad descrita en el acta.
3. **Seleccionar el Formato de Salida**:
   - Despliega las opciones de *Formato Salida* y elige entre:
     - `PPTX + PDF (Recomendado)`: Genera ambos formatos.
     - `Solo PPTX`: Genera únicamente los archivos de PowerPoint editables.
     - `Solo PDF`: Genera certificados listos para imprimir en PDF.
4. **Ejecutar**:
   - Haz clic en **"🚀 Generar Certificados (Lote)"**. Los archivos generados se guardarán en la carpeta `output/`.

---

### 💻 2. Uso por Línea de Comandos (CLI)

Para usuarios avanzados o integración de procesos automatizados:

#### Modo Lote / Masivo (`batch`)
Procesa el archivo de acta de examen y genera todos los certificados de los alumnos aprobados:
```bash
# En Windows:
python main.py batch --excel "documentacion/Acta de examen.xlsx" --trayecto MM11 --format all

# En Linux:
python3 main.py batch --excel "documentacion/Acta de examen.xlsx" --trayecto MM11 --format all
```

Opciones principales del comando `batch`:
- `--excel` / `-e`: Ruta al archivo Excel de Acta de Examen.
- `--trayecto` / `-k`: Código del trayecto en el catálogo (ej: `MM11`, `CL01`).
- `--format` / `-f`: Formato de salida (`pptx`, `pdf`, `all`).
- `--outdir` / `-o`: Carpeta de destino (por defecto `output/`).

#### Modo Selección Individual (`select`)
Muestra el listado de egresados en pantalla para elegir interactivamente a cuál emitirle el certificado:
```bash
python main.py select
```

#### Modo Formulario Interactivo (`form`)
Permite confeccionar un certificado cargando los datos del estudiante campo por campo desde la consola:
```bash
python main.py form
```

---

## 📂 Estructura del Proyecto

```text
titulador/
├── documentacion/                   # Fuentes de referencia y modelos oficiales
│   ├── Modelo base.pptx             # Plantilla base por defecto en PowerPoint A5
│   ├── Acta de examen.xlsx          # Planilla Excel modelo de examen
│   ├── CATALOGO 2025 FP.pdf         # Catálogo de certificaciones y módulos de FP
│   └── Circular 4-2020...pdf        # Normativa e instructivo técnico A5
├── ejemplos/
│   └── CATALOGO_2025_FP_agrupado.json # Catálogo FP estructurado para búsqueda rápida
├── output/                          # Carpeta donde se exportan los certificados generados
├── src/
│   ├── catalog.py                   # Búsqueda y normalización de trayectos del catálogo FP
│   ├── cli.py                       # Comandos CLI (batch, select, form, gui)
│   ├── excel_parser.py              # Parser y filtrado de egresados aprobados en Excel
│   ├── exporter.py                  # Conversión de PPTX a PDF vía LibreOffice
│   ├── gui_app.py                   # Aplicación de escritorio (CustomTkinter)
│   ├── models.py                    # Modelo de datos Pydantic (TituloData)
│   └── pptx_generator.py            # Motor de posicionamiento y formato en PowerPoint
├── app_gui.py                       # Lanzador directo de la GUI
├── main.py                          # Lanzador directo de la CLI
└── requirements.txt                 # Dependencias Python
```

---

## 📏 Resumen de Reglas de Confección (Circular 04-2020)

El sistema aplica automáticamente las reglas oficiales de confección para certificados A5:

- **Nombre y Apellido**: MAYÚSCULAS, fuente Arial 14 estilo Negrita.
- **Documento (D.U.)**: Formateado con puntos separadores de miles (ej: `35.140.353`).
- **Trayecto / Curso**: Arial 12, alineado exactamente tras el texto preimpreso.
- **Horas Cursada**: Arial 12, posicionado respetando la línea base preimpresa.
- **Certificación DE**: Respetando mayúsculas/minúsculas del catálogo, dividiendo en dos renglones de ser necesario hasta 202 mm de ancho.
- **Módulos (Dorso)**: Un módulo por renglón en Arial 10 con su código entre paréntesis, testando el espacio sobrante con líneas de puntos (`.....`).
