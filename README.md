# Titulador 🎓 - Sistema de Generación de Certificados y Títulos (A5)

**Titulador** es una herramienta en Python (CLI y **Aplicación de Escritorio GUI**) con soporte de **Docker** para automatizar la generación, formateo y exportación de títulos y certificados oficiales en formato **PPTX editable** para hojas **A5** (210 mm x 148.5 mm).

Toma datos directamente desde planillas de examen en **Excel (`Acta de examen.xlsx`)** y completa dinámicamente plantillas personalizables de **PowerPoint (`Modelo base.pptx`)**.

---

## 📌 Características Principales

- 💻 **Aplicación de Escritorio GUI (CustomTkinter)**: Interfaz gráfica moderna con tema oscuro/claro para Windows, Linux y macOS.
- 📊 **Carga de Datos Automática desde Excel**: Extrae especialidades, fechas de egreso, número de egresado, número de CPF, distrito y datos personales de los alumnos (con formateo automático de DNI).
- 📄 **Salida Editable en PowerPoint (`.pptx`)**: Genera archivos PPTX en tamaño A5 listos para abrir y editar manualmente en Microsoft PowerPoint si fuera necesario.
- 📝 **Rellenado Inteligente de Módulos**: Asigna las materias/módulos en la Diapositiva 2 y completa automáticamente los casilleros no utilizados con guiones discontinuos (`----------------------------------`).
- ⚡ **3 Modos de Operación**:
  - **`batch` (Generación Masiva)**: Procesamiento automático de todos los egresados aprobados del acta con barra de progreso.
  - **`select` (Selección Individual)**: Listado de egresados para elegir interactiva y puntualmente a quién emitirle el certificado.
  - **`form` (Formulario Manual)**: Formulario paso a paso para carga rápida sin necesidad de archivo Excel.
- 🐳 **Compilación Docker para Windows (`.exe`)**: Script automatizado `./windows_build.sh` que genera un archivo ejecutable `.exe` en el directorio `dist/` mediante Docker sin necesidad de estar en Windows.

---

## 🛠️ Requisitos Previos

### Opción 1: Ejecución Local en Python

- **Python**: 3.10 o superior.
- **Dependencias**: Instalar paquetes requeridos:
  ```bash
  pip install -r requirements.txt
  ```

### Opción 2: Compilación / Ejecución con Docker
Solo requiere tener instalado **Docker** y **Docker Compose**.

---

## 🚀 Guía de Uso

### 🖥️ Aplicación de Escritorio (GUI)

#### Ejecución Local:
```bash
python app_gui.py
```

#### Ejecución con Docker (Linux con X11):
```bash
./docker-run.sh --gui
```
*(Permite seleccionar archivos Excel/PPTX mediante selectores de archivos, procesar lotes, cargar formularios manuales y abrir los PPTX generados directamente con un clic).*

---

### 💻 Uso de la Línea de Comandos (CLI)

#### 1. Modo Lote / Masivo (`batch`)
Procesa el archivo Excel y genera los certificados de **todos los egresados aprobados**:
```bash
python main.py batch
# O mediante Docker:
./docker-run.sh batch
```

#### 2. Modo Selección (`select`)
Permite elegir mediante su índice cuál certificado procesar:
```bash
python main.py select
```

#### 3. Modo Formulario Interactivo (`form`)
Ingreso manual a través de preguntas en la consola:
```bash
python main.py form
```

---

## 📦 Compilación del Ejecutable de Windows (`.exe`) con Docker

Para compilar un paquete ejecutable standalone `.exe` para Windows desde cualquier sistema operativo (Linux/macOS) mediante Docker:

```bash
# Otorgar permisos de ejecución (solo la primera vez)
chmod +x windows_build.sh

# Compilar el ejecutable
./windows_build.sh
```

Al finalizar el proceso, el archivo ejecutable resultante estará disponible en la carpeta de volumen:
```text
dist/Titulador.exe
```

---

## 🔑 Comandos SSH para Compilación y Ejecución Remota

Para compilar o ejecutar el proyecto en un servidor remoto mediante SSH:

1. **Compilar `.exe` de Windows remotamente vía SSH:**
   ```bash
   ssh usuario@servidor "cd /ruta/a/titulador && ./windows_build.sh"
   ```

2. **Descargar el ejecutable `.exe` compilado vía SCP:**
   ```bash
   scp usuario@servidor:/ruta/a/titulador/dist/Titulador.exe ./dist/Titulador.exe
   ```

3. **Lanzar la GUI remota en tu pantalla local (X11 Forwarding):**
   ```bash
   ssh -X usuario@servidor "cd /ruta/a/titulador && ./docker-run.sh --gui"
   ```

*(Para más detalles técnicos de arquitectura y despliegue, consultar [`agent.md`](agent.md)).*

---

## 📂 Estructura del Proyecto

```text
titulador/
├── ejemplos/
│   ├── Acta de examen.xlsx    # Archivo Excel modelo con egresados
│   └── Modelo base.pptx       # Plantilla base en formato A5 (Diapositivas 1 y 2)
├── output/                    # Carpeta donde se exportan los PPTX generados
├── dist/                      # Carpeta donde se genera el ejecutable Titulador.exe para Windows
├── src/
│   ├── cli.py                 # Definición de comandos Typer (batch, select, form)
│   ├── excel_parser.py        # Extracción y formateo de datos desde Excel
│   ├── gui_app.py             # Aplicación de escritorio CustomTkinter
│   ├── models.py              # Esquema de datos Pydantic (TituloData)
│   └── pptx_generator.py      # Manipulación de formas y textos en PowerPoint
├── app_gui.py                 # Punto de entrada de la aplicación de escritorio GUI
├── agent.md                   # Documentación técnica de arquitectura y SSH
├── Dockerfile                 # Configuración de la imagen Docker estándar
├── Dockerfile.windows         # Configuración de compilación cruzada Wine/Python 3.11/PyInstaller
├── docker-compose.yml         # Orquestación Docker CLI y GUI
├── docker-compose.windows.yml # Orquestación Docker para build de Windows
├── docker-run.sh              # Bash script ejecutor para Docker (CLI / GUI)
├── windows_build.sh           # Bash script ejecutor para compilar dist/Titulador.exe
├── main.py                    # Punto de entrada de la aplicación CLI
└── requirements.txt           # Dependencias de Python
```

---

## 📐 Mapeo de Formas de la Plantilla PowerPoint (`Modelo base.pptx`)

### 📄 Diapositiva 1 (Frente)
| Campo | ID Forma PPTX | Descripción / Ejemplo |
| :--- | :--- | :--- |
| **Apellido y Nombre** | `85` | Nombre completo del alumno |
| **Documento / DNI** | `86` | DNI formateado con puntos (Ej: `35.140.353`) |
| **Horas de cursada** | `87` | Duración del curso (Ej: `230`) |
| **Nombre del Título (Línea 1)** | `88` | Nombre de la especialidad |
| **Nombre del Título (Línea 2)** | `89` | Segunda línea para nombres largos |
| **Resolución** | `90` | Texto de resolución oficial |
| **Día Emisión** | `91` | Día de emisión (Ej: `02`) |
| **Mes Emisión** | `92` | Mes de emisión (Ej: `Septiembre`) |
| **Año Emisión** | `93` | Año en dos dígitos (Ej: `25`) |

### 📄 Diapositiva 2 (Dorso)
| Campo | ID Forma PPTX | Descripción / Ejemplo |
| :--- | :--- | :--- |
| **Módulos 1 al 10** | `99` al `108` | Lista de materias con código asignado |
| **Fecha de Egreso** | `109` | Fecha oficial (Ej: `14 de Julio de 2025`) |
| **Número de Egresado** | `110` | N° correlativo de egresado (Ej: `354`) |
| **Número CPF** | `111` | Identificador del CFP (Ej: `412`) |
| **Distrito CPF** | `112` | Ubicación geográfica (Ej: `Lomas de Zamora`) |
