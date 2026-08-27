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

## 📜 Reglas de Confección (Circular 04-2020 - Frente del Certificado)

El sistema aplica las reglas oficiales de confección para el **Frente del Certificado** (descartando la referencia a números del acta):

### 📄 Frente del Certificado

1. **Nombre y Apellido (POR CUANTO)**:
   - Datos tomados del D.U. (se descarta que el acta contenga los números a los que hace referencia).
   - Completar nombre y apellido tal cual figuran en el D.U.
   - Texto en **MAYÚSCULAS**, fuente **Arial 14**, estilo **negrita**.
   - No superponer con el renglón pre-impreso ni extenderse más allá de su final.

2. **NÚMERO DE D.U.**:
   - Transcripción tal cual figura en el D.U. separando unidades por puntos (ej: `35.140.353`). Consignar letra si la tuviera.
   - Fuente **Arial 12**.
   - Respetar los límites del renglón pre-impreso.

3. **TRAYECTO / CURSO FORMACIÓN PROFESIONAL DE**:
   - Datos del Catálogo de Certificaciones (columna *"Nombre del Trayecto"*).
   - Fuente **Arial 12**. Respetar exactamente mayúsculas, minúsculas, puntos y abreviaturas del catálogo.
   - Para usar el segundo renglón se completa el primero y se puede separar en sílabas la última palabra.
   - Testar con línea de puntos el excedente de espacio no utilizado.

4. **CANTIDAD DE HORAS**:
   - Horas reloj asignadas en el Catálogo de Certificaciones (columna *"Hs. Reloj"*).
   - **Excepción**: Certificaciones de *Gasista de 3ra Categoría*, *Gasista de 2da Categoría*, *Montador Electricista* y *Electricista Instalador* se consignan en **horas cátedra** (Res. 3993/11 y 2265/01).
   - Fuente **Arial 12**.

5. **CERTIFICACIÓN DE**:
   - Datos del Catálogo de Certificaciones (columna *"Certificación"*).
   - Fuente **Arial 12**. Respetar mayúsculas, minúsculas, puntos y abreviaturas del catálogo.
   - Completar renglones en orden, separando en sílabas la última palabra si es necesario. Testar excedente con línea de puntos.

6. **FECHA**:
   - Día de confección del certificado (nunca anterior a la fecha de egreso).
   - Formato: Nº de día, nombre del mes y últimas dos cifras del año (ej. `"01 de Julio 2018"`).
   - Fuente **Arial 12**.

7. **FIRMAS (Frente)**:
   - Renglón sin texto pre-impreso destinado a la firma del Director (o a cargo de dirección).
   - Firma del Director acompañada de la firma del Inspector del servicio en el lugar correspondiente.

### 📄 Anverso / Reverso del Certificado

1. **MÓDULOS**:
   - Tomar del Catálogo de certificaciones (columna *"Denominación del Módulo"*), respetando mayúsculas, minúsculas, puntos y abreviaturas.
   - Ubicar un módulo por renglón (hasta 2 niveles de escritura si es necesario) y consignar el código entre paréntesis al final.
   - Fuente **Arial 10**. Testar excedente con línea de puntos. No superponer ni extender del renglón.

2. **FECHA DE EGRESO**:
   - Tomada del acta de examen (día del examen final).
   - Formato: N° de día (2 cifras), nombre de mes completo y 4 cifras del año (ej. `"29 de Noviembre 2017"`). Fuente **Arial 12**.

3. **Nº DE EGRESADO**:
   - Asignar a alumnos aprobados en el acta de examen. Alumno que adeuda documentación no recibe número hasta completarla.
   - Completar en el certificado con fuente **Arial 12**.

4. **C.F.P. Nº**:
   - Consignar las 3 cifras del número de servicio en el distrito. Fuente **Arial 12**.

5. **DISTRITO**:
   - Nombre completo del distrito sin abreviaturas (ej. `"Lomas de Zamora"`). Fuente **Arial 12**.

6. **FIRMAS (Anverso)**:
   - Firmado por Secretario/a (o a cargo) o Director.
   - Autenticación por autoridad superior (Director autentica a Secretario, Inspector autentica a Director).
   - Firmas obligatoriamente en **tinta azul**.

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
