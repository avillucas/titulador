# Documentación Operativa y Guía de Agente: Titulador Desktop & Docker

Este documento sirve como guía operativa para desarrolladores y agentes de IA encargados del mantenimiento, compilación y despliegue de la aplicación **Titulador**.

---

## 1. Arquitectura General

La aplicación **Titulador** está diseñada con una arquitectura modular desacoplada:

- **`src/excel_parser.py`**: Procesa libros de trabajo Excel (`.xlsx`), valida la estructura de actas de examen de titulación e extrae la nómina de alumnos.
- **`src/pptx_generator.py`**: Generador de documentos `.pptx` (formato A5) basado en plantillas PowerPoint.
- **`src/html_generator.py`**: Nuevo motor central de generación en **HTML + CSS** A5 (210mm x 148.5mm) y conversión directa a **PDF** vía Chrome Headless.
- **`templates/template_certificate.html`**: Plantilla HTML5/CSS3 con posicionamiento estricto de campos según la Circular 04-2020.
- **`templates/assets/`**: Almacena las imágenes de fondo en alta resolución de la plantilla (`frente.jpg` y `dorso.jpg`).
- **`src/cli.py` / `main.py`**: Interfaz de línea de comandos (CLI) construida sobre `typer` y `rich` (soporta `--format html|pdf|pptx|all`).
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

El proceso de confección de certificados cumple estrictamente con el **ANEXO INSTRUCTIVO TÉCNICO** de la **Circular N° 04-2020** de la Dirección General de Cultura y Educación (DGCyE).

*Nota importante: Se toma la información del D.U. y del Catálogo de Certificaciones, descartando que el acta contenga los números a los que hace referencia.*

---FRENTE DEL CERTIFICADO:

1. **POR CUANTO - (Nombre y Apellido)**
   - **a)** Tomar los datos del D.U.
   - **b)** Completar el nombre tal cual figura en el D.U.
   - **c)** Completar el apellido tal cual figura en el D.U.
   - **d)** Texto todo en **MAYÚSCULAS**, formato de fuente **Arial 14**, estilo **negrita**.
   - **e)** No superponer el texto con el renglón pre-impreso.
   - **f)** No extenderse más allá del final del renglón pre-impreso.

2. **NÚMERO DE D.U.**
   - **a)** Tomar la fotocopia del D.U. y transcribirlo tal cual figura separando por puntos las unidades (ej: `35.140.353`).
   - **b)** Formato de fuente **Arial 12**.
   - **c)** No superponer el texto con el renglón pre-impreso.
   - **d)** No extenderse más allá del final del renglón pre-impreso.
   - **e)** Si tuviera una letra, consignarla.

3. **TRAYECTO / CURSO FORMACIÓN PROFESIONAL DE (Nombre del Trayecto / Curso)**
   - **a)** Tomar del Catálogo de certificaciones.
   - **b)** Formato de fuente **Arial 12**.
   - **c)** Completar el nombre tal cual figura en el Catálogo respetando las letras mayúsculas, minúsculas, puntos y abreviaturas que figuran en la columna *"Nombre del Trayecto"*.
   - **d)** No superponer el texto con el renglón pre-impreso.
   - **e)** No extenderse más allá del final del renglón pre-impreso.
   - **f)** Para utilizar el segundo renglón se debe utilizar el primero por completo, pudiendo separar en sílabas la última palabra, siempre que sea ortográficamente posible.
   - **g)** Testar con línea de puntos el excedente de espacio no utilizado.

4. **CANTIDAD DE HORAS**
   - **a)** Tomar del Catálogo de certificaciones.
   - **b)** Completar la cantidad de horas reloj asignada al Trayecto/Curso en el catálogo de certificaciones (columna *"Hs. Reloj"*).
   - **c)** **Excepción Reglamentaria**: Las certificaciones de *Gasista de 3ra Categoría*, *Gasista de 2da Categoría*, *Montador Electricista* y *Electricista Instalador* se deberán consignar en **horas cátedra** para dar cumplimiento a las Resoluciones 3993/11 y 2265/01 respectivamente.
   - **d)** Formato de fuente **Arial 12**.

5. **CERTIFICACIÓN DE (Denominación de la certificación)**
   - **a)** Tomar del Catálogo de certificaciones.
   - **b)** Formato de fuente **Arial 12**.
   - **c)** Completar el nombre tal cual figura en el Catálogo respetando las letras mayúsculas, minúsculas, puntos y abreviaturas que figuran en la columna *"Certificación"*.
   - **d)** No superponer el texto con el renglón pre-impreso.
   - **e)** No extenderse más allá del final del renglón pre-impreso.
   - **f)** Para utilizar el segundo renglón se debe utilizar el primero por completo, pudiendo separar en sílabas la última palabra, siempre que sea ortográficamente posible.
   - **g)** Testar con línea de puntos el excedente de espacio no utilizado.

6. **FECHA**
   - **a)** Completar con el día en que se confecciona el certificado (nunca puede ser anterior a la fecha de egreso).
   - **b)** Usar el N° de día, nombre de mes y últimas dos cifras del año completando lo pre-impreso en el certificado.
   - **c)** Ej. `"01 de Julio 2018"`.
   - **d)** Formato de fuente **Arial 12**.

7. **FIRMAS**
   - **a)** En el renglón sin texto pre-impreso debe firmar el Director o quien se encuentre a cargo de la dirección previo acto administrativo correspondiente que acredite dicha situación.
   - **b)** La firma del Director debe ir acompañada en el lugar correspondiente pre-impreso y destinado a tal fin con la firma del inspector del servicio.

---DORSO / ANVERSO DEL CERTIFICADO:

1. **MÓDULOS**
   - **a)** Tomar del Catálogo de certificaciones.
   - **b)** Completar el nombre tal cual figura en el Catálogo respetando las letras mayúsculas, minúsculas, puntos y abreviaturas que figuran en la columna *"Denominación del Módulo"*.
   - **c)** No superponer el texto con el renglón pre-impreso.
   - **d)** No extenderse más allá del final del renglón pre-impreso.
   - **e)** Ubicar un módulo por renglón.
   - **f)** Por cada renglón, de ser necesario, se puede disponer de 2 niveles de escritura; al finalizar el texto consignar el código del módulo entre paréntesis.
   - **g)** Testar con línea de puntos el excedente de espacio.
   - **h)** Formato de fuente **Arial 10**.

2. **FECHA DE EGRESO**
   - **a)** Tomar del acta de examen.
   - **b)** Completar con el día que se asignó al examen final.
   - **c)** Usar el N° de día con 2 cifras, nombre del mes completo en formato texto y 4 cifras para el año (Ej. `"29 de Noviembre 2017"`).
   - **d)** Formato de fuente **Arial 12**.

3. **Nº DE EGRESADO**
   - **a)** Asignar los números de egresados a los alumnos que figuran en el acta de examen como aprobados.
   - **b)** Alumno que adeuda documentación no se le asignará el número de egresado. Una vez completa la misma, se asignará un número pudiendo este no ser correlativo al Trayecto / Curso.
   - **c)** Completar el certificado con el número asignado.
   - **d)** Formato de fuente **Arial 12**.

4. **C.F.P. Nº**
   - **a)** Consignar las 3 cifras del número que identifica al servicio en el distrito.
   - **b)** Formato de fuente **Arial 12**.

5. **DISTRITO**
   - **a)** Completar el nombre del distrito completo sin abreviaturas (Ej. `"Lomas de Zamora"`).
   - **b)** Formato de fuente **Arial 12**.

6. **FIRMAS (Anverso)**
   - **a)** Los datos consignados los puede firmar el secretario/a o quien esté a cargo de la secretaría previo acto administrativo correspondiente que acredite dicha situación. De no existir el cargo, estar descubierto o por razones operativas del servicio, debe firmar el Director o quien esté a cargo de la Dirección previo acto administrativo correspondiente.
   - **b)** La autenticación de la firma que antecede la debe hacer una autoridad superior (si firmase el secretario será autenticada por el Director; si firma el Director solo podrá ser autenticada por el inspector a cargo del servicio).
   - **c)** Para mayor seguridad y evitar duplicaciones, las firmas se realizarán en **tinta azul**.

---

## 7. Diagnóstico y Registro de Errores (Logs) en Windows

Cuando el ejecutable `Titulador.exe` se empaqueta en modo GUI con PyInstaller (`--windowed` / `console=False`), la consola del sistema operativo (STDERR/STDOUT) queda deshabilitada. Si ocurre una excepción o fallo de inicio (por falta de assets, DLLs o dependencias), el programa se cierra inmediatamente sin mostrar ninguna ventana de error.

### 7.1. Captura de Errores Integrada (Log & Popup)
El archivo `app_gui.py` incluye un manejador global `sys.excepthook` que:
1. Genera un archivo **`titulador_error.log`** en la misma carpeta donde reside `Titulador.exe` con la traza completa de la excepción (`traceback`).
2. Muestra una ventana de diálogo de error de Windows (`tkinter.messagebox`) indicando el fallo visualmente antes de cerrar.

### 7.2. Compilación en Modo Consola para Depuración (Debug Mode)
Si se desea ver la salida de terminal en vivo al ejecutar desde `cmd.exe` o `PowerShell` en Windows:
1. Reemplazar `--windowed` por `--console` en `Dockerfile.windows` (o quitar `console=False` en `Titulador.spec`).
2. Ejecutar `./windows_build.sh`.
3. Al ejecutar `.\Titulador.exe` desde la terminal de Windows, se mantendrá una consola visible mostrando logs en tiempo real y errores de importación/runtime.

### 7.3. Visor de Eventos de Windows (Event Viewer)
Si la aplicación ni siquiera inicia la máquina virtual de Python (crash nativo o falta de DLLs como `vcruntime140.dll` / `tcl86t.dll`):
1. Presionar `Win + R`, escribir `eventvwr.msc` y presionar Enter.
2. Navegar a **Registros de Windows** -> **Aplicación**.
3. Buscar eventos con nivel **Error** con origen `Application Error` asociados a `Titulador.exe` para identificar la DLL o librería faltante.


