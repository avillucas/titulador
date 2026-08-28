import os
import sys
import platform
import subprocess
from typing import List, Dict, Any, Optional

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
except ImportError:
    ctk = None

from src.models import TituloData
from src.excel_parser import ExcelParser
from src.pptx_generator import PPTXGenerator
from src.exporter import convert_pptx_to_pdf

from src.catalog import TrayectoCatalog, format_trayecto_data


DEFAULT_TEMPLATE = "ejemplos/Modelo base.pptx"
DEFAULT_EXCEL = "ejemplos/Acta de examen.xlsx"
OUTPUT_DIR = "output"

def open_in_system(path: str):
    """Opens a file or folder using the default system application."""
    if not os.path.exists(path):
        return False
    try:
        if platform.system() == "Windows":
            os.startfile(os.path.normpath(path))
        elif platform.system() == "Darwin":
            subprocess.run(["open", path], check=True)
        else:
            subprocess.run(["xdg-open", path], check=True)
        return True
    except Exception as e:
        print(f"Error opening {path}: {e}")
        return False

class TituladorGUI(ctk.CTk if ctk else object):
    def __init__(self):
        if ctk is None:
            raise RuntimeError("customtkinter is not installed.")
            
        super().__init__()

        # Window settings
        self.title("Titulador 🎓 - Generador de Certificados y Títulos A5")
        self.geometry("980x750")
        self.minsize(880, 650)

        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        # Catalog manager
        self.catalog = TrayectoCatalog()

        # State variables
        self.excel_path = ctk.StringVar(value=DEFAULT_EXCEL)
        self.template_path = ctk.StringVar(value=DEFAULT_TEMPLATE)
        self.output_dir = ctk.StringVar(value=OUTPUT_DIR)
        
        self.excel_data: Optional[Dict[str, Any]] = None
        self.current_trayecto_data: Optional[Dict[str, Any]] = None
        self.last_generated_pptx: Optional[str] = None
        self.last_generated_pdf: Optional[str] = None


        self._build_ui()
        self._load_excel_data()

    def _build_ui(self):
        # Grid layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame, 
            text="🎓 Titulador A5", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.grid(row=0, column=0, padx=15, pady=(10, 0), sticky="w")

        subtitle_label = ctk.CTkLabel(
            header_frame, 
            text="Confección y emisión de certificados de Trayectos Formativos de FP (Circular 04-2020)",
            font=ctk.CTkFont(size=12, slant="italic")
        )

        subtitle_label.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        theme_switch = ctk.CTkSwitch(
            header_frame, 
            text="Modo Oscuro", 
            command=self._toggle_theme
        )
        theme_switch.select()
        theme_switch.grid(row=0, column=1, rowspan=2, padx=15, pady=10, sticky="e")

        # 2. File Pickers & Catalog Frame
        paths_frame = ctk.CTkFrame(self, corner_radius=10)
        paths_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        paths_frame.grid_columnconfigure(1, weight=1)

        # Excel Picker
        ctk.CTkLabel(paths_frame, text="Planilla Excel:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.excel_entry = ctk.CTkEntry(paths_frame, textvariable=self.excel_path)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(paths_frame, text="Buscar...", width=80, command=self._browse_excel).grid(row=0, column=2, padx=10, pady=5)

        # Template PPTX Picker
        ctk.CTkLabel(paths_frame, text="Plantilla PPTX:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.template_entry = ctk.CTkEntry(paths_frame, textvariable=self.template_path)
        self.template_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(paths_frame, text="Buscar...", width=80, command=self._browse_template).grid(row=1, column=2, padx=10, pady=5)

        # Output Directory Picker
        ctk.CTkLabel(paths_frame, text="Carpeta Salida:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.output_entry = ctk.CTkEntry(paths_frame, textvariable=self.output_dir)
        self.output_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(paths_frame, text="Buscar...", width=80, command=self._browse_output).grid(row=2, column=2, padx=10, pady=5)

        # Output Format Selector
        ctk.CTkLabel(paths_frame, text="Formato Salida:", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.format_dropdown = ctk.CTkOptionMenu(
            paths_frame,
            values=["PPTX + PDF (Recomendado)", "Solo PPTX", "Solo PDF"]
        )

        self.format_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # Trayecto Catalog Dropdown
        ctk.CTkLabel(paths_frame, text="Trayecto Catálogo:", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, padx=10, pady=5, sticky="e")
        display_list = self.catalog.get_trayecto_display_list() or ["Sin catálogo disponible"]
        self.trayecto_dropdown = ctk.CTkOptionMenu(
            paths_frame, 
            values=display_list,
            command=self._on_trayecto_selected
        )
        self.trayecto_dropdown.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="ew")


        # 3. Main Tabview
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=2, column=0, padx=15, pady=5, sticky="nsew")

        self.tab_batch = self.tabview.add("📦 Generación Masiva")
        self.tab_select = self.tabview.add("👤 Selección Individual")
        self.tab_form = self.tabview.add("✍️ Formulario Manual")

        self._setup_batch_tab()
        self._setup_select_tab()
        self._setup_form_tab()

        # 4. Bottom Frame (Console & Actions)
        bottom_frame = ctk.CTkFrame(self, corner_radius=10)
        bottom_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(bottom_frame, height=90, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="ew")

        btn_open_out = ctk.CTkButton(
            bottom_frame, 
            text="📁 Abrir Carpeta de Salida", 
            command=self._open_output_folder,
            fg_color="#2b5b84"
        )
        btn_open_out.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        btn_open_pptx = ctk.CTkButton(
            bottom_frame, 
            text="📄 Abrir Último PPTX", 
            command=self._open_last_pptx,
            fg_color="#2b845b"
        )
        btn_open_pptx.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        btn_open_pdf = ctk.CTkButton(
            bottom_frame, 
            text="✍️ Abrir Último PDF Editable", 
            command=self._open_last_pdf,
            fg_color="#844b2b"
        )
        btn_open_pdf.grid(row=1, column=2, padx=5, pady=10, sticky="w")

        self.log("Aplicación iniciada. Catálogo JSON cargado con 322 trayectos.")


    def log(self, message: str):
        """Appends a log message to the log textbox."""
        self.log_textbox.insert("end", f"> {message}\n")
        self.log_textbox.see("end")

    def _toggle_theme(self):
        current = ctk.get_appearance_mode()
        ctk.set_appearance_mode("Light" if current == "Dark" else "Dark")

    def _browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if filename:
            self.excel_path.set(filename)
            self._load_excel_data()

    def _browse_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Plantillas PowerPoint", "*.pptx")])
        if filename:
            self.template_path.set(filename)

    def _browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir.set(folder)

    def _on_trayecto_selected(self, choice: str):
        match = self.catalog.find_by_display_name(choice)
        if match:
            self.current_trayecto_data = format_trayecto_data(match)
            self.log(f"Trayecto seleccionado del catálogo: {self.current_trayecto_data['codigo']} - {self.current_trayecto_data['nombre_trayecto']}")
            self._update_batch_info()
            self._populate_form_from_catalog()

    def _update_batch_info(self):
        if not self.excel_data:
            return
        
        egresados = self.excel_data.get("egresados", [])
        omitidos = self.excel_data.get("omitidos", [])
        especialidad_acta = self.excel_data.get("especialidad", "")
        
        if self.current_trayecto_data:
            t_str = f"Categoría JSON: [{self.current_trayecto_data['codigo']}] {self.current_trayecto_data['nombre_trayecto']} ({self.current_trayecto_data['sector']})\n" \
                    f"Horas: {self.current_trayecto_data['horas_cursada']} hs | Res: {self.current_trayecto_data['resolucion']} | Módulos: {len(self.current_trayecto_data['modulos'])}"
        else:
            t_str = "Trayecto JSON: No asignado"

        info_text = (
            f"Especialidad del Acta: {especialidad_acta}\n"
            f"{t_str}\n"
            f"Fecha Egreso: {self.excel_data.get('fecha_egreso', '')}\n"
            f"CPF N°: {self.excel_data.get('numero_cpf', '')} - {self.excel_data.get('distrito_cpf', '')}\n"
            f"Total Egresados Aprobados: {len(egresados)}\n"
            f"Total Filas Evitadas (no aprobó): {len(omitidos)}"
        )
        self.batch_info_label.configure(text=info_text)

    def _load_excel_data(self):
        path = self.excel_path.get()
        if not os.path.exists(path):
            self.log(f"Advertencia: Archivo Excel no encontrado en '{path}'")
            return
        try:
            parser = ExcelParser(path)
            self.excel_data = parser.load_data(catalog=self.catalog)
            egresados = self.excel_data["egresados"]
            omitidos = self.excel_data.get("omitidos", [])
            
            # Auto-select matched trayecto from catalog
            matched_trayecto = self.excel_data.get("trayecto_matched")
            if matched_trayecto:
                self.current_trayecto_data = self.excel_data.get("trayecto_data")
                code = matched_trayecto.get("Código", "")
                name = matched_trayecto.get("Nombre del Trayecto", "")
                sector = matched_trayecto.get("Sector", "")
                disp_str = f"{code} - {name}"
                if sector:
                    disp_str += f" [{sector}]"
                
                self.trayecto_dropdown.set(disp_str)
                self.log(f"✔ Trayecto identificado en catálogo: [{code}] {name}")
            else:
                self.log(f"⚠️ No se identificó coincidencia exacta en el catálogo para especialidad '{self.excel_data['especialidad']}'. Seleccione un trayecto manualmente.")
                self.current_trayecto_data = None

            self._update_batch_info()

            # Update Select Option Menu
            options = [f"{e['num_egresado']} - {e['apellido_nombre']} ({e['documento']})" for e in egresados]
            if options:
                self.select_dropdown.configure(values=options)
                self.select_dropdown.set(options[0])
            else:
                self.select_dropdown.configure(values=["Sin egresados aprobados"])

            self.log(f"Excel cargado: {len(egresados)} egresados aprobados encontrados.")
            if omitidos:
                self.log(f"⚠️ {len(omitidos)} fila(s) evitada(s) por no contar con Número de Egresado (no aprobó):")
                for o in omitidos:
                    self.log(f"   • [Fila {o['fila']}] {o['apellido_nombre']} (DNI: {o['documento']})")
        except Exception as e:
            self.log(f"Error al leer Excel: {e}")

    # --- TAB 1: BATCH ---
    def _setup_batch_tab(self):
        self.tab_batch.grid_columnconfigure(0, weight=1)

        info_frame = ctk.CTkFrame(self.tab_batch)
        info_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")

        self.batch_info_label = ctk.CTkLabel(
            info_frame, 
            text="Cargando información del Excel y Catálogo...", 
            justify="left",
            font=ctk.CTkFont(size=13)
        )
        self.batch_info_label.pack(padx=15, pady=15, anchor="w")

        self.progress_bar = ctk.CTkProgressBar(self.tab_batch)
        self.progress_bar.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        self.progress_bar.set(0)

        btn_batch = ctk.CTkButton(
            self.tab_batch, 
            text="🚀 Generar Todos los Certificados (PPTX)", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self._run_batch_generation
        )
        btn_batch.grid(row=2, column=0, padx=15, pady=15)

    def _run_batch_generation(self):
        excel_p = self.excel_path.get()
        tmpl_p = self.template_path.get()
        out_d = self.output_dir.get()
        fmt_choice = self.format_dropdown.get()

        if not os.path.exists(excel_p):
            messagebox.showerror("Error", "Por favor verifique que la ruta del Excel exista.")
            return

        try:
            parser = ExcelParser(excel_p)
            data = parser.load_data(catalog=self.catalog)
            egresados = data["egresados"]
            omitidos = data.get("omitidos", [])

            if omitidos:
                self.log(f"⚠️ Omite {len(omitidos)} fila(s) por no contar con Número de Egresado (no aprobó):")
                for o in omitidos:
                    self.log(f"   • [Fila {o['fila']}] {o['apellido_nombre']} (DNI: {o['documento']})")

            if not egresados:
                messagebox.showinfo("Información", "No hay egresados aprobados en la planilla.")
                return

            generator = PPTXGenerator(tmpl_p) if os.path.exists(tmpl_p) else None
            
            t_data = self.current_trayecto_data or data.get("trayecto_data")
            if t_data:
                t_line1 = t_data["titulo_linea1"]
                t_line2 = t_data["titulo_linea2"]
                horas = t_data["horas_cursada"]
                res = t_data["resolucion"]
                modules = t_data["modulos"]
            else:
                t_line1 = data["especialidad"].title()
                t_line2 = data["especialidad"].title() if len(data["especialidad"]) > 40 else ""
                horas = "540"
                res = "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"
                modules = []

            total = len(egresados)
            for idx, eg in enumerate(egresados, 1):
                filename_base = f"{eg['num_egresado']}_{eg['apellido_nombre'].replace(' ', '_')}"

                titulo_data = TituloData(
                    apellido_nombre=eg['apellido_nombre'],
                    documento=eg['documento'],
                    horas_cursada=horas,
                    titulo_linea1=t_line1,
                    titulo_linea2=t_line2,
                    resolucion=res,
                    emision_dia="02",
                    emision_mes="Septiembre",
                    emision_ano="25",
                    modulos=modules,
                    fecha_egreso=data["fecha_egreso"],
                    numero_egresado=eg["num_egresado"],
                    numero_cpf=data["numero_cpf"],
                    distrito_cpf=data["distrito_cpf"]
                )

                outs = []
                out_pptx = os.path.join(out_d, f"{filename_base}.pptx")
                generator.generate(titulo_data, out_pptx)
                self.last_generated_pptx = out_pptx

                if "PPTX" in fmt_choice:
                    outs.append("PPTX")

                if "PDF" in fmt_choice or "Recomendado" in fmt_choice:
                    out_pdf = convert_pptx_to_pdf(out_pptx, out_d)
                    if out_pdf:
                        self.last_generated_pdf = out_pdf
                        outs.append("PDF")

                self.log(f"[{idx}/{total}] Generado ({', '.join(outs)}): {filename_base}")
                self.progress_bar.set(idx / total)
                self.update_idletasks()

            messagebox.showinfo("¡Éxito!", f"Se procesaron {total} certificados ({fmt_choice}) en:\n{out_d}")

        except Exception as e:
            messagebox.showerror("Error", f"Fallo durante la generación: {e}")
            self.log(f"ERROR Batch: {e}")


    # --- TAB 2: SELECCION INDIVIDUAL ---
    def _setup_select_tab(self):
        self.tab_select.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.tab_select, 
            text="Seleccione un alumno egresado para generar su certificado PPTX:",
            font=ctk.CTkFont(size=13)
        ).pack(padx=15, pady=(20, 5), anchor="w")

        self.select_dropdown = ctk.CTkOptionMenu(self.tab_select, values=["Cargando..."], width=400)
        self.select_dropdown.pack(padx=15, pady=10, anchor="w")

        btn_single = ctk.CTkButton(
            self.tab_select, 
            text="📄 Generar PPTX del Alumno Seleccionado", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=35,
            command=self._run_single_generation
        )
        btn_single.pack(padx=15, pady=20, anchor="w")

    def _run_single_generation(self):
        selected_text = self.select_dropdown.get()
        if not selected_text or "Sin egresados" in selected_text or "Cargando" in selected_text:
            messagebox.showwarning("Atención", "Seleccione un egresado válido.")
            return

        num_egresado = selected_text.split(" - ")[0].strip()
        data = self.excel_data
        if not data:
            return

        target_eg = next((e for e in data["egresados"] if e["num_egresado"] == num_egresado), None)
        if not target_eg:
            messagebox.showerror("Error", "No se encontró la información del egresado.")
            return

        tmpl_p = self.template_path.get()
        out_d = self.output_dir.get()
        generator = PPTXGenerator(tmpl_p)

        t_data = self.current_trayecto_data or data.get("trayecto_data")
        if t_data:
            t_line1 = t_data["titulo_linea1"]
            t_line2 = t_data["titulo_linea2"]
            horas = t_data["horas_cursada"]
            res = t_data["resolucion"]
            modules = t_data["modulos"]
        else:
            t_line1 = data["especialidad"].title()
            t_line2 = data["especialidad"].title() if len(data["especialidad"]) > 40 else ""
            horas = "230"
            res = "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"
            modules = []

        titulo_data = TituloData(
            apellido_nombre=target_eg['apellido_nombre'],
            documento=target_eg['documento'],
            horas_cursada=horas,
            titulo_linea1=t_line1,
            titulo_linea2=t_line2,
            resolucion=res,
            emision_dia="02",
            emision_mes="Septiembre",
            emision_ano="25",
            modulos=modules,
            fecha_egreso=data["fecha_egreso"],
            numero_egresado=target_eg["num_egresado"],
            numero_cpf=data["numero_cpf"],
            distrito_cpf=data["distrito_cpf"]
        )

        safe_name = target_eg['apellido_nombre'].replace(" ", "_")
        out_pptx = os.path.join(out_d, f"{target_eg['num_egresado']}_{safe_name}.pptx")
        generator.generate(titulo_data, out_pptx)
        self.last_generated_pptx = out_pptx

        out_pdf = convert_pptx_to_pdf(out_pptx, out_d)
        if out_pdf:
            self.last_generated_pdf = out_pdf
            self.log(f"✔ PPTX y PDF generado: {out_pdf}")
            messagebox.showinfo("¡Éxito!", f"Certificado PPTX y PDF generados correctamente:\n{out_pptx}")
        else:
            self.log(f"✔ PPTX generado: {out_pptx}")
            messagebox.showinfo("¡Éxito!", f"Certificado PPTX generado correctamente:\n{out_pptx}")

    # --- TAB 3: FORMULARIO MANUAL ---
    def _setup_form_tab(self):
        top_frame = ctk.CTkFrame(self.tab_form)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        btn_apply_cat = ctk.CTkButton(
            top_frame,
            text="📋 Cargar campos desde el Trayecto Seleccionado del Catálogo",
            command=self._populate_form_from_catalog,
            fg_color="#1f538d"
        )
        btn_apply_cat.pack(padx=10, pady=8, anchor="w")

        scroll_frame = ctk.CTkScrollableFrame(self.tab_form)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        scroll_frame.grid_columnconfigure(1, weight=1)

        self.form_entries: Dict[str, ctk.CTkEntry] = {}

        fields = [
            ("apellido_nombre", "Apellido y Nombre:", "Pérez Juan Manuel"),
            ("documento", "DNI / Documento:", "35.140.353"),
            ("horas_cursada", "Horas de Cursada:", "230"),
            ("titulo_linea1", "Título (Línea 1):", "Operadora/or de Carpintería y"),
            ("titulo_linea2", "Título (Línea 2):", "Fabricación de Mobiliario"),
            ("resolucion", "Resolución:", "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"),
            ("emision_dia", "Emisión Día:", "02"),
            ("emision_mes", "Emisión Mes:", "Septiembre"),
            ("emision_ano", "Emisión Año (2 dígitos):", "25"),
            ("fecha_egreso", "Fecha de Egreso:", "14 de Julio de 2025"),
            ("numero_egresado", "Número de Egresado:", "354"),
            ("numero_cpf", "Número CPF:", "412"),
            ("distrito_cpf", "Distrito CPF:", "Lomas de Zamora"),
        ]

        row = 0
        for key, label_text, default_val in fields:
            ctk.CTkLabel(scroll_frame, text=label_text).grid(row=row, column=0, padx=10, pady=3, sticky="e")
            entry = ctk.CTkEntry(scroll_frame)
            entry.insert(0, default_val)
            entry.grid(row=row, column=1, padx=10, pady=3, sticky="ew")
            self.form_entries[key] = entry
            row += 1

        # Modules Sub-section
        ctk.CTkLabel(scroll_frame, text="--- Módulos (hasta 10) ---", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        self.module_entries: List[ctk.CTkEntry] = []
        for i in range(1, 11):
            ctk.CTkLabel(scroll_frame, text=f"Módulo {i}:").grid(row=row, column=0, padx=10, pady=2, sticky="e")
            mod_entry = ctk.CTkEntry(scroll_frame)
            mod_entry.grid(row=row, column=1, padx=10, pady=2, sticky="ew")
            self.module_entries.append(mod_entry)
            row += 1

        btn_form = ctk.CTkButton(
            scroll_frame, 
            text="✨ Generar Certificado (PPTX y PDF Editable)", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=35,
            command=self._run_form_generation
        )
        btn_form.grid(row=row, column=0, columnspan=2, pady=15)

        # Pre-fill form if current_trayecto_data exists
        self._populate_form_from_catalog()

    def _populate_form_from_catalog(self):
        if not hasattr(self, 'form_entries') or not self.current_trayecto_data:
            return
            
        t_data = self.current_trayecto_data
        
        self.form_entries["titulo_linea1"].delete(0, "end")
        self.form_entries["titulo_linea1"].insert(0, t_data.get("titulo_linea1", ""))

        self.form_entries["titulo_linea2"].delete(0, "end")
        self.form_entries["titulo_linea2"].insert(0, t_data.get("titulo_linea2", ""))

        self.form_entries["horas_cursada"].delete(0, "end")
        self.form_entries["horas_cursada"].insert(0, t_data.get("horas_cursada", ""))

        self.form_entries["resolucion"].delete(0, "end")
        self.form_entries["resolucion"].insert(0, t_data.get("resolucion", ""))

        # Populate modules
        mods = t_data.get("modulos", [])
        for i, entry in enumerate(self.module_entries):
            entry.delete(0, "end")
            if i < len(mods):
                entry.insert(0, mods[i])

    def _run_form_generation(self):
        tmpl_p = self.template_path.get()
        out_d = self.output_dir.get()

        if not os.path.exists(tmpl_p):
            messagebox.showerror("Error", "La plantilla PPTX no existe.")
            return

        modulos = [m.get().strip() for m in self.module_entries if m.get().strip()]

        titulo_data = TituloData(
            apellido_nombre=self.form_entries["apellido_nombre"].get(),
            documento=self.form_entries["documento"].get(),
            horas_cursada=self.form_entries["horas_cursada"].get(),
            titulo_linea1=self.form_entries["titulo_linea1"].get(),
            titulo_linea2=self.form_entries["titulo_linea2"].get(),
            resolucion=self.form_entries["resolucion"].get(),
            emision_dia=self.form_entries["emision_dia"].get(),
            emision_mes=self.form_entries["emision_mes"].get(),
            emision_ano=self.form_entries["emision_ano"].get(),
            fecha_egreso=self.form_entries["fecha_egreso"].get(),
            numero_egresado=self.form_entries["numero_egresado"].get(),
            numero_cpf=self.form_entries["numero_cpf"].get(),
            distrito_cpf=self.form_entries["distrito_cpf"].get(),
            modulos=modulos
        )

        safe_name = (titulo_data.apellido_nombre or "manual").replace(" ", "_")
        num_eg = titulo_data.numero_egresado or "000"
        out_pptx = os.path.join(out_d, f"{num_eg}_{safe_name}.pptx")

        generator = PPTXGenerator(tmpl_p)
        generator.generate(titulo_data, out_pptx)
        self.last_generated_pptx = out_pptx

        out_pdf = convert_pptx_to_pdf(out_pptx, out_d)
        if out_pdf:
            self.last_generated_pdf = out_pdf
            self.log(f"✔ PPTX y PDF generado desde formulario: {out_pptx}")
            messagebox.showinfo("¡Éxito!", f"Certificado PPTX generado correctamente:\n{out_pptx}")
        else:
            self.log(f"✔ PPTX generado desde formulario: {out_pptx}")
            messagebox.showinfo("¡Éxito!", f"Certificado PPTX generado correctamente:\n{out_pptx}")

    # --- ACTIONS ---
    def _open_output_folder(self):
        out_d = self.output_dir.get()
        os.makedirs(out_d, exist_ok=True)
        open_in_system(out_d)

    def _open_last_pptx(self):
        if self.last_generated_pptx and os.path.exists(self.last_generated_pptx):
            open_in_system(self.last_generated_pptx)
        else:
            messagebox.showinfo("Información", "Aún no se ha generado ningún archivo PPTX en esta sesión.")

    def _open_last_pdf(self):
        if self.last_generated_pdf and os.path.exists(self.last_generated_pdf):
            open_in_system(self.last_generated_pdf)
        else:
            messagebox.showinfo("Información", "Aún no se ha generado ningún archivo PDF Editable en esta sesión.")


def main():
    if ctk is None:
        print("Error: customtkinter no está instalado.")
        sys.exit(1)
    app = TituladorGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
