import os
import sys
from typing import Optional, List, Dict, Any
import typer
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm
from rich.panel import Panel

from src.models import TituloData
from src.excel_parser import ExcelParser
from src.pptx_generator import PPTXGenerator
from src.exporter import convert_pptx_to_pdf, generate_editable_pdf, print_pdf_a5


from src.catalog import TrayectoCatalog, format_trayecto_data


app = typer.Typer(name="titulador", help="Sistema de generación de certificados y títulos en A5")
console = Console()

DEFAULT_TEMPLATE = "ejemplos/Modelo base.pptx"
DEFAULT_EXCEL = "ejemplos/Acta de examen.xlsx"
OUTPUT_DIR = "output"

def resolve_trayecto(catalog: TrayectoCatalog, code_or_name: Optional[str], excel_trayecto_data: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Resolves trayecto data from CLI option or Excel auto-matched trayecto."""
    if code_or_name:
        match = catalog.find_by_code(code_or_name) or catalog.find_by_display_name(code_or_name) or catalog.search_trayecto(code_or_name)
        if match:
            return format_trayecto_data(match)
        else:
            console.print(f"[bold yellow]⚠️  No se encontró el trayecto '{code_or_name}' en el catálogo JSON.[/bold yellow]")
    
    if excel_trayecto_data:
        return excel_trayecto_data
        
    return None

@app.command("batch")
def batch_generate(
    excel_path: str = typer.Option(DEFAULT_EXCEL, "--excel", "-e", help="Ruta al archivo Excel de Acta de examen"),
    template_path: str = typer.Option(DEFAULT_TEMPLATE, "--template", "-t", help="Ruta a la plantilla PPTX"),
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o", help="Carpeta de salida"),
    trayecto_code: Optional[str] = typer.Option(None, "--trayecto", "-k", help="Código o nombre del trayecto en el catálogo (ej. MM11)"),
    convert_pdf: bool = typer.Option(True, "--pdf/--no-pdf", help="Convertir a PDF automáticamente si LibreOffice está disponible")
):
    """Procesa un archivo Excel de Acta de examen y genera los certificados de todos los egresados aprobados."""
    console.print(Panel("[bold green]Titulador - Generación en Lote (Batch)[/bold green]"))

    if not os.path.exists(excel_path):
        console.print(f"[bold red]Error:[bold red] El archivo Excel no existe: {excel_path}")
        raise typer.Exit(code=1)

    if not os.path.exists(template_path):
        console.print(f"[bold red]Error:[bold red] La plantilla PPTX no existe: {template_path}")
        raise typer.Exit(code=1)

    catalog = TrayectoCatalog()
    parser = ExcelParser(excel_path)
    data = parser.load_data(catalog=catalog)
    
    egresados = data["egresados"]
    omitidos = data.get("omitidos", [])

    trayecto_data = resolve_trayecto(catalog, trayecto_code, data.get("trayecto_data"))

    console.print(f"Especialidad Acta: [cyan]{data['especialidad']}[/cyan]")
    if trayecto_data:
        console.print(f"[bold green]✔ Trayecto Catálogo JSON:[/bold green] [cyan]{trayecto_data['codigo']}[/cyan] - [cyan]{trayecto_data['nombre_trayecto']}[/cyan] ({trayecto_data['sector']})")
        console.print(f"  • Horas: [yellow]{trayecto_data['horas_cursada']}[/yellow] | Res: [yellow]{trayecto_data['resolucion']}[/yellow] | Módulos: [yellow]{len(trayecto_data['modulos'])}[/yellow]")
    else:
        console.print("[bold yellow]⚠️  No se identificó trayecto en catálogo JSON. Se usarán datos por defecto del acta.[/bold yellow]")

    console.print(f"Fecha Egreso: [cyan]{data['fecha_egreso']}[/cyan]")
    console.print(f"Total egresados aprobados: [bold yellow]{len(egresados)}[/bold yellow]")

    if omitidos:
        console.print(f"Total filas evitadas (no aprobó): [bold red]{len(omitidos)}[/bold red]\n")
        console.print("[bold yellow]⚠️  Filas evitadas por no contar con Número de Egresado (no aprobó):[/bold yellow]")
        for o in omitidos:
            console.print(f"  • [yellow]{o['apellido_nombre']}[/yellow] (DNI: {o['documento']}) - Fila {o['fila']}")
        console.print("")
    else:
        console.print("")

    generator = PPTXGenerator(template_path)
    
    # Determine certificate template fields from trayecto_data or fallback
    if trayecto_data:
        t_line1 = trayecto_data["titulo_linea1"]
        t_line2 = trayecto_data["titulo_linea2"]
        horas = trayecto_data["horas_cursada"]
        res = trayecto_data["resolucion"]
        modules = trayecto_data["modulos"]
    else:
        t_line1 = data["especialidad"].title()
        t_line2 = data["especialidad"].title() if len(data["especialidad"]) > 40 else ""
        horas = "540"
        res = "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"
        modules = []

    table = Table(title="Certificados a Generar")
    table.add_column("N° Egresado", style="cyan")
    table.add_column("Nombre y Apellido", style="bold white")
    table.add_column("DNI", style="yellow")
    table.add_column("Archivo PPTX", style="green")

    generated_files = []

    for eg in egresados:
        filename_base = f"{eg['num_egresado']}_{eg['apellido_nombre'].replace(' ', '_')}"
        out_pptx = os.path.join(output_dir, f"{filename_base}.pptx")
        
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

        generator.generate(titulo_data, out_pptx)
        
        pdf_file = None
        if convert_pdf:
            pdf_file = convert_pptx_to_pdf(out_pptx, output_dir)
            
        generated_files.append((out_pptx, pdf_file))
        table.add_row(eg["num_egresado"], eg["apellido_nombre"], eg["documento"], os.path.basename(out_pptx))

    console.print(table)
    console.print(f"\n[bold green]✔ ¡Se generaron {len(generated_files)} certificados PPTX exitosamente en {output_dir}![/bold green]")


@app.command("form")
def interactive_form(
    template_path: str = typer.Option(DEFAULT_TEMPLATE, "--template", "-t", help="Ruta a la plantilla PPTX"),
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o", help="Carpeta de salida"),
    trayecto_code: Optional[str] = typer.Option(None, "--trayecto", "-k", help="Código o nombre del trayecto en el catálogo (ej. MM11)")
):
    """Crea un certificado completando el formulario campo por campo en la consola."""
    console.print(Panel("[bold green]Titulador - Formulario Interactivo[/bold green]"))

    generator = PPTXGenerator(template_path)
    catalog = TrayectoCatalog()

    selected_code = trayecto_code
    if not selected_code:
        selected_code = Prompt.ask("Código o nombre de trayecto en catálogo (ej. MM11, presione ENTER para omitir)", default="")

    trayecto_data = None
    if selected_code:
        match = catalog.find_by_code(selected_code) or catalog.find_by_display_name(selected_code) or catalog.search_trayecto(selected_code)
        if match:
            trayecto_data = format_trayecto_data(match)
            console.print(f"[bold green]✔ Carga desde catálogo:[/bold green] {trayecto_data['codigo']} - {trayecto_data['nombre_trayecto']}")
        else:
            console.print(f"[bold yellow]⚠️  Trayecto '{selected_code}' no encontrado en el catálogo.[/bold yellow]")

    def_l1 = trayecto_data["titulo_linea1"] if trayecto_data else "Operadora/or de Carpintería y Fabricación de Mobiliario"
    def_l2 = trayecto_data["titulo_linea2"] if trayecto_data else ""
    def_hrs = trayecto_data["horas_cursada"] if trayecto_data else "230"
    def_res = trayecto_data["resolucion"] if trayecto_data else "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"
    def_mods = trayecto_data["modulos"] if trayecto_data else []

    apellido_nombre = Prompt.ask("Apellido y Nombre")
    documento = Prompt.ask("DNI / Documento")
    horas_cursada = Prompt.ask("Cantidad de horas de cursada", default=def_hrs)
    titulo_linea1 = Prompt.ask("Nombre del Título (Línea 1)", default=def_l1)
    titulo_linea2 = Prompt.ask("Nombre del Título (Línea 2)", default=def_l2)
    resolucion = Prompt.ask("Resolución", default=def_res)
    emision_dia = Prompt.ask("Emisión Día", default="02")
    emision_mes = Prompt.ask("Emisión Mes", default="Septiembre")
    emision_ano = Prompt.ask("Emisión Año", default="25")
    
    fecha_egreso = Prompt.ask("Fecha de Egreso", default="14 de Julio de 2025")
    numero_egresado = Prompt.ask("Número de Egresado", default="354")
    numero_cpf = Prompt.ask("Número CPF", default="412")
    distrito_cpf = Prompt.ask("Distrito CPF", default="Lomas de Zamora")

    if def_mods:
        console.print("\n[bold yellow]Módulos cargados automáticamente desde el catálogo:[/bold yellow]")
        for i, m in enumerate(def_mods, 1):
            console.print(f"  {i}. {m}")
        if not Confirm.ask("¿Desea conservar estos módulos del catálogo?", default=True):
            def_mods = []

    modulos = def_mods
    if not modulos:
        console.print("\n[bold yellow]Ingrese los módulos (hasta 10, presione ENTER para finalizar):[/bold yellow]")
        modulos = []
        for i in range(1, 11):
            mod = Prompt.ask(f"Módulo {i}", default="")
            if not mod:
                break
            modulos.append(mod)

    data = TituloData(
        apellido_nombre=apellido_nombre,
        documento=documento,
        horas_cursada=horas_cursada,
        titulo_linea1=titulo_linea1,
        titulo_linea2=titulo_linea2,
        resolucion=resolucion,
        emision_dia=emision_dia,
        emision_mes=emision_mes,
        emision_ano=emision_ano,
        modulos=modulos,
        fecha_egreso=fecha_egreso,
        numero_egresado=numero_egresado,
        numero_cpf=numero_cpf,
        distrito_cpf=distrito_cpf
    )

    safe_name = apellido_nombre.strip().replace(" ", "_")
    output_pptx = os.path.join(output_dir, f"{numero_egresado}_{safe_name}.pptx")
    generator.generate(data, output_pptx)

    console.print(f"\n[bold green]✔ Certificado PPTX generado en: {output_pptx}[/bold green]")
    
    pdf_path = convert_pptx_to_pdf(output_pptx, output_dir)
    if pdf_path:
        console.print(f"[bold cyan]✔ Versión PDF generada en: {pdf_path}[/bold cyan]")
        if Confirm.ask("¿Desea enviar a imprimir en A5?"):
            if print_pdf_a5(pdf_path):
                console.print("[bold green]✔ Enviado a la impresora en A5.[/bold green]")
            else:
                console.print("[bold red]No se pudo imprimir automáticamente.[/bold red]")

@app.command("select")
def select_egresado(
    excel_path: str = typer.Option(DEFAULT_EXCEL, "--excel", "-e"),
    template_path: str = typer.Option(DEFAULT_TEMPLATE, "--template", "-t"),
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o"),
    trayecto_code: Optional[str] = typer.Option(None, "--trayecto", "-k", help="Código o nombre del trayecto en el catálogo (ej. MM11)")
):
    """Muestra la lista de egresados del Excel y permite elegir uno para emitir su certificado."""
    if not os.path.exists(excel_path):
        console.print(f"[bold red]Error:[bold red] El archivo Excel no existe: {excel_path}")
        raise typer.Exit(code=1)

    catalog = TrayectoCatalog()
    parser = ExcelParser(excel_path)
    data = parser.load_data(catalog=catalog)
    egresados = data["egresados"]
    omitidos = data.get("omitidos", [])

    trayecto_data = resolve_trayecto(catalog, trayecto_code, data.get("trayecto_data"))

    if omitidos:
        console.print("[bold yellow]⚠️  Filas evitadas por no contar con Número de Egresado (no aprobó):[/bold yellow]")
        for o in omitidos:
            console.print(f"  • [yellow]{o['apellido_nombre']}[/yellow] (DNI: {o['documento']}) - Fila {o['fila']}")
        console.print("")

    table = Table(title=f"Egresados Aprobados en {os.path.basename(excel_path)}")
    table.add_column("Índice", style="cyan")
    table.add_column("N° Egresado", style="yellow")
    table.add_column("Nombre y Apellido", style="bold white")
    table.add_column("DNI", style="magenta")
    table.add_column("Estado", style="green")

    for idx, eg in enumerate(egresados, 1):
        table.add_row(str(idx), eg["num_egresado"] or "-", eg["apellido_nombre"], eg["documento"], eg["estado"])

    console.print(table)
    
    selected_idx = Prompt.ask("\nSeleccione el número de índice del egresado a procesar", default="1")
    try:
        idx_num = int(selected_idx) - 1
        selected_eg = egresados[idx_num]
    except (ValueError, IndexError):
        console.print("[bold red]Índice inválido.[/bold red]")
        raise typer.Exit(code=1)

    generator = PPTXGenerator(template_path)
    
    if trayecto_data:
        t_line1 = trayecto_data["titulo_linea1"]
        t_line2 = trayecto_data["titulo_linea2"]
        horas = trayecto_data["horas_cursada"]
        res = trayecto_data["resolucion"]
        modules = trayecto_data["modulos"]
    else:
        t_line1 = data["especialidad"].title()
        t_line2 = data["especialidad"].title() if len(data["especialidad"]) > 40 else ""
        horas = "230"
        res = "Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE"
        modules = []

    titulo_data = TituloData(
        apellido_nombre=selected_eg['apellido_nombre'],
        documento=selected_eg['documento'],
        horas_cursada=horas,
        titulo_linea1=t_line1,
        titulo_linea2=t_line2,
        resolucion=res,
        emision_dia="02",
        emision_mes="Septiembre",
        emision_ano="25",
        modulos=modules,
        fecha_egreso=data["fecha_egreso"],
        numero_egresado=selected_eg["num_egresado"] or "000",
        numero_cpf=data["numero_cpf"],
        distrito_cpf=data["distrito_cpf"]
    )

    safe_name = selected_eg['apellido_nombre'].replace(" ", "_")
    output_pptx = os.path.join(output_dir, f"{selected_eg['num_egresado']}_{safe_name}.pptx")
    generator.generate(titulo_data, output_pptx)
    
    console.print(f"\n[bold green]✔ Certificado para {selected_eg['apellido_nombre']} generado en: {output_pptx}[/bold green]")
    
    pdf_path = convert_pptx_to_pdf(output_pptx, output_dir)
    if pdf_path:
        console.print(f"[bold cyan]✔ PDF generado en: {pdf_path}[/bold cyan]")


if __name__ == "__main__":
    app()
