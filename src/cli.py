import os
import sys
import typer
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm
from rich.panel import Panel

from src.models import TituloData
from src.excel_parser import ExcelParser
from src.pptx_generator import PPTXGenerator
from src.exporter import convert_pptx_to_pdf, print_pdf_a5

app = typer.Typer(name="titulador", help="Sistema de generación de certificados y títulos en A5")
console = Console()

DEFAULT_TEMPLATE = "ejemplos/Modelo base.pptx"
DEFAULT_EXCEL = "ejemplos/Acta de examen.xlsx"
OUTPUT_DIR = "output"

@app.command("batch")
def batch_generate(
    excel_path: str = typer.Option(DEFAULT_EXCEL, "--excel", "-e", help="Ruta al archivo Excel de Acta de examen"),
    template_path: str = typer.Option(DEFAULT_TEMPLATE, "--template", "-t", help="Ruta a la plantilla PPTX"),
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o", help="Carpeta de salida"),
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

    parser = ExcelParser(excel_path)
    data = parser.load_data()
    
    egresados = data["egresados"]
    aprobados = [e for e in egresados if e["estado"] == "Aprobado"]

    console.print(f"Curso/Especialidad: [cyan]{data['especialidad']}[/cyan]")
    console.print(f"Fecha Egreso: [cyan]{data['fecha_egreso']}[/cyan]")
    console.print(f"Total egresados aprobados: [bold yellow]{len(aprobados)}[/bold yellow]\n")

    generator = PPTXGenerator(template_path)
    
    # Modules for Carpinteria example if matching course
    modules = []
    if "carpintería" in data["especialidad"].lower():
        modules = [
            "Relaciones laborales y Orientación profesional. (MM0011.1)",
            "Tecnología de la Madera y materiales derivados.",
            "(MM 0012.1)",
            "Documentación técnica en carpintería. (MM0013.1)",
            "Trazado y corte de la madera y derivados. (MM0014.2)--",
            "Mecanizado, ensamble, unión y calidad de terminación en productos de madera y derivados. (MM0015.2)"
        ]

    table = Table(title="Certificados a Generar")
    table.add_column("N° Egresado", style="cyan")
    table.add_column("Nombre y Apellido", style="bold white")
    table.add_column("DNI", style="yellow")
    table.add_column("Archivo PPTX", style="green")

    generated_files = []

    for eg in aprobados:
        filename_base = f"{eg['num_egresado']}_{eg['apellido_nombre'].replace(' ', '_')}"
        out_pptx = os.path.join(output_dir, f"{filename_base}.pptx")
        
        titulo_data = TituloData(
            apellido_nombre=eg['apellido_nombre'],
            documento=eg['documento'],
            horas_cursada="230" if "carpintería" in data["especialidad"].lower() else "540",
            titulo_linea1=data["especialidad"].title(),
            titulo_linea2=data["especialidad"].title() if len(data["especialidad"]) > 40 else "",
            resolucion="Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE",
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
    console.print(f"\n[bold green]✔ ¡Se generaron {len(generated_files)} certificados exitosamente en {output_dir}![/bold green]")

@app.command("form")
def interactive_form(
    template_path: str = typer.Option(DEFAULT_TEMPLATE, "--template", "-t", help="Ruta a la plantilla PPTX"),
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o", help="Carpeta de salida")
):
    """Crea un certificado completando el formulario campo por campo en la consola."""
    console.print(Panel("[bold green]Titulador - Formulario Interactivo[/bold green]"))

    generator = PPTXGenerator(template_path)

    apellido_nombre = Prompt.ask("Apellido y Nombre")
    documento = Prompt.ask("DNI / Documento")
    horas_cursada = Prompt.ask("Cantidad de horas de cursada", default="230")
    titulo_linea1 = Prompt.ask("Nombre del Título (Línea 1)", default="Operadora/or de Carpintería y Fabricación de Mobiliario")
    titulo_linea2 = Prompt.ask("Nombre del Título (Línea 2)", default="")
    resolucion = Prompt.ask("Resolución", default="Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE")
    emision_dia = Prompt.ask("Emisión Día", default="02")
    emision_mes = Prompt.ask("Emisión Mes", default="Septiembre")
    emision_ano = Prompt.ask("Emisión Año", default="25")
    
    fecha_egreso = Prompt.ask("Fecha de Egreso", default="14 de Julio de 2025")
    numero_egresado = Prompt.ask("Número de Egresado", default="354")
    numero_cpf = Prompt.ask("Número CPF", default="412")
    distrito_cpf = Prompt.ask("Distrito CPF", default="Lomas de Zamora")

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

    console.print(f"\n[bold green]✔ Certificado generado en: {output_pptx}[/bold green]")
    
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
    output_dir: str = typer.Option(OUTPUT_DIR, "--outdir", "-o")
):
    """Muestra la lista de egresados del Excel y permite elegir uno para emitir su certificado."""
    if not os.path.exists(excel_path):
        console.print(f"[bold red]Error:[bold red] El archivo Excel no existe: {excel_path}")
        raise typer.Exit(code=1)

    parser = ExcelParser(excel_path)
    data = parser.load_data()
    egresados = data["egresados"]

    table = Table(title=f"Egresados en {os.path.basename(excel_path)}")
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
    
    modules = [
        "Relaciones laborales y Orientación profesional. (MM0011.1)",
        "Tecnología de la Madera y materiales derivados.",
        "(MM 0012.1)",
        "Documentación técnica en carpintería. (MM0013.1)",
        "Trazado y corte de la madera y derivados. (MM0014.2)--",
        "Mecanizado, ensamble, unión y calidad de terminación en productos de madera y derivados. (MM0015.2)"
    ]

    titulo_data = TituloData(
        apellido_nombre=selected_eg['apellido_nombre'],
        documento=selected_eg['documento'],
        horas_cursada="230",
        titulo_linea1=data["especialidad"].title(),
        titulo_linea2=data["especialidad"].title() if len(data["especialidad"]) > 40 else "",
        resolucion="Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE",
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
