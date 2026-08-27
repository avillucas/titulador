import os
import shutil
import pymupdf as fitz
from typing import Optional
from .models import TituloData

DEFAULT_DASH_LINE = "--------------------------------------------------------------------------"

def make_pdf_editable(pdf_path: str, data: TituloData, output_pdf_path: Optional[str] = None) -> str:
    """
    Transforms a standard PDF into an interactive, fillable PDF (AcroForm) pre-populated 
    with certificate data, adhering strictly to Circular 4-2020 layout standards.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if output_pdf_path is None:
        output_pdf_path = pdf_path

    doc = fitz.open(pdf_path)

    # Page 1: Certificate Front (Frente)
    if len(doc) > 0:
        p1 = doc[0]

        # Calculate dynamic font sizes
        len_t1 = len(data.titulo_linea1 or "")
        size_t1 = 9.5 if len_t1 > 42 else (10.5 if len_t1 > 34 else 12)

        len_t2 = len(data.titulo_linea2 or "")
        size_t2 = 9 if len_t2 > 35 else (10 if len_t2 > 26 else 12)

        len_nom = len(data.apellido_nombre or "")
        size_nom = 12 if len_nom > 32 else 14

        fields_page1 = [
            ("apellido_nombre", fitz.Rect(123.3, 165.0, 511.4, 189.2), "Helv-Bold", size_nom, data.apellido_nombre.upper()),
            ("documento", fitz.Rect(100.8, 196.0, 185.2, 217.8), "Helv", 12, data.documento),
            ("horas_cursada", fitz.Rect(415.8, 219.0, 460.8, 240.8), "Helv", 12, data.horas_cursada),
            ("titulo_linea1", fitz.Rect(83.7, 219.5, 404.3, 241.3), "Helv", size_t1, data.titulo_linea1),
            ("titulo_linea2", fitz.Rect(308.9, 241.5, 550.8, 263.3), "Helv", size_t2, data.titulo_linea2),
            ("resolucion", fitz.Rect(67.1, 264.0, 534.0, 285.8), "Helv", 12, data.resolucion),
            ("emision_dia", fitz.Rect(235.8, 304.0, 292.1, 325.8), "Helv", 12, data.emision_dia),
            ("emision_mes", fitz.Rect(303.3, 304.0, 466.4, 325.8), "Helv", 12, data.emision_mes),
            ("emision_ano", fitz.Rect(494.6, 304.0, 550.9, 325.8), "Helv", 12, data.emision_ano),
        ]

        # Redact static text under field locations to prevent bleed-through
        for _, rect, _, _, _ in fields_page1:
            p1.add_redact_annot(rect, fill=(1, 1, 1))
        p1.apply_redactions()

        # Add interactive AcroForm text widgets
        for field_name, rect, font_name, font_size, value in fields_page1:
            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.field_name = field_name
            widget.field_value = str(value or "")
            widget.rect = rect
            widget.text_font = font_name
            widget.text_fontsize = font_size
            widget.text_color = [0, 0, 0]
            p1.add_widget(widget)

    # Page 2: Certificate Back (Anverso)
    if len(doc) > 1:
        p2 = doc[1]

        module_rects = [
            fitz.Rect(33.3, 33.0, 297.7, 52.4),    # Slot 1
            fitz.Rect(33.3, 66.8, 297.7, 86.2),    # Slot 2
            fitz.Rect(33.3, 94.9, 297.7, 114.3),   # Slot 3
            fitz.Rect(33.3, 126.1, 297.7, 145.5),  # Slot 4
            fitz.Rect(33.3, 156.8, 297.7, 176.2),  # Slot 5
            fitz.Rect(308.9, 33.0, 573.3, 52.4),   # Slot 6
            fitz.Rect(308.9, 66.8, 573.3, 86.2),   # Slot 7
            fitz.Rect(308.9, 94.9, 573.3, 114.3),  # Slot 8
            fitz.Rect(308.9, 126.1, 573.3, 145.5), # Slot 9
            fitz.Rect(308.9, 156.8, 573.3, 176.2), # Slot 10
        ]

        fields_page2 = []
        for i, rect in enumerate(module_rects):
            val = data.modulos[i].strip() if i < len(data.modulos) and data.modulos[i].strip() else DEFAULT_DASH_LINE
            mod_size = 9 if len(val) > 42 else 10
            fields_page2.append((f"modulo_{i+1}", rect, "Helv", mod_size, val))

        fields_page2.extend([
            ("fecha_egreso", fitz.Rect(112.1, 198.0, 286.5, 219.8), "Helv", 12, data.fecha_egreso),
            ("numero_egresado", fitz.Rect(376.4, 198.0, 573.3, 219.8), "Helv", 12, data.numero_egresado),
            ("numero_cpf", fitz.Rect(89.6, 331.0, 286.5, 352.8), "Helv", 12, data.numero_cpf),
            ("distrito_cpf", fitz.Rect(365.2, 331.0, 562.1, 352.8), "Helv", 12, data.distrito_cpf),
        ])

        # Redact static text under field locations
        for _, rect, _, _, _ in fields_page2:
            p2.add_redact_annot(rect, fill=(1, 1, 1))
        p2.apply_redactions()

        # Add interactive AcroForm text widgets
        for field_name, rect, font_name, font_size, value in fields_page2:
            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.field_name = field_name
            widget.field_value = str(value or "")
            widget.rect = rect
            widget.text_font = font_name
            widget.text_fontsize = font_size
            widget.text_color = [0, 0, 0]
            p2.add_widget(widget)

    # Save output PDF using temporary file to allow overwriting input path safely
    target_dir = os.path.dirname(os.path.abspath(output_pdf_path))
    os.makedirs(target_dir, exist_ok=True)
    temp_path = os.path.join(target_dir, f"_tmp_{os.path.basename(output_pdf_path)}")

    doc.save(temp_path)
    doc.close()
    shutil.move(temp_path, output_pdf_path)

    return output_pdf_path
