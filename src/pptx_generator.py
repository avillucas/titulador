import os
import pptx
from pptx.util import Pt
from typing import Dict, Any, List
from src.models import TituloData

DEFAULT_DASH_LINE = "--------------------------------------------------------------------------"

# Standard shape mapping based on template inspection
SHAPE_MAP_SLIDE1 = {
    85: "apellido_nombre",
    86: "documento",
    87: "horas_cursada",
    88: "titulo_linea1",
    89: "titulo_linea2",
    90: "resolucion",
    91: "emision_dia",
    92: "emision_mes",
    93: "emision_ano",
}

SHAPE_MAP_SLIDE2 = {
    99: 0,   # Slot 1  (Columna izquierda, fila 1)
    101: 1,  # Slot 2  (Columna izquierda, fila 2)
    103: 2,  # Slot 3  (Columna izquierda, fila 3)
    105: 3,  # Slot 4  (Columna izquierda, fila 4)
    107: 4,  # Slot 5  (Columna izquierda, fila 5)
    100: 5,  # Slot 6  (Columna derecha, fila 1)
    102: 6,  # Slot 7  (Columna derecha, fila 2)
    104: 7,  # Slot 8  (Columna derecha, fila 3)
    106: 8,  # Slot 9  (Columna derecha, fila 4)
    108: 9,  # Slot 10 (Columna derecha, fila 5)
    109: "fecha_egreso",
    110: "numero_egresado",
    111: "numero_cpf",
    112: "distrito_cpf",
}

def update_shape_text(shape, new_text: str):
    """Updates the shape's text frame preserving run styles where possible, scaling font size for long titles to avoid overflow."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    
    # Disable word wrap for title shapes (88, 89) and module shapes (99..108) to prevent vertical wrapping
    if shape.shape_id in (88, 89) or shape.shape_id in range(99, 109):
        tf.word_wrap = False

    if tf.paragraphs:
        p = tf.paragraphs[0]
        
        # Apply font scaling for long titles and modules
        if shape.shape_id == 89:  # titulo_linea2 (Width ~241 pt)
            if len(new_text) > 45:
                p.font.size = Pt(7.0)
            elif len(new_text) > 35:
                p.font.size = Pt(7.5)
            elif len(new_text) > 26:
                p.font.size = Pt(8.5)
            else:
                p.font.size = Pt(11.0)
        elif shape.shape_id == 88:  # titulo_linea1 (Width ~320 pt)
            if len(new_text) > 45:
                p.font.size = Pt(8.0)
            elif len(new_text) > 35:
                p.font.size = Pt(9.0)
            else:
                p.font.size = Pt(11.5)

        elif shape.shape_id in range(99, 109):  # Module slots
            if len(new_text) > 50:
                p.font.size = Pt(7.5)
            elif len(new_text) > 38:
                p.font.size = Pt(8.5)
            else:
                p.font.size = Pt(9.5)
        elif shape.shape_id == 85:  # apellido_nombre (Arial 14 Bold)
            p.font.bold = True
            if len(new_text) > 32:
                p.font.size = Pt(12.0)
            else:
                p.font.size = Pt(14.0)

        if p.runs:
            p.runs[0].text = new_text
            # Clear remaining runs in paragraph
            for r in p.runs[1:]:
                r.text = ""
            # Clear remaining paragraphs
            for p_extra in tf.paragraphs[1:]:
                p_extra.text = ""
        else:
            p.text = new_text
    else:
        tf.text = new_text



class PPTXGenerator:
    def __init__(self, template_path: str):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found at {template_path}")
        self.template_path = template_path

    def generate(self, data: TituloData, output_path: str) -> str:
        prs = pptx.Presentation(self.template_path)
        
        # Process Slide 1
        if len(prs.slides) > 0:
            slide1 = prs.slides[0]
            for shape in slide1.shapes:
                if shape.shape_id in SHAPE_MAP_SLIDE1:
                    field_name = SHAPE_MAP_SLIDE1[shape.shape_id]
                    val = getattr(data, field_name, "")
                    if field_name == "apellido_nombre" and val:
                        val = str(val).upper()
                    update_shape_text(shape, str(val))
                    
        # Process Slide 2
        if len(prs.slides) > 1:
            slide2 = prs.slides[1]
            mod_index = 0
            for shape in slide2.shapes:
                if shape.shape_id in SHAPE_MAP_SLIDE2:
                    target = SHAPE_MAP_SLIDE2[shape.shape_id]
                    if isinstance(target, int):
                        # Module slot 0..9
                        if target < len(data.modulos) and data.modulos[target].strip():
                            update_shape_text(shape, data.modulos[target].strip())
                        else:
                            update_shape_text(shape, DEFAULT_DASH_LINE)
                    else:
                        field_name = target
                        val = getattr(data, field_name, "")
                        update_shape_text(shape, str(val))

        # Ensure output directory exists
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        prs.save(output_path)
        return output_path
