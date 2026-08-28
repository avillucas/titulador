import os
import re
import base64
import subprocess
from typing import List, Tuple
from src.models import TituloData

DEFAULT_DASH_LINE = "--------------------------------------------------------------------------"

def get_image_as_data_uri(file_path: str) -> str:
    """Converts a local image file to a base64 Data URI for self-contained HTML."""
    if not os.path.exists(file_path):
        return ""
    ext = os.path.splitext(file_path)[1].lower().replace('.', '')
    mime = 'image/png' if ext == 'png' else 'image/jpeg'
    with open(file_path, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('utf-8')
    return f"data:{mime};base64,{encoded}"


def split_certificacion_title(title: str, max_line1: int = 38) -> Tuple[str, str]:
    """Splits a certification title so part 1 fits on Line 1 and part 2 wraps to Line 2."""
    title = title.strip()
    if not title or len(title) <= max_line1:
        return title, ""

    words = title.split()
    l1_words = []
    curr_len = 0
    for w in words:
        if curr_len + len(w) + (1 if l1_words else 0) <= max_line1:
            l1_words.append(w)
            curr_len += len(w) + 1
        else:
            break

    if l1_words and len(" ".join(l1_words)) >= max_line1 - 10:
        part1 = " ".join(l1_words)
        part2 = title[len(part1):].strip()
        return part1, part2

    cut = max_line1
    part1 = title[:cut] + "-"
    part2 = title[cut:].lstrip()
    return part1, part2


def format_module_text(module_str: str, max_single_line_chars: int = 52, max_line_chars: int = 42) -> str:
    """Formats module text with dot-testing per Circular 04-2020 f) and g)."""
    module_str = module_str.strip()
    if not module_str or module_str.startswith("----"):
        return DEFAULT_DASH_LINE

    if len(module_str) <= max_single_line_chars:
        num_dots = max(3, 60 - len(module_str))
        return f"{module_str} {'.' * num_dots}"

    match = re.search(r'^(.*?)\s*(\([A-Z0-9\.]+\))?$', module_str)
    if match and match.group(2):
        denom, code = match.group(1).strip(), match.group(2).strip()
    else:
        denom = module_str
        code = ""

    if len(denom) <= max_line_chars:
        num_dots = max(3, 60 - len(denom))
        line1 = f"{denom} {'.' * num_dots}"
        line2 = f" {code}" if code else ""
        return f"{line1}\n{line2}".strip()

    words = denom.split()
    l1_words = []
    curr = 0
    for w in words:
        if curr + len(w) + (1 if l1_words else 0) <= max_line_chars:
            l1_words.append(w)
            curr += len(w) + 1
        else:
            break

    if l1_words:
        part1 = " ".join(l1_words)
        part2 = denom[len(part1):].strip()
    else:
        part1 = denom[:max_line_chars]
        part2 = denom[max_line_chars:].strip()

    num_dots = max(3, 60 - len(part1))
    line1 = f"{part1} {'.' * num_dots}"
    line2 = f" {part2} {code}".strip()
    return f"{line1}\n {line2}".strip()


class HTMLGenerator:
    def __init__(self, template_path: str = None, assets_dir: str = None):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if template_path is None:
            template_path = os.path.join(base_dir, 'templates', 'template_certificate.html')
        if assets_dir is None:
            assets_dir = os.path.join(base_dir, 'templates', 'assets')

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found at {template_path}")

        self.template_path = template_path
        self.assets_dir = assets_dir

        self.frente_bg_path = os.path.join(self.assets_dir, 'frente.jpg')
        self.dorso_bg_path = os.path.join(self.assets_dir, 'dorso.jpg')

    def generate(self, data: TituloData, output_path: str, embed_images: bool = True) -> str:
        with open(self.template_path, 'r', encoding='utf-8') as f:
            template_content = f.read()

        if embed_images:
            bg_frente = get_image_as_data_uri(self.frente_bg_path)
            bg_dorso = get_image_as_data_uri(self.dorso_bg_path)
        else:
            rel_dir = os.path.dirname(os.path.abspath(output_path))
            bg_frente = os.path.relpath(self.frente_bg_path, rel_dir)
            bg_dorso = os.path.relpath(self.dorso_bg_path, rel_dir)

        cert_part1, cert_part2 = split_certificacion_title(data.titulo_linea2, max_line1=32)


        if cert_part2:
            cert_line2_full = f"{cert_part2} {data.resolucion}"
        else:
            cert_line2_full = str(data.resolucion)

        nombre_val = str(data.apellido_nombre).upper()
        font_size_nombre = "12pt" if len(nombre_val) > 32 else "14pt"

        trayecto_val = str(data.titulo_linea1)
        font_size_trayecto = "10.5pt" if len(trayecto_val) > 42 else "11.5pt"

        # Modules slot preparation (slots 0..9)
        slot_values = {}
        for idx in range(10):
            if idx < len(data.modulos) and data.modulos[idx].strip():
                formatted_mod = format_module_text(data.modulos[idx].strip(), max_line_chars=44)
                slot_values[f"slot_{idx}"] = formatted_mod
            else:
                slot_values[f"slot_{idx}"] = DEFAULT_DASH_LINE

        # Perform template replacements
        replacements = {
            "{{ bg_frente }}": bg_frente,
            "{{ bg_dorso }}": bg_dorso,
            "{{ apellido_nombre }}": nombre_val,
            "{{ font_size_nombre }}": font_size_nombre,
            "{{ documento }}": str(data.documento),
            "{{ horas_cursada }}": str(data.horas_cursada),
            "{{ titulo_linea1 }}": trayecto_val,
            "{{ font_size_trayecto }}": font_size_trayecto,
            "{{ cert_linea1 }}": cert_part1,
            "{{ cert_linea2 }}": cert_line2_full,
            "{{ emision_dia }}": str(data.emision_dia),
            "{{ emision_mes }}": str(data.emision_mes),
            "{{ emision_ano }}": str(data.emision_ano),
            "{{ fecha_egreso }}": str(data.fecha_egreso),
            "{{ numero_egresado }}": str(data.numero_egresado),
            "{{ numero_cpf }}": str(data.numero_cpf),
            "{{ distrito_cpf }}": str(data.distrito_cpf),
        }
        for idx in range(10):
            replacements[f"{{{{ slot_{idx} }}}}"] = slot_values[f"slot_{idx}"]

        rendered = template_content
        for key, val in replacements.items():
            rendered = rendered.replace(key, str(val))

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(rendered)

        return output_path

    def export_to_pdf(self, html_path: str, pdf_path: str) -> str:
        """Converts an HTML certificate file to PDF using headless Chrome/Chromium."""
        abs_html = os.path.abspath(html_path)
        abs_pdf = os.path.abspath(pdf_path)
        os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)

        cmd = [
            "google-chrome",
            "--headless",
            "--disable-gpu",
            "--no-pdf-header-footer",
            f"--print-to-pdf={abs_pdf}",
            f"file://{abs_html}"
        ]

        try:
            res = subprocess.run(cmd, capture_output=True, text=True, check=True)
            return abs_pdf
        except Exception as e:
            # Fallback to standard chromium if available
            try:
                cmd[0] = "chromium"
                subprocess.run(cmd, capture_output=True, text=True, check=True)
                return abs_pdf
            except Exception as e2:
                raise RuntimeError(f"Error converting HTML to PDF via headless browser: {e}")
