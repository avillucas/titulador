import os
import subprocess
import shutil
from typing import Optional
from src.models import TituloData

def convert_pptx_to_pdf(pptx_path: str, output_dir: Optional[str] = None) -> Optional[str]:
    """
    Converts a PPTX file to PDF using LibreOffice/Soffice in headless mode.
    Returns path to converted PDF if successful, or None if LibreOffice is not available.
    """
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd:
        return None

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(pptx_path))
    else:
        os.makedirs(output_dir, exist_ok=True)

    cmd = [
        soffice_cmd,
        "--headless",
        "--convert-to", "pdf",
        pptx_path,
        "--outdir", output_dir
    ]

    try:
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        filename = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
        pdf_path = os.path.join(output_dir, filename)
        if os.path.exists(pdf_path):
            return pdf_path
    except (subprocess.SubprocessError, Exception) as e:
        print(f"PDF conversion error: {e}")
    return None


def print_pdf_a5(pdf_path: str, printer_name: Optional[str] = None) -> bool:
    """
    Sends PDF to default printer or specified printer with A5 paper size setting.
    """
    lp_cmd = shutil.which("lp")
    if not lp_cmd:
        print("Printer utility 'lp' not found on system.")
        return False

    cmd = [lp_cmd, "-o", "media=A5"]
    if printer_name:
        cmd.extend(["-d", printer_name])
    cmd.append(pdf_path)

    try:
        subprocess.run(cmd, check=True)
        return True
    except subprocess.SubprocessError as e:
        print(f"Print error: {e}")
        return False
