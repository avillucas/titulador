import os
import sys
import json
import re
import unicodedata
from typing import List, Dict, Any, Optional, Tuple

def normalize_text(text: str) -> str:
    """Normalizes text by removing accents, lowercasing, and stripping whitespace."""
    if not text:
        return ""
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text

def split_title(title: str, max_len: int = 40) -> Tuple[str, str]:
    """Splits a title string into line 1 and line 2 for pptx formatting if it exceeds max_len."""
    if not title:
        return "", ""
    title = title.strip()
    if len(title) <= max_len:
        return title, ""
    
    words = title.split()
    half = len(title) / 2
    best_split = len(words) // 2
    curr_len = 0
    
    for i, w in enumerate(words):
        curr_len += len(w) + 1
        if curr_len >= half:
            best_split = max(1, i + 1)
            break
            
    line1 = " ".join(words[:best_split])
    line2 = " ".join(words[best_split:])
    return line1, line2

def format_trayecto_data(trayecto: Dict[str, Any]) -> Dict[str, Any]:
    """Formats a raw trayecto dictionary from the catalog into structured data for TituloData."""
    if not trayecto:
        return {}
        
    codigo = str(trayecto.get("Código", "")).strip()
    sector = str(trayecto.get("Sector", "")).strip()
    nombre_trayecto = str(trayecto.get("Nombre del Trayecto", "")).strip()
    certificacion = str(trayecto.get("Certificación", "")).strip()
    
    # Preferred title text is Nombre del Trayecto, fallback to Certificación
    title_text = nombre_trayecto or certificacion
    t_line1, t_line2 = split_title(title_text)
    
    # Hours: Hs. Reloj Trayecto preferred, fallback to Hs. Cat. Trayecto
    horas = str(trayecto.get("Hs. Reloj Trayecto", "") or trayecto.get("Hs. Cat. Trayecto", "")).strip()
    
    # Resolution formatting
    res_raw = str(trayecto.get("Res Jurisdiccional", "")).strip()
    if res_raw:
        if not res_raw.lower().startswith("resoluci") and not res_raw.lower().startswith("res"):
            resolucion = f"Resolución Nro. {res_raw}"
        elif res_raw.lower().startswith("res") and not res_raw.lower().startswith("resoluci"):
            resolucion = f"Resolución Nro. {res_raw}"
        else:
            resolucion = res_raw
    else:
        resolucion = ""
        
    # Format Accreditation Modules
    modulos: List[str] = []
    raw_mods = trayecto.get("Módulos Acreditables", [])
    if isinstance(raw_mods, list):
        for m in raw_mods:
            if not isinstance(m, dict):
                continue
            denom = str(m.get("Denominación del Módulo", "")).strip()
            cod = str(m.get("Código de Planificación", "")).strip()
            if not denom:
                continue
            if not denom.endswith("."):
                denom_str = f"{denom}."
            else:
                denom_str = denom
            
            if cod:
                modulos.append(f"{denom_str} ({cod})")
            else:
                modulos.append(denom_str)
                
    return {
        "codigo": codigo,
        "sector": sector,
        "nombre_trayecto": nombre_trayecto,
        "certificacion": certificacion,
        "titulo_linea1": t_line1,
        "titulo_linea2": t_line2,
        "horas_cursada": horas,
        "resolucion": resolucion,
        "modulos": modulos
    }

class TrayectoCatalog:
    def __init__(self, catalog_path: Optional[str] = None):
        self.catalog_path = catalog_path or self._find_catalog_file()
        self.trayectos: List[Dict[str, Any]] = []
        self._load_catalog()

    def _find_catalog_file(self) -> str:
        candidates = [
            "utils/CATALOGO_2025_FP_agrupado.json",
            "CATALOGO_2025_FP_agrupado.json",
            os.path.join(os.path.dirname(__file__), "..", "utils", "CATALOGO_2025_FP_agrupado.json"),
            os.path.join(os.path.dirname(__file__), "CATALOGO_2025_FP_agrupado.json"),
        ]
        # PyInstaller sys._MEIPASS support for executable builds
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            candidates.insert(0, os.path.join(sys._MEIPASS, "utils", "CATALOGO_2025_FP_agrupado.json"))
            candidates.insert(1, os.path.join(sys._MEIPASS, "CATALOGO_2025_FP_agrupado.json"))

        for path in candidates:
            if os.path.exists(path):
                return os.path.abspath(path)
        return os.path.abspath("utils/CATALOGO_2025_FP_agrupado.json")

    def _load_catalog(self):
        if not os.path.exists(self.catalog_path):
            print(f"Advertencia: No se encontró el archivo de catálogo en '{self.catalog_path}'.")
            return
        try:
            with open(self.catalog_path, "r", encoding="utf-8") as f:
                self.trayectos = json.load(f)
        except Exception as e:
            print(f"Error al cargar el catálogo '{self.catalog_path}': {e}")

    def get_all_trayectos(self) -> List[Dict[str, Any]]:
        return self.trayectos

    def get_trayecto_display_list(self) -> List[str]:
        """Returns formatted list of strings suitable for GUI dropdowns or CLI selection."""
        items = []
        for t in self.trayectos:
            code = t.get("Código", "").strip()
            name = t.get("Nombre del Trayecto", "").strip()
            sector = t.get("Sector", "").strip()
            display = f"{code} - {name}"
            if sector:
                display += f" [{sector}]"
            items.append(display)
        return items

    def find_by_code(self, code: str) -> Optional[Dict[str, Any]]:
        code_norm = normalize_text(code)
        for item in self.trayectos:
            if normalize_text(str(item.get("Código", ""))) == code_norm:
                return item
        return None

    def find_by_display_name(self, display_name: str) -> Optional[Dict[str, Any]]:
        if not display_name:
            return None
        # Display name format: "MM11 - Operadora/or..."
        code_part = display_name.split(" - ")[0].strip()
        match = self.find_by_code(code_part)
        if match:
            return match
            
        norm_target = normalize_text(display_name)
        for item in self.trayectos:
            disp = normalize_text(f"{item.get('Código', '')} - {item.get('Nombre del Trayecto', '')}")
            if norm_target in disp or disp in norm_target:
                return item
        return None

    def search_trayecto(self, especialidad: str) -> Optional[Dict[str, Any]]:
        """
        Searches for a trayecto matching especialidad from an exam record (acta).
        Matches against Código, Nombre del Trayecto, and Certificación.
        """
        if not especialidad or not self.trayectos:
            return None

        target_norm = normalize_text(especialidad)
        if not target_norm:
            return None

        # 1. Exact match on Code
        for item in self.trayectos:
            code = normalize_text(str(item.get("Código", "")))
            if code and code == target_norm:
                return item

        # 2. Exact normalized match on Nombre del Trayecto or Certificación
        for item in self.trayectos:
            name = normalize_text(str(item.get("Nombre del Trayecto", "")))
            cert = normalize_text(str(item.get("Certificación", "")))
            if name == target_norm or cert == target_norm:
                return item

        # 3. Substring match
        best_match = None
        best_score = 0

        for item in self.trayectos:
            name = normalize_text(str(item.get("Nombre del Trayecto", "")))
            cert = normalize_text(str(item.get("Certificación", "")))
            code = normalize_text(str(item.get("Código", "")))

            score = 0
            if name and (target_norm in name or name in target_norm):
                score = 80
            elif cert and (target_norm in cert or cert in target_norm):
                score = 70
            elif code and code in target_norm:
                score = 60

            if score > best_score:
                best_score = score
                best_match = item

        if best_match:
            return best_match

        # 4. Token overlap matching
        target_words = set(target_norm.split())
        if target_words:
            for item in self.trayectos:
                name_words = set(normalize_text(str(item.get("Nombre del Trayecto", ""))).split())
                cert_words = set(normalize_text(str(item.get("Certificación", ""))).split())

                overlap_name = len(target_words.intersection(name_words))
                overlap_cert = len(target_words.intersection(cert_words))
                score = max(overlap_name, overlap_cert)

                if score > best_score:
                    best_score = score
                    best_match = item

        if best_score >= 2 and best_match:
            return best_match

        return None
