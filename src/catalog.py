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
    """Formats a raw trayecto dictionary from the catalog into structured data for TituloData per Circular 4-2020."""
    if not trayecto:
        return {}
        
    codigo = str(trayecto.get("codigo", "") or trayecto.get("Código", "")).strip()
    sector = str(trayecto.get("sector", "") or trayecto.get("Sector", "")).strip()
    nombre_trayecto = str(trayecto.get("trayecto_titulo", "") or trayecto.get("Nombre del Trayecto", "")).strip()
    certificacion_raw = str(trayecto.get("certificacion", "") or trayecto.get("Certificación", "") or nombre_trayecto).strip()
    
    # Clean certification text from trailing "según Resolución..." suffixes
    certificacion = re.sub(r"\s+según\s+Resolución.*$", "", certificacion_raw, flags=re.IGNORECASE).strip()
    certificacion = re.sub(r"\s+según\s+Res.*$", "", certificacion, flags=re.IGNORECASE).strip()
    
    # Line 1 (Nombre del Trayecto - Shape 88) and Line 2 (Certificación - Shape 89)
    t_line1 = nombre_trayecto
    t_line2 = certificacion if certificacion else nombre_trayecto

    # Hours calculation: Normative exception check (Circular 4-2020 Anexo, Sección 4.c)
    # Gasista 3ra, Gasista 2da, Montador Electricista, Electricista Instalador MUST use Hs. Cátedra.
    full_title_check = f"{nombre_trayecto} {certificacion}".lower()
    is_horas_catedra_exception = any(
        term in full_title_check
        for term in ["gasista de 3", "gasista 3", "gasista de 2", "gasista 2", "montador electricista", "electricista instalador"]
    )
    
    raw_mods = trayecto.get("modulos", []) or trayecto.get("Módulos Acreditables", [])
    if is_horas_catedra_exception:
        horas = str(trayecto.get("Hs. Cat. Trayecto", "") or trayecto.get("Hs. Reloj Trayecto", "")).strip()
        if not horas and isinstance(raw_mods, list):
            cat_sum = sum(int(m.get("Hs. Cat. Módulo", 0) or 0) for m in raw_mods if isinstance(m, dict))
            if cat_sum > 0:
                horas = str(cat_sum)
    else:
        horas = str(trayecto.get("Hs. Reloj Trayecto", "") or trayecto.get("Hs. Cat. Trayecto", "")).strip()
        if not horas and isinstance(raw_mods, list):
            rel_sum = sum(int(m.get("Hs. Reloj Módulo", 0) or 0) for m in raw_mods if isinstance(m, dict))
            if rel_sum > 0:
                horas = str(rel_sum)
            else:
                cat_sum = sum(int(m.get("Hs. Cat. Módulo", 0) or 0) for m in raw_mods if isinstance(m, dict))
                if cat_sum > 0:
                    horas = str(cat_sum)

    # Resolution formatting:
    # First jurisdictional, then national (if present) separated by ' / ' and padded with '--'
    res_jur = str(trayecto.get("resolucion_jurisdiccional", "") or trayecto.get("Res Jurisdiccional", "")).strip()
    res_nac = str(trayecto.get("resolucion_validez_nacional", "") or trayecto.get("Res Validez Nacional", "")).strip()

    clean_regex = r'^(resoluci[oó]n\s*(?:nro\.?|n°|nº)?|res\.\s*(?:nro\.?|n°|nº)?|res\s+(?:nro\.?|n°|nº))\s*:?\s*'
    clean_jur = re.sub(clean_regex, '', res_jur, flags=re.IGNORECASE).strip() if res_jur else ""
    clean_nac = re.sub(clean_regex, '', res_nac, flags=re.IGNORECASE).strip() if res_nac else ""

    clean_jur = clean_jur.rstrip("-").strip()
    clean_nac = clean_nac.rstrip("-").strip()

    res_parts = []
    if clean_jur:
        res_parts.append(clean_jur)
    if clean_nac:
        res_parts.append(clean_nac)

    if res_parts:
        combined_res = " / ".join(res_parts)
        resolucion = f"Resolución Nro. {combined_res}--"
    else:
        resolucion = ""
        
    # Format Accreditation Modules (Section 1 Anverso: "Módulo (CÓDIGO)")
    modulos: List[str] = []
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
            "utils/catalogo_FP_con_nacionales_2025_agrupado.json",
            "catalogo_FP_con_nacionales_2025_agrupado.json",
            "utils/CATALOGO_2025_FP_agrupado.json",
            "CATALOGO_2025_FP_agrupado.json",
            os.path.join(os.path.dirname(__file__), "..", "utils", "catalogo_FP_con_nacionales_2025_agrupado.json"),
            os.path.join(os.path.dirname(__file__), "catalogo_FP_con_nacionales_2025_agrupado.json"),
            os.path.join(os.path.dirname(__file__), "..", "utils", "CATALOGO_2025_FP_agrupado.json"),
            os.path.join(os.path.dirname(__file__), "CATALOGO_2025_FP_agrupado.json"),
        ]
        # PyInstaller sys._MEIPASS support for executable builds
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            candidates.insert(0, os.path.join(sys._MEIPASS, "utils", "catalogo_FP_con_nacionales_2025_agrupado.json"))
            candidates.insert(1, os.path.join(sys._MEIPASS, "catalogo_FP_con_nacionales_2025_agrupado.json"))
            candidates.insert(2, os.path.join(sys._MEIPASS, "utils", "CATALOGO_2025_FP_agrupado.json"))
            candidates.insert(3, os.path.join(sys._MEIPASS, "CATALOGO_2025_FP_agrupado.json"))

        for path in candidates:
            if os.path.exists(path):
                return os.path.abspath(path)
        return os.path.abspath("utils/catalogo_FP_con_nacionales_2025_agrupado.json")

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
            code = str(t.get("codigo", "") or t.get("Código", "")).strip()
            name = str(t.get("trayecto_titulo", "") or t.get("Nombre del Trayecto", "")).strip()
            sector = str(t.get("sector", "") or t.get("Sector", "")).strip()
            display = f"{code} - {name}"
            if sector:
                display += f" [{sector}]"
            items.append(display)
        return items

    def find_by_code(self, code: str) -> Optional[Dict[str, Any]]:
        code_norm = normalize_text(code)
        for item in self.trayectos:
            item_code = str(item.get("codigo", "") or item.get("Código", ""))
            if normalize_text(item_code) == code_norm:
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
            code = str(item.get("codigo", "") or item.get("Código", ""))
            name = str(item.get("trayecto_titulo", "") or item.get("Nombre del Trayecto", ""))
            disp = normalize_text(f"{code} - {name}")
            if norm_target in disp or disp in norm_target:
                return item
        return None

    def search_trayecto(self, especialidad: str) -> Optional[Dict[str, Any]]:
        """
        Searches for a trayecto matching especialidad from an exam record (acta).
        Matches against Código/codigo, Nombre del Trayecto/trayecto_titulo, and Certificación/certificacion.
        """
        if not especialidad or not self.trayectos:
            return None

        target_norm = normalize_text(especialidad)
        if not target_norm:
            return None

        # 1. Exact match on Code
        for item in self.trayectos:
            code = normalize_text(str(item.get("codigo", "") or item.get("Código", "")))
            if code and code == target_norm:
                return item

        # 2. Exact normalized match on Nombre del Trayecto / trayecto_titulo or Certificación
        for item in self.trayectos:
            name = normalize_text(str(item.get("trayecto_titulo", "") or item.get("Nombre del Trayecto", "")))
            cert = normalize_text(str(item.get("certificacion", "") or item.get("Certificación", "")))
            if name == target_norm or cert == target_norm:
                return item

        # 3. Substring match
        best_match = None
        best_score = 0

        for item in self.trayectos:
            name = normalize_text(str(item.get("trayecto_titulo", "") or item.get("Nombre del Trayecto", "")))
            cert = normalize_text(str(item.get("certificacion", "") or item.get("Certificación", "")))
            code = normalize_text(str(item.get("codigo", "") or item.get("Código", "")))

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
                name_words = set(normalize_text(str(item.get("trayecto_titulo", "") or item.get("Nombre del Trayecto", ""))).split())
                cert_words = set(normalize_text(str(item.get("certificacion", "") or item.get("Certificación", ""))).split())

                overlap_name = len(target_words.intersection(name_words))
                overlap_cert = len(target_words.intersection(cert_words))
                score = max(overlap_name, overlap_cert)

                if score > best_score:
                    best_score = score
                    best_match = item

        if best_score >= 2 and best_match:
            return best_match

        return None

