import re
import datetime
import calendar
from typing import Optional, Tuple

MESES_ESPANOL = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

MESES_MAP = {m.lower(): i + 1 for i, m in enumerate(MESES_ESPANOL)}
# Add common variations
MESES_MAP["setiembre"] = 9

def parse_date_string(date_str: str) -> Optional[datetime.date]:
    """
    Parses a date string in various formats into a datetime.date object.
    Supports:
    - Textual Spanish: '14 de Julio de 2025', '14 de julio 2025', '14 de julio del 2025'
    - Formats with slashes/dashes: '14/07/2025', '14-07-2025', '14/7/25', '2025-07-14'
    """
    if not date_str:
        return None
    date_str = date_str.strip()

    # Textual Spanish pattern
    match_text = re.search(
        r'(\d{1,2})\s+de\s+([a-zA-ZáéíóúñÁÉÍÓÚÑ]+)(?:\s+de|\s+del)?\s+(\d{2,4})',
        date_str,
        re.IGNORECASE
    )
    if match_text:
        day = int(match_text.group(1))
        month_name = match_text.group(2).lower()
        year = int(match_text.group(3))
        if year < 100:
            year += 2000
        month = MESES_MAP.get(month_name)
        if month:
            try:
                return datetime.date(year, month, day)
            except ValueError:
                pass

    # ISO Format: YYYY-MM-DD
    match_iso = re.search(r'^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$', date_str)
    if match_iso:
        year = int(match_iso.group(1))
        month = int(match_iso.group(2))
        day = int(match_iso.group(3))
        try:
            return datetime.date(year, month, day)
        except ValueError:
            pass

    # Standard numeric format: DD/MM/YYYY or DD-MM-YYYY or DD/MM/YY
    match_num = re.search(r'^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$', date_str)
    if match_num:
        day = int(match_num.group(1))
        month = int(match_num.group(2))
        year = int(match_num.group(3))
        if year < 100:
            year += 2000
        try:
            return datetime.date(year, month, day)
        except ValueError:
            pass

    return None


def add_one_month(dt: datetime.date) -> datetime.date:
    """
    Adds exactly 1 month to a given datetime.date object.
    Handles month rollover (Dec -> Jan of next year) and day clamping for shorter months.
    """
    month = dt.month % 12 + 1
    year = dt.year + (dt.month // 12)
    max_day = calendar.monthrange(year, month)[1]
    day = min(dt.day, max_day)
    return datetime.date(year, month, day)


def format_emision_tuple(dt: datetime.date) -> Tuple[str, str, str]:
    """
    Formats a datetime.date object into a (dia, mes, ano) tuple for certificates.
    Example: (dia='14', mes='Agosto', ano='25')
    """
    dia_str = f"{dt.day:02d}"
    mes_str = MESES_ESPANOL[dt.month - 1]
    ano_str = f"{dt.year % 100:02d}"
    return dia_str, mes_str, ano_str


def calculate_default_emision_date(fecha_egreso_str: str) -> Tuple[str, str, str]:
    """
    Calculates the default certificate emission date (1 month after fecha_egreso).
    Returns (emision_dia, emision_mes, emision_ano).
    Falls back to ('02', 'Septiembre', '25') if fecha_egreso cannot be parsed.
    """
    dt = parse_date_string(fecha_egreso_str)
    if dt:
        dt_next = add_one_month(dt)
        return format_emision_tuple(dt_next)
    return "02", "Septiembre", "25"


def parse_emision_input(emision_str: str) -> Optional[Tuple[str, str, str]]:
    """
    Parses a user-provided emission date string into (emision_dia, emision_mes, emision_ano).
    Supports formats like '14/08/2025', '14 de Agosto de 2025', '14-08-25', etc.
    """
    dt = parse_date_string(emision_str)
    if dt:
        return format_emision_tuple(dt)
    return None
