import re
from typing import List, Optional
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText


def normalize_str(s: Optional[str]) -> str:
    return str(s or "").strip()


def lower_clean(s: Optional[str]) -> str:
    return normalize_str(s).lower()


def bullet_join(items: List[str]) -> str:
    cleaned = [normalize_str(x) for x in items if normalize_str(x)]
    return "\n".join(cleaned)


def tc_in_out_str(tc_in: Optional[str], tc_out: Optional[str]) -> str:
    tin = normalize_str(tc_in)
    tout = normalize_str(tc_out)

    tin_lines = [line.strip() for line in tin.splitlines() if line.strip()]
    tout_lines = [line.strip() for line in tout.splitlines() if line.strip()]

    if len(tin_lines) > 1 or len(tout_lines) > 1:
        max_len = max(len(tin_lines), len(tout_lines))
        pairs = []
        for i in range(max_len):
            in_val = tin_lines[i] if i < len(tin_lines) else "—"
            out_val = tout_lines[i] if i < len(tout_lines) else "—"
            pairs.append(f"TC IN {in_val} / TC OUT {out_val}")
        return "\n".join(pairs)

    tin_val = tin_lines[0] if tin_lines else "—"
    tout_val = tout_lines[0] if tout_lines else "—"
    return f"{tin_val}\n{tout_val}"


def replace_month_in_title(text: str, month_name_upper: str) -> str:
    if not text:
        return text
    return text.replace("(MONTH)", month_name_upper).replace("MONTH", month_name_upper)

def join_comments(*parts: str) -> str:
    cleaned = [normalize_str(p) for p in parts if normalize_str(p)]
    return " | ".join(cleaned)

def process_multiline_branding(text: str, max_lines: int) -> list:
    """
    Toma un texto con saltos de línea (\n), lo divide y rellena con 
    strings vacíos ("") si faltan líneas, o recorta si sobran,
    asegurando que siempre devuelva exactamente 'max_lines' elementos.
    """
    if not text:
        return [""] * max_lines
        
    # Separar por saltos de línea reales de Airtable
    lines = str(text).split('\n')
    
    # Limpiar espacios extra
    lines = [line.strip() for line in lines]
    
    # Rellenar con blancos si faltan líneas
    if len(lines) < max_lines:
        lines.extend([""] * (max_lines - len(lines)))
        
    # Recortar si sobran líneas (evita desbordamientos en Excel)
    return lines[:max_lines]

def markdown_to_richtext(text: str):
    """
    Convierte un texto con asteriscos (*negrita*) en un objeto CellRichText
    para que Excel lo pinte con negritas parciales en la misma celda.
    """
    if not text or '*' not in str(text):
        return str(text) if text else ""
        
    chunks = str(text).split('*')
    rich_text = CellRichText()
    bold_font = InlineFont(b=True)
    
    for i, chunk in enumerate(chunks):
        if not chunk:
            continue
        # Los índices impares son los que estaban entre asteriscos
        if i % 2 == 1:
            rich_text.append(TextBlock(font=bold_font, text=chunk))
        else:
            rich_text.append(chunk)
            
    return rich_text