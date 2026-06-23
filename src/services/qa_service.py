# src/services/qa_service.py

from collections import defaultdict
from typing import Dict, List, Tuple
from config.fields import AT_FIELDS
from config.content_types import MONTHLY_OVERVIEW_MAPPING
from core.dates import parse_date
from core.text_formatting import normalize_str
from services.planner_service import hub_role

# Master list de todos los Content Types válidos unificados (DC + DJR)
VALID_CONTENT_TYPES = {
    "full episode", "exclusive", "editorial short", "lift", "musical lift",
    "magical moment", "compilation", "live stream", "podcast", "vertical short",
    "interactive", "sneak peek", "videoclip", "event", "exclusive short",
    "scene lifts / sneak peeks / videoclips"
}

def run_qa_on_records(records: List[Dict]) -> Dict:
    """
    Evaluates Airtable records against business rules to generate
    high-signal Blockers, Warnings, and Info logs.
    """
    qa_results = {
        "blockers": [],
        "warnings": [],
        "info": []
    }
    
    # Estructura para validar capacidad máxima: {(fecha, hub): cuenta}
    daily_hub_counts = defaultdict(int)

    for rec in records:
        f = rec.get("fields", rec)
        
        raw_title = f.get(AT_FIELDS["youtube_title"], f.get(AT_FIELDS["title"], ""))
        ct_raw = f.get(AT_FIELDS["content_type"], "")
        hub_raw = f.get(AT_FIELDS["owner_hub"], "")
        
        # ─── FILTRO DE FILAS GHOST / VACÍAS ───
        if not raw_title and not ct_raw and not hub_raw:
            continue

        # Configurar nombre visible para los logs
        title = normalize_str(raw_title) if raw_title else "Untitled Asset"
        ct_clean = normalize_str(ct_raw).strip().lower()
        role = hub_role(hub_raw)

        # ─── 1. BLOCKERS ───
        # A. Missing Live Date
        live_date = parse_date(f.get(AT_FIELDS["live_date"]))
        if not live_date:
            qa_results["blockers"].append(f"🛑 [Blocker] '{title}': Live Date is missing or invalid.")
            continue
            
        # B. Capacidad (Max 2 por slot diario por Hub)
        if role and ct_clean != "event":
            key = (live_date, role)
            daily_hub_counts[key] += 1
            if daily_hub_counts[key] == 3:
                qa_results["blockers"].append(
                    f"🛑 [Blocker] Capacidad excedida: Más de 2 contenidos programados para {role} el día {live_date.strftime('%Y-%m-%d')}."
                )

        # ─── 2. WARNINGS ───
        # A. Content Type Desconocido
        if ct_clean and ct_clean not in VALID_CONTENT_TYPES:
             qa_results["warnings"].append(f"⚠️ [Warning] '{title}': Unknown Content Type '{ct_raw}'.")

        # B. REGLAS EXCLUSIVAS PARA CONTENIDOS (Ignora los EVENT totalmente)
        if ct_clean != "event":
            # B.1 Missing Region
            region = f.get(AT_FIELDS["region"])
            if not region:
                 qa_results["warnings"].append(f"⚠️ [Warning] '{title}': Region/Tier is missing.")

            # B.2 Missing Property or Title (Solo para Programming)
            if role == "Programming":
                prop = f.get(AT_FIELDS["property"])
                if not prop or not raw_title:
                    qa_results["warnings"].append(f"⚠️ [Warning] '{title}': Missing Property or Title in Programming Hub.")

            # B.3 NUEVO: Validación de Metadatos Técnicos (TX ID, TC IN, TC OUT)
            tx_id = f.get(AT_FIELDS.get("tx_id", "TX ID"))
            tc_in = f.get(AT_FIELDS.get("tc_in", "TC IN"))
            tc_out = f.get(AT_FIELDS.get("tc_out", "TC OUT"))
            
            if not tx_id or not tc_in or not tc_out:
                qa_results["warnings"].append(
                    f"⚠️ [Warning] '{title}': Faltan metadatos técnicos críticos (TX ID, TC IN o TC OUT) en Airtable."
                )

            # B.4 NUEVO: Alerta de Compliance MENA (Bloqueo si es NO PASS)
            # Nota: Si consideras que 'NO PASS' debe impedir la creación del Excel por completo,
            # puedes mover este 'append' a la sección de 'blockers' de arriba.
            compliance_mena = f.get(AT_FIELDS.get("compliance_mena", "Compliance Pass (MENA)"))
            if str(compliance_mena).strip().upper() == "NO PASS":
                qa_results["warnings"].append(
                    f"🚨 [CRITICAL COMPLIANCE] '{title}': Marcado como 'NO PASS' para MENA. ¡Revisar urgente!"
                )

            # B.5 Alerta de desmapeo en el Monthly Overview
            mapped_category = ct_raw
            for k, v in MONTHLY_OVERVIEW_MAPPING.items():
                if k.lower() == ct_raw.lower():
                    mapped_category = v
                    break
            
            valid_overview_buckets = {"full episode", "exclusive", "exclusive short", "scene lifts / sneak peeks / videoclips", 
                                      "lift", "musical lift", "sneak peek", "videoclip", "podcast", "magical moment", 
                                      "interactive", "editorial short", "compilation", "live stream", "vertical short"}
            
            if mapped_category.lower() not in valid_overview_buckets:
                qa_results["warnings"].append(f"⚠️ [Warning] '{title}': Content Type '{ct_raw}' might not map to any Monthly Overview row.")

        # ─── 3. INFO ───
        if role == "Mediaplanning" and ("placeholder" in title.lower() or "tbd" in title.lower()):
            qa_results["info"].append(f"ℹ️ [Info] Mediaplanning placeholder detected on {live_date}.")

    return qa_results