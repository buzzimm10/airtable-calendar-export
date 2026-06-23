# src/config/content_types.py

"""
Centralized configuration for Airtable Content Types.
Handles aliases, abbreviations, and mapping to Excel Monthly Overview rows.
"""

# Mapeo de Content Types a las filas del Excel "Monthly Overview"
MONTHLY_OVERVIEW_MAPPING = {
    "Sneak Peek": "Scene Lifts / Sneak Peeks / Videoclips",
    "Videoclip": "Scene Lifts / Sneak Peeks / Videoclips",
    "Lift": "Scene Lifts / Sneak Peeks / Videoclips",
    "Musical Lift": "Scene Lifts / Sneak Peeks / Videoclips",
}

# Alias y normalizaciones básicas (Airtable value -> Normalized value)
CONTENT_TYPE_ALIASES = {
    "Exclusive": "Exclusive Short"
}

# Tipos de contenido que siempre deben ir en MAYÚSCULAS en el export
UPPERCASE_CONTENT_TYPES = {
    "Vertical Short"
}