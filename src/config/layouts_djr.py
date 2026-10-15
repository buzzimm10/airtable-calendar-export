from openpyxl.utils import column_index_from_string

DJR_TARGET_SHEET = "FULL CALENDAR YT DJR EMEA"
DJR_MONTHLY_OVERVIEW_SHEET = "MONTHLY OVERVIEW YT DJR EMEA"

DJR_WEEK_ANCHOR_CELLS = ["D7", "D69", "D131", "D193", "D255", "D317"]

DJR_WEEK_BLOCK_HEIGHT = 62

DJR_DAY_LAYOUT_LETTERS = {
    "Monday": [
        ("Programming", "D"),
        ("Extra", "E"),
        ("Mediaplanning", "F"),
        ("UK", "G"),
    ],
    "Tuesday": [
        ("Programming", "H"),
        ("Extra", "I"),
        ("Mediaplanning", "J"),
        ("UK", "K"),
    ],
    "Wednesday": [
        ("Programming", "L"),
        ("Extra", "M"),
        ("Mediaplanning", "N"),
        ("UK", "O"),
    ],
    "Thursday": [
        ("Programming", "P"),
        ("Extra", "Q"),
        ("Mediaplanning", "R"),
        ("UK", "S"),
    ],
    "Friday": [
        ("Programming", "T"),
        ("Extra", "U"),
        ("Mediaplanning", "V"),
        ("UK", "W"),
    ],
    "Saturday": [
        ("Programming", "X"),
        ("Extra", "Y"),
        ("Mediaplanning", "Z"),
        ("UK", "AA"),
    ],
    "Sunday": [
        ("Programming", "AB"),
        ("Extra", "AC"),
        ("Mediaplanning", "AD"),
        ("UK", "AE"),
    ],
}

def _to_indices(layout_letters):
    return {
        day: [(role, column_index_from_string(col)) for role, col in cols]
        for day, cols in layout_letters.items()
    }

DJR_DAY_LAYOUT = _to_indices(DJR_DAY_LAYOUT_LETTERS)

DJR_WEEK_LAYOUT = {
    1: {"start_row": 6, "date_row": 7, "deadline_row": 8, "owner_row": 9, "property_row": 10, "content_type_row": 11, "title_row": 12, "synopsis_row": 13, "strategy_row": 14, "country_rows": (15, 30), "tx_row": 31, "tc_row": 32, "branding_rows": (33, 46), "comments_rows": {"Programming": 47, "Mediaplanning": 48, "UK": 49}, "status_rows": (50, 65), "status_comments_row": 66},
    2: {"start_row": 68, "date_row": 69, "deadline_row": 70, "owner_row": 71, "property_row": 72, "content_type_row": 73, "title_row": 74, "synopsis_row": 75, "strategy_row": 76, "country_rows": (77, 92), "tx_row": 93, "tc_row": 94, "branding_rows": (95, 108), "comments_rows": {"Programming": 109, "Mediaplanning": 110, "UK": 111}, "status_rows": (112, 127), "status_comments_row": 128},
    3: {"start_row": 130, "date_row": 131, "deadline_row": 132, "owner_row": 133, "property_row": 134, "content_type_row": 135, "title_row": 136, "synopsis_row": 137, "strategy_row": 138, "country_rows": (139, 154), "tx_row": 155, "tc_row": 156, "branding_rows": (157, 170), "comments_rows": {"Programming": 171, "Mediaplanning": 172, "UK": 173}, "status_rows": (174, 189), "status_comments_row": 190},
    4: {"start_row": 192, "date_row": 193, "deadline_row": 194, "owner_row": 195, "property_row": 196, "content_type_row": 197, "title_row": 198, "synopsis_row": 199, "strategy_row": 200, "country_rows": (201, 216), "tx_row": 217, "tc_row": 218, "branding_rows": (219, 232), "comments_rows": {"Programming": 233, "Mediaplanning": 234, "UK": 235}, "status_rows": (236, 251), "status_comments_row": 252},
    5: {"start_row": 254, "date_row": 255, "deadline_row": 256, "owner_row": 257, "property_row": 258, "content_type_row": 259, "title_row": 260, "synopsis_row": 261, "strategy_row": 262, "country_rows": (263, 278), "tx_row": 279, "tc_row": 280, "branding_rows": (281, 294), "comments_rows": {"Programming": 295, "Mediaplanning": 296, "UK": 297}, "status_rows": (298, 313), "status_comments_row": 314},
    6: {"start_row": 316, "date_row": 317, "deadline_row": 318, "owner_row": 319, "property_row": 320, "content_type_row": 321, "title_row": 322, "synopsis_row": 323, "strategy_row": 324, "country_rows": (325, 340), "tx_row": 341, "tc_row": 342, "branding_rows": (343, 356), "comments_rows": {"Programming": 357, "Mediaplanning": 358, "UK": 359}, "status_rows": (360, 375), "status_comments_row": 376},
}

DJR_BRANDING_LAYOUT = {
    "watermark": [33, 34, 35, 36], 
    "bug": [37, 38, 39],           
    "copyright": 40,               
    "endcard": [41, 42, 43, 44, 45, 46] 
}

# Inyección dinámica de las filas de BRANDING para Disney Junior (Offset: 62)
# CAMBIO CRÍTICO: range(1, 7) para abarcar la semana 6
for week_num in range(1, 7):
    offset = (week_num - 1) * 62
    if week_num in DJR_WEEK_LAYOUT:
        DJR_WEEK_LAYOUT[week_num]["branding_rows"] = {
            "watermark": [33 + offset, 34 + offset, 35 + offset, 36 + offset],
            "bug": [37 + offset, 38 + offset, 39 + offset],
            "copyright": 40 + offset,  
            "endcard": [41 + offset, 42 + offset, 43 + offset, 44 + offset, 45 + offset, 46 + offset]
        }