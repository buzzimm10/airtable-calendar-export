from openpyxl.utils import column_index_from_string

TARGET_SHEET = "FULL CALENDAR YT DC EMEA"
MONTHLY_OVERVIEW_SHEET = "MONTHLY OVERVIEW YT DC EMEA"

DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

DAY_LAYOUT_LETTERS = {
    "Monday":    [("Programming", "D"), ("Programming", "E"), ("Mediaplanning", "F"), ("UK", "G")],
    "Tuesday":   [("Programming", "H"), ("Programming", "I"), ("Mediaplanning", "J"), ("UK", "K")],
    "Wednesday": [("Programming", "L"), ("Programming", "M"), ("Mediaplanning", "N"), ("UK", "O")],
    "Thursday":  [("Programming", "P"), ("Programming", "Q"), ("Mediaplanning", "R"), ("UK", "S")],
    "Friday":    [("Programming", "T"), ("Programming", "U"), ("Mediaplanning", "V"), ("UK", "W")],
    "Saturday":  [("Programming", "X"), ("Programming", "Y"), ("Mediaplanning", "Z"), ("UK", "AA")],
    "Sunday":    [("Programming", "AB"), ("Programming", "AC"), ("Mediaplanning", "AD"), ("UK", "AE")],
}

def _to_indices(layout_letters):
    return {
        day: [(role, column_index_from_string(letter)) for role, letter in cols]
        for day, cols in layout_letters.items()
    }

DAY_LAYOUT = _to_indices(DAY_LAYOUT_LETTERS)

FULL_CALENDAR_COUNTRY_ROW_RANGES = {
    1: (14, 30),
    2: (67, 83),
    3: (120, 136),
    4: (173, 189),
    5: (226, 242),
}

FULL_CALENDAR_WEEKLY_COUNT_COL = 33  # AG

OVERVIEW_DAY_COLS = {
    "Monday": 4,
    "Tuesday": 5,
    "Wednesday": 6,
    "Thursday": 7,
    "Friday": 8,
    "Saturday": 9,
    "Sunday": 10,
}

OVERVIEW_WEEK_LAYOUT = {
    1: {"header_row": 4, "date_row": 5, "event_rows": (6, 7), "base_row": 8, "end_row": 24},
    2: {"header_row": 26, "date_row": 27, "event_rows": (28, 29), "base_row": 30, "end_row": 46},
    3: {"header_row": 48, "date_row": 49, "event_rows": (50, 51), "base_row": 52, "end_row": 68},
    4: {"header_row": 70, "date_row": 71, "event_rows": (72, 73), "base_row": 74, "end_row": 90},
    5: {"header_row": 92, "date_row": 93, "event_rows": (94, 95), "base_row": 96, "end_row": 112},
}

BASE_DAY_TEMPLATE = {
    "Monday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Tuesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Wednesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Thursday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Friday": {"programming": 2, "mediaplanning": 1, "uk": 1},
    "Saturday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Sunday": {"programming": 2, "mediaplanning": 1, "uk": 1},
}

MAX_ASSETS_PER_HUB_PER_DAY = 2