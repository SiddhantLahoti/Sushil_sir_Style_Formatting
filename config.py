# ==============================================================================
# FILE PATHS
# ==============================================================================
STYLE_BREAKUP_PATH = "Style Breakup.xlsx"
COSTING_TEMPLATE_PATH = "Costing-Format_5_formats.xlsx"
OUTPUT_PATH = "Costing_Breakup_Generated.xlsx"

# ==============================================================================
# TEMPLATE STRUCTURE
# ==============================================================================
HEADER_LAST_ROW = 4
SR_NO_COL = "A"
STYLE_NO_COL = "B"
SPACER_ROWS = 1

# Template definitions from Costing-Format_5_formats.xlsx
# Each format block is 9 rows tall (start_row to end_row)
FORMAT_BLOCK_RANGES = {
    "ring":     {"start_row": 5,  "end_row": 13},
    "earring":  {"start_row": 14, "end_row": 22},
    "pendant":  {"start_row": 23, "end_row": 31},
    "necklace": {"start_row": 32, "end_row": 40},
    "bracelet": {"start_row": 41, "end_row": 49},
}

# Category name normalization (maps 'bangle' and plurals to target formats)
CATEGORY_MAPPING = {
    # Ring
    "ring": "ring",
    "rings": "ring",
    # Earring
    "earring": "earring",
    "earrings": "earring",
    # Pendant
    "pendant": "pendant",
    "pendants": "pendant",
    # Necklace
    "necklace": "necklace",
    "necklaces": "necklace",
    "neck piece": "necklace",
    "neckpiece": "necklace",
    # Bracelet & Bangle
    "bracelet": "bracelet",
    "bracelets": "bracelet",
    "bolo bracelet": "bracelet",
    "bolobracelet": "bracelet",
    "bangle": "bracelet",
    "bangles": "bracelet",
}

# Destination columns for diamond fields
TEMPLATE_DIAMOND_COLS = {
    "sieve": "M",          # SIEVE
    "type_shape": "N",     # Type/ Shape
    "qty": "P",            # QTY
    "each_dia_wt": "Q",    # Each Dia Wt
    "setting": "Z",        # Setting
}

# Source columns in Style Breakup
BREAKUP_HEADERS = {
    "cust": ["cust", "customer"],
    "style": ["style no", "style no.", "style_no", "style", "style code", "style #"],
    "rm_type": ["rm type", "rm_type", "type"],
    "sieve": ["sieve size", "sieve_size", "sieve"],
    "rm_code": ["rm code", "rm_code"],
    "setting": ["setting", "setting type"],
    "each_dia_wt": ["each dia wt", "each dia. wt", "each dia weight", "dia wt"],
    "rm_qty": ["rm qty", "rm_qty", "qty", "dia qty"],
}