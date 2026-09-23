import streamlit as st
import config
from generator import generate_costing_sheet
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="Costing Breakup Generator", layout="wide")
st.title("💎 Costing Breakup Generator")

# ==============================================================================
# SIDEBAR: DYNAMIC COORDINATE & MAPPING SETTINGS
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Column & Row Settings")
    st.caption("Adjust these if any rows or columns shift in your input file.")

    # 1. Style Breakup Input Controls
    with st.expander("📥 Style Breakup (Input Settings)", expanded=False):
        st.subheader("Header Detection Keywords")
        style_kw = st.text_input("Style No Header Keywords", value=", ".join(config.BREAKUP_HEADERS["style"]))
        rm_type_kw = st.text_input("RM Type Header Keywords", value=", ".join(config.BREAKUP_HEADERS["rm_type"]))
        sieve_kw = st.text_input("Sieve Header Keywords", value=", ".join(config.BREAKUP_HEADERS["sieve"]))
        rm_code_kw = st.text_input("RM Code Header Keywords", value=", ".join(config.BREAKUP_HEADERS["rm_code"]))
        qty_kw = st.text_input("Quantity Header Keywords", value=", ".join(config.BREAKUP_HEADERS["rm_qty"]))
        wt_kw = st.text_input("Dia Wt Header Keywords", value=", ".join(config.BREAKUP_HEADERS["each_dia_wt"]))
        setting_kw = st.text_input("Setting Header Keywords", value=", ".join(config.BREAKUP_HEADERS["setting"]))
        cust_kw = st.text_input("Customer Header Keywords", value=", ".join(config.BREAKUP_HEADERS["cust"]))

        st.divider()
        st.subheader("Row & Column Shifts")
        cat_offset = st.number_input(
            "Category Row Offset (relative to Style row)", 
            value=2, 
            help="Default is 2 (i.e. Category is 2 rows below Style Number)"
        )
        img_col_letter = st.text_input(
            "Product Image Column", 
            value="B", 
            help="Column in the breakup file where photos are located"
        ).upper()

    # 2. Costing Template Output Controls
    with st.expander("📤 Costing Template (Output Settings)", expanded=False):
        header_last_row = st.number_input("Header Last Row", value=config.HEADER_LAST_ROW, min_value=1)
        spacer_rows = st.number_input("Spacer Rows Between Products", value=config.SPACER_ROWS, min_value=0)
        sr_col = st.text_input("Serial No Column", value=config.SR_NO_COL).upper()
        style_col = st.text_input("Style No Column", value=config.STYLE_NO_COL).upper()

        st.divider()
        st.subheader("Destination Diamond Columns")
        col_sieve = st.text_input("Sieve Size Col", value=config.TEMPLATE_DIAMOND_COLS["sieve"]).upper()
        col_shape = st.text_input("Type/Shape Col", value=config.TEMPLATE_DIAMOND_COLS["type_shape"]).upper()
        col_qty = st.text_input("Qty Col", value=config.TEMPLATE_DIAMOND_COLS["qty"]).upper()
        col_wt = st.text_input("Dia Wt Col", value=config.TEMPLATE_DIAMOND_COLS["each_dia_wt"]).upper()
        col_setting = st.text_input("Setting Col", value=config.TEMPLATE_DIAMOND_COLS["setting"]).upper()

        st.divider()
        st.subheader("Format Block Row Ranges")
        format_blocks = {}
        for cat, rng in config.FORMAT_BLOCK_RANGES.items():
            c1, c2 = st.columns(2)
            s_row = c1.number_input(f"{cat.title()} Start", value=rng["start_row"], key=f"fmt_{cat}_start")
            e_row = c2.number_input(f"{cat.title()} End", value=rng["end_row"], key=f"fmt_{cat}_end")
            format_blocks[cat] = {"start_row": int(s_row), "end_row": int(e_row)}


# ==============================================================================
# MAIN PAGE: SINGLE FILE UPLOAD
# ==============================================================================
uploaded_file = st.file_uploader(
    "Upload Style Breakup Excel File (.xlsx)",
    type=["xlsx"],
    help="Upload the client style breakup sheet containing products and photos."
)

if uploaded_file:
    if st.button("Generate Costing Breakup", type="primary", use_container_width=True):
        with st.spinner("Processing products, translating formulas, and placing images..."):
            try:
                # Build custom configurations from UI fields
                ui_config = {
                    "header_aliases": {
                        "style": [k.strip() for k in style_kw.split(",") if k.strip()],
                        "rm_type": [k.strip() for k in rm_type_kw.split(",") if k.strip()],
                        "sieve": [k.strip() for k in sieve_kw.split(",") if k.strip()],
                        "rm_code": [k.strip() for k in rm_code_kw.split(",") if k.strip()],
                        "rm_qty": [k.strip() for k in qty_kw.split(",") if k.strip()],
                        "each_dia_wt": [k.strip() for k in wt_kw.split(",") if k.strip()],
                        "setting": [k.strip() for k in setting_kw.split(",") if k.strip()],
                        "cust": [k.strip() for k in cust_kw.split(",") if k.strip()],
                    },
                    "category_offset": int(cat_offset),
                    "image_col": column_index_from_string(img_col_letter),
                    "header_last_row": int(header_last_row),
                    "spacer_rows": int(spacer_rows),
                    "sr_no_col": sr_col,
                    "style_no_col": style_col,
                    "template_diamond_cols": {
                        "sieve": col_sieve,
                        "type_shape": col_shape,
                        "qty": col_qty,
                        "each_dia_wt": col_wt,
                        "setting": col_setting,
                    },
                    "format_block_ranges": format_blocks,
                    "category_mapping": config.CATEGORY_MAPPING,
                }

                output_stream = generate_costing_sheet(
                    style_breakup_file=uploaded_file,
                    user_config=ui_config,
                )

                st.success("Costing sheet generated successfully!")
                st.download_button(
                    label="📥 Download Costing Breakup (.xlsx)",
                    data=output_stream,
                    file_name="Costing_Breakup_Generated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error: {e}")