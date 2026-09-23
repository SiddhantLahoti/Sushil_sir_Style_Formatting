import io
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from config import (
    BREAKUP_HEADERS,
    CATEGORY_MAPPING,
    COSTING_TEMPLATE_PATH,
    FORMAT_BLOCK_RANGES,
    HEADER_LAST_ROW,
    OUTPUT_PATH,
    SPACER_ROWS,
    SR_NO_COL,
    STYLE_BREAKUP_PATH,
    STYLE_NO_COL,
    TEMPLATE_DIAMOND_COLS,
)
from excel_utils import copy_cell
from extractor import extract_products_from_breakup


def generate_costing_sheet(style_breakup_file=STYLE_BREAKUP_PATH, user_config=None):
    cfg = user_config or {}

    # Extract UI / dynamic overrides with config fallbacks
    header_aliases = cfg.get("header_aliases", BREAKUP_HEADERS)
    cat_offset = cfg.get("category_offset", 2)
    img_col = cfg.get("image_col", 2)
    cat_mapping = cfg.get("category_mapping", CATEGORY_MAPPING)
    fmt_block_ranges = cfg.get("format_block_ranges", FORMAT_BLOCK_RANGES)
    dia_cols_config = cfg.get("template_diamond_cols", TEMPLATE_DIAMOND_COLS)
    header_end = cfg.get("header_last_row", HEADER_LAST_ROW)
    sr_col_letter = cfg.get("sr_no_col", SR_NO_COL)
    style_col_letter = cfg.get("style_no_col", STYLE_NO_COL)
    spacer_rows = cfg.get("spacer_rows", SPACER_ROWS)
    template_path = cfg.get("costing_template_path", COSTING_TEMPLATE_PATH)
    

    # 1. Parse products, images, and diamond rows from Style Breakup
    products = extract_products_from_breakup(
        file_source=style_breakup_file,
        header_aliases=header_aliases,
        category_offset=cat_offset,
        image_col=img_col,
        category_map=cat_mapping,
    )
    print(f"Extracted {len(products)} products from style breakup.")

    # 2. Load Costing Template from repository
    wb_template = openpyxl.load_workbook(template_path)
    ws_template = wb_template["Costing"] if "Costing" in wb_template.sheetnames else wb_template.active
    

    tpl_max_col = 45  # Columns A through AP
    sr_no_col_idx = column_index_from_string(sr_col_letter)
    style_col_idx = column_index_from_string(style_col_letter)
    dia_col_map = {k: column_index_from_string(v) for k, v in dia_cols_config.items()}
    dia_start_col = column_index_from_string("M")
    dia_end_col = column_index_from_string("AC")

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Costing"

    # Copy column widths from template
    for col_idx in range(1, tpl_max_col + 1):
        col_letter = get_column_letter(col_idx)
        if col_letter in ws_template.column_dimensions:
            ws_out.column_dimensions[col_letter].width = ws_template.column_dimensions[col_letter].width

    current_out_row = 1

    # 3. Write Header Section Once (Rows 1 to 4)
    for r in range(1, header_end + 1):
        if r in ws_template.row_dimensions:
            ws_out.row_dimensions[r].height = ws_template.row_dimensions[r].height
        for c in range(1, tpl_max_col + 1):
            copy_cell(ws_template.cell(row=r, column=c), ws_out.cell(row=r, column=c), row_delta=0)

    current_out_row = header_end + 1

    # 4. Generate Product Blocks
    for idx, prod in enumerate(products):
        if idx > 0:
            current_out_row += spacer_rows

        cat_key = prod["category"]
        fmt_range = fmt_block_ranges.get(cat_key, fmt_block_ranges.get("ring", {"start_row": 5, "end_row": 13}))
        fmt_start = fmt_range["start_row"]
        fmt_end = fmt_range["end_row"]
        fmt_height = fmt_end - fmt_start + 1

        block_target_start = current_out_row
        row_delta = block_target_start - fmt_start

        # Duplicate base format block
        for src_r in range(fmt_start, fmt_end + 1):
            tgt_r = src_r + row_delta
            if src_r in ws_template.row_dimensions:
                ws_out.row_dimensions[tgt_r].height = ws_template.row_dimensions[src_r].height

            for c in range(1, tpl_max_col + 1):
                src_cell = ws_template.cell(row=src_r, column=c)
                tgt_cell = ws_out.cell(row=tgt_r, column=c)
                copy_cell(src_cell, tgt_cell, row_delta=row_delta)

        # Update Serial Number (Column A) and Style Number (Column B)
        ws_out.cell(row=block_target_start, column=sr_no_col_idx).value = idx + 1
        ws_out.cell(row=block_target_start, column=style_col_idx).value = prod["style_no"]

        # Anchor Product Image into Column C
        if prod.get("image"):
            ws_out.add_image(prod["image"], f"C{block_target_start}")

        # Populate Diamond Information
        num_diamonds = len(prod["diamonds"])
        summary_cols = [
            column_index_from_string("S"),
            column_index_from_string("Y"),
            column_index_from_string("AC"),
        ]

        for dia_idx, dia in enumerate(prod["diamonds"]):
            tgt_dia_row = block_target_start + dia_idx

            # Duplicate formatting and translate formulas for diamond rows past row 1
            if dia_idx >= 1:
                dia_delta = tgt_dia_row - fmt_start

                if fmt_start in ws_template.row_dimensions:
                    ws_out.row_dimensions[tgt_dia_row].height = ws_template.row_dimensions[fmt_start].height

                for c in range(dia_start_col, dia_end_col + 1):
                    copy_cell(
                        ws_template.cell(row=fmt_start, column=c),
                        ws_out.cell(row=tgt_dia_row, column=c),
                        row_delta=dia_delta,
                    )
                    if c in summary_cols:
                        ws_out.cell(row=tgt_dia_row, column=c).value = None

            # Insert extracted diamond properties
            for field_name, val in dia.items():
                if field_name in dia_col_map:
                    ws_out.cell(row=tgt_dia_row, column=dia_col_map[field_name]).value = val

        # Compute full block height
        actual_product_end = max(block_target_start + fmt_height - 1, block_target_start + num_diamonds - 1)

        # Set summary formulas covering the entire product block height
        ws_out.cell(row=block_target_start, column=column_index_from_string("S")).value = f"=SUM(R{block_target_start}:R{actual_product_end})"
        ws_out.cell(row=block_target_start, column=column_index_from_string("Y")).value = f"=SUM(X{block_target_start}:X{actual_product_end})"
        ws_out.cell(row=block_target_start, column=column_index_from_string("AC")).value = f"=SUM(AB{block_target_start}:AB{actual_product_end})"

        print(
            f"Product {idx + 1}/{len(products)} -> Style: {prod['style_no']} | "
            f"Category: {cat_key.upper()} (from '{prod['raw_category']}') | "
            f"Diamonds: {num_diamonds} | Rows: {block_target_start}-{actual_product_end}"
        )

        current_out_row = actual_product_end + 1

    # If running locally from CLI, also save directly to disk
    if isinstance(style_breakup_file, str):
        wb_out.save(OUTPUT_PATH)
        print(f"\nSuccessfully generated {OUTPUT_PATH}")

    # Return in-memory buffer for Streamlit web download
    output_stream = io.BytesIO()
    wb_out.save(output_stream)
    output_stream.seek(0)
    return output_stream