
import io
import posixpath
import xml.etree.ElementTree as ET
import zipfile
import openpyxl
from openpyxl.drawing.image import Image
from config import BREAKUP_HEADERS, CATEGORY_MAPPING

def extract_images_from_xlsx(file_path):
    """
    Extracts floating images and their (row, col) cell coordinates directly
    from the .xlsx zip archive (bypassing openpyxl's read limitation).
    """
    images = []
    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            drawing_files = [f for f in zf.namelist() if f.startswith("xl/drawings/drawing") and f.endswith(".xml")]

            for draw_path in drawing_files:
                rels_path = posixpath.join(
                    posixpath.dirname(draw_path), "_rels", posixpath.basename(draw_path) + ".rels"
                )
                if rels_path not in zf.namelist():
                    continue

                # Map relationship IDs to media paths
                rels_tree = ET.fromstring(zf.read(rels_path))
                rel_map = {
                    rel.attrib.get("Id"): posixpath.normpath(
                        posixpath.join(posixpath.dirname(draw_path), rel.attrib.get("Target", ""))
                    )
                    for rel in rels_tree
                }

                # Parse anchors for row/column coordinates and image embed IDs
                draw_tree = ET.fromstring(zf.read(draw_path))
                for anchor in draw_tree:
                    tag = anchor.tag.split("}")[-1]
                    if tag in ("twoCellAnchor", "oneCellAnchor"):
                        from_elem = None
                        blip_elem = None
                        for elem in anchor.iter():
                            el_tag = elem.tag.split("}")[-1]
                            if el_tag == "from":
                                from_elem = elem
                            elif el_tag == "blip":
                                blip_elem = elem

                        if from_elem is not None and blip_elem is not None:
                            r_elem = next((e for e in from_elem if e.tag.split("}")[-1] == "row"), None)
                            c_elem = next((e for e in from_elem if e.tag.split("}")[-1] == "col"), None)
                            r_id = next((v for k, v in blip_elem.attrib.items() if k.endswith("embed")), None)

                            if r_elem is not None and c_elem is not None and r_id in rel_map:
                                row_idx = int(r_elem.text) + 1  # 1-indexed row
                                col_idx = int(c_elem.text) + 1  # 1-indexed col (Col B = 2)
                                target_path = rel_map[r_id]

                                if target_path in zf.namelist():
                                    img = Image(io.BytesIO(zf.read(target_path)))
                                    # Scale down high-res photos so they fit neatly in Column C
                                    if img.width and img.height:
                                        max_dim = 130
                                        scale = min(max_dim / img.width, max_dim / img.height, 1.0)
                                        img.width = int(img.width * scale)
                                        img.height = int(img.height * scale)

                                    images.append({"row": row_idx, "col": col_idx, "img": img})
    except Exception as e:
        print(f"Warning: Could not extract images: {e}")
    return images

def _get_image_position(img):
    anchor = getattr(img, "anchor", None)
    if hasattr(anchor, "_from") and anchor._from is not None:
        # Drawing markers are 0-indexed (col 0 = A, col 1 = B)
        return anchor._from.row + 1, anchor._from.col + 1
    if isinstance(anchor, str):
        from openpyxl.utils import coordinate_to_tuple
        return coordinate_to_tuple(anchor)
    return None, None


def _match_column(header_cell_value, alias_dict):
    val = str(header_cell_value or "").strip().lower()
    for field_key, aliases in alias_dict.items():
        if val in aliases or any(a in val for a in aliases):
            return field_key
    return None


def extract_products_from_breakup(
    file_source,
    header_aliases=None,
    category_offset=2,
    image_col=2,
    category_map=None,
):
    header_aliases = header_aliases or BREAKUP_HEADERS
    cat_mapping = category_map or CATEGORY_MAPPING

    # Reset stream pointer if it's an uploaded file object
    if hasattr(file_source, "seek"):
        file_source.seek(0)

    sheet_images = extract_images_from_xlsx(file_source)

    if hasattr(file_source, "seek"):
        file_source.seek(0)
    wb = openpyxl.load_workbook(file_source, data_only=True)
    ws = wb.active
    

    col_map = {}
    header_row_idx = None

    for row in ws.iter_rows(values_only=False):
        for cell in row:
            matched_key = _match_column(cell.value, header_aliases)
            if matched_key:
                col_map[matched_key] = cell.column
                if header_row_idx is None:
                    header_row_idx = cell.row

        if "cust" in col_map and "style" in col_map and "rm_type" in col_map:
            break

    required = ["cust", "style", "rm_type", "sieve", "rm_code", "setting", "each_dia_wt", "rm_qty"]
    missing = [f for f in required if f not in col_map]
    if missing:
        wb.close()
        raise ValueError(f"Missing required columns in {file_path}: {missing}")

    products = []
    max_row = ws.max_row
    r = header_row_idx + 1

    while r <= max_row:
        cust_val = ws.cell(row=r, column=col_map["cust"]).value
        style_val = ws.cell(row=r, column=col_map["style"]).value

        if cust_val is not None and str(cust_val).strip() != "" and style_val is not None and str(style_val).strip() != "":
            style_num = str(style_val).strip()
            # Match image located in Column B (col 2) near this style row
            prod_image = None
            matched_item = None
            best_dist = float("inf")
            for item in sheet_images:
                if item["col"] in (image_col - 1, image_col, image_col + 1):  # Catches Col B even if slightly overlapping A or C
                    dist = abs(item["row"] - r)
                    if dist < best_dist and dist <= 6:
                        best_dist = dist
                        matched_item = item

            if matched_item:
                sheet_images.remove(matched_item)
                prod_image = matched_item["img"]
                
            # Category is located 2 rows below the style number row in the style column
            raw_cat = " ".join(str(ws.cell(row=r + category_offset, column=col_map["style"]).value or "").strip().lower().split())
            category = cat_mapping.get(raw_cat, "ring")

            diamond_rows = []
            curr_r = r
            while curr_r <= max_row:
                rm_type_val = ws.cell(row=curr_r, column=col_map["rm_type"]).value
                if rm_type_val is None or str(rm_type_val).strip() == "":
                    break

                diamond_rows.append({
                    "sieve": ws.cell(row=curr_r, column=col_map["sieve"]).value,
                    "type_shape": ws.cell(row=curr_r, column=col_map["rm_code"]).value,
                    "setting": ws.cell(row=curr_r, column=col_map["setting"]).value,
                    "each_dia_wt": ws.cell(row=curr_r, column=col_map["each_dia_wt"]).value,
                    "qty": ws.cell(row=curr_r, column=col_map["rm_qty"]).value,
                })
                curr_r += 1

            products.append({
                "style_no": style_num,
                "category": category,
                "raw_category": raw_cat,
                "diamond_count": len(diamond_rows),
                "diamonds": diamond_rows,
                "image": prod_image,
            })

            r = max(curr_r, r + 3)
        else:
            r += 1

    wb.close()
    return products