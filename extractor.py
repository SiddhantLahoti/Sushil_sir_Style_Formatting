
import io
import posixpath
import xml.etree.ElementTree as ET
import zipfile
import openpyxl
from openpyxl.drawing.image import Image
from config import BREAKUP_HEADERS, CATEGORY_MAPPING

def extract_images_from_xlsx(file_path):
    """
    Extracts floating images, legacy cell images, and modern Office 365
    'Place in Cell' (Rich Data) images mapped directly to (row, col).
    """
    images = []
    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            file_names = set(zf.namelist())

            # -------------------------------------------------------------
            # 1. FLOATING IMAGES (xl/drawings/drawing*.xml)
            # -------------------------------------------------------------
            drawing_files = [f for f in file_names if f.startswith("xl/drawings/drawing") and f.endswith(".xml")]
            for draw_path in drawing_files:
                rels_path = posixpath.join(
                    posixpath.dirname(draw_path), "_rels", posixpath.basename(draw_path) + ".rels"
                )
                if rels_path not in file_names:
                    continue

                rels_tree = ET.fromstring(zf.read(rels_path))
                rel_map = {
                    rel.attrib.get("Id"): posixpath.normpath(
                        posixpath.join(posixpath.dirname(draw_path), rel.attrib.get("Target", ""))
                    )
                    for rel in rels_tree
                }

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
                                row_idx = int(r_elem.text) + 1
                                col_idx = int(c_elem.text) + 1
                                target_path = rel_map[r_id]

                                if target_path in file_names:
                                    img = Image(io.BytesIO(zf.read(target_path)))
                                    if img.width and img.height:
                                        max_dim = 130
                                        scale = min(max_dim / img.width, max_dim / img.height, 1.0)
                                        img.width = int(img.width * scale)
                                        img.height = int(img.height * scale)
                                    images.append({"row": row_idx, "col": col_idx, "img": img})

            # -------------------------------------------------------------
            # 2. MICROSOFT 365 "PLACE IN CELL" (Rich Data Architecture)
            # -------------------------------------------------------------
            rich_rel_file = "xl/richData/richValueRel.xml"
            rich_rels_file = "xl/richData/_rels/richValueRel.xml.rels"
            metadata_file = "xl/metadata.xml"

            if rich_rels_file in file_names and rich_rel_file in file_names:
                # Map rId -> media path
                rels_tree = ET.fromstring(zf.read(rich_rels_file))
                rel_map = {}
                for rel in rels_tree:
                    target = rel.attrib.get("Target", "")
                    rel_map[rel.attrib.get("Id")] = posixpath.normpath(
                        posixpath.join("xl/richData", target)
                    )

                # Map 0-indexed position -> media path
                rich_rel_tree = ET.fromstring(zf.read(rich_rel_file))
                rich_images = []
                for rel_elem in rich_rel_tree.iter():
                    if rel_elem.tag.split("}")[-1] == "rel":
                        r_id = next((v for k, v in rel_elem.attrib.items() if k.endswith("id")), None)
                        if r_id and r_id in rel_map:
                            rich_images.append(rel_map[r_id])

                # Read valueMetadata index mapping (vm attribute on cells)
                vm_to_image_idx = {}
                if metadata_file in file_names:
                    meta_tree = ET.fromstring(zf.read(metadata_file))
                    vm_counter = 1
                    for bk in meta_tree.iter():
                        if bk.tag.split("}")[-1] == "bk":
                            for rc in bk:
                                if rc.tag.split("}")[-1] == "rc":
                                    val_idx = int(rc.attrib.get("v", 0))
                                    vm_to_image_idx[vm_counter] = val_idx
                                    vm_counter += 1

                # Map cells with vm attribute to actual image
                sheet_files = [f for f in file_names if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]
                from openpyxl.utils import coordinate_to_tuple
                for sheet_path in sheet_files:
                    sheet_tree = ET.fromstring(zf.read(sheet_path))
                    for c_elem in sheet_tree.iter():
                        if c_elem.tag.split("}")[-1] == "c":
                            vm_attr = c_elem.attrib.get("vm")
                            if vm_attr:
                                vm_val = int(vm_attr)
                                img_idx = vm_to_image_idx.get(vm_val, vm_val - 1)
                                if 0 <= img_idx < len(rich_images):
                                    media_path = rich_images[img_idx]
                                    if media_path in file_names:
                                        cell_ref = c_elem.attrib.get("r")
                                        r_idx, c_idx = coordinate_to_tuple(cell_ref)
                                        img = Image(io.BytesIO(zf.read(media_path)))
                                        if img.width and img.height:
                                            max_dim = 130
                                            scale = min(max_dim / img.width, max_dim / img.height, 1.0)
                                            img.width = int(img.width * scale)
                                            img.height = int(img.height * scale)
                                        images.append({"row": r_idx, "col": c_idx, "img": img})

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
        raise ValueError(f"Missing required columns: {missing}")

    products = []
    max_row = ws.max_row
    r = header_row_idx + 1

    while r <= max_row:
        cust_val = ws.cell(row=r, column=col_map["cust"]).value
        style_val = ws.cell(row=r, column=col_map["style"]).value

        if cust_val is not None and str(cust_val).strip() != "" and style_val is not None and str(style_val).strip() != "":
            style_num = str(style_val).strip()
            
            # 1. Match image: Check directly in-cell first, then nearby floating images
            prod_image = None
            matched_item = None

            # 1. Direct cell match: only in Column O (col 15)
            for item in sheet_images:
                if item["col"] == 15 and r <= item["row"] <= r + 12:
                    matched_item = item
                    break

            # 2. Proximity match: only anchor to Column O (col 15)
            if not matched_item:
                best_dist = float("inf")
                for item in sheet_images:
                    if item["col"] == 15:
                        if r <= item["row"] <= r + 12:
                            dist = abs(item["row"] - r)
                            if dist < best_dist:
                                best_dist = dist
                                matched_item = item

            if matched_item:
                sheet_images.remove(matched_item)
                prod_image = matched_item["img"]
                
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


def extract_products_from_sketch(file_source, category_map=None):
    cat_mapping = category_map or CATEGORY_MAPPING

    if hasattr(file_source, "seek"):
        file_source.seek(0)

    sheet_images = extract_images_from_xlsx(file_source)

    if hasattr(file_source, "seek"):
        file_source.seek(0)
    wb = openpyxl.load_workbook(file_source, data_only=True)
    ws = wb.active

    products = []
    max_row = ws.max_row

    for r in range(1, max_row + 1):
        cell_a = ws.cell(row=r, column=1).value
        if cell_a and "sketch no" in str(cell_a).strip().lower():
            style_val = ws.cell(row=r, column=3).value
            if not style_val:
                continue
            style_num = str(style_val).strip()

            raw_cat = str(ws.cell(row=r + 15, column=16).value or "").strip().lower()
            category = cat_mapping.get(raw_cat, "ring")

            # Extract Sketch Image (Column O is col 15, Column P is col 16)
            prod_image = None
            matched_item = None

            # 1. Direct cell match: only in Column O (col 15)
            for item in sheet_images:
                if item["col"] == 15 and r <= item["row"] <= r + 12:
                    matched_item = item
                    break

            # 2. Proximity match: only anchor to Column O (col 15)
            if not matched_item:
                best_dist = float("inf")
                for item in sheet_images:
                    if item["col"] == 15:
                        if r <= item["row"] <= r + 12:
                            dist = abs(item["row"] - r)
                            if dist < best_dist:
                                best_dist = dist
                                matched_item = item

            if matched_item:
                sheet_images.remove(matched_item)
                prod_image = matched_item["img"]

            diamond_rows = []
            curr_r = r + 2
            first_shape = ws.cell(row=curr_r, column=4).value
            first_qty = ws.cell(row=curr_r, column=8).value
            if (first_shape is None or str(first_shape).strip() == "") and (first_qty is None or str(first_qty).strip() == ""):
                curr_r += 1

            while curr_r <= max_row:
                qty_val = ws.cell(row=curr_r, column=8).value
                if qty_val is None or str(qty_val).strip() == "":
                    break

                col_c = ws.cell(row=curr_r, column=3).value
                raw_shape = ws.cell(row=curr_r, column=4).value
                shape_map = {"RDFC": "RD", "RDSC": "SC RD"}
                col_d = shape_map.get(str(raw_shape).strip().upper(), raw_shape) if raw_shape is not None else raw_shape
                col_e = ws.cell(row=curr_r, column=5).value
                col_f = ws.cell(row=curr_r, column=6).value
                col_g = ws.cell(row=curr_r, column=7).value
                col_h = qty_val
                col_j = ws.cell(row=curr_r, column=10).value
                col_k = ws.cell(row=curr_r, column=11).value

                setting_parts = [str(x).strip() for x in (col_j, col_k) if x is not None and str(x).strip() != ""]
                setting_val = " ".join(setting_parts) if setting_parts else None

                diamond_rows.append({
                    "mm_size": col_f,
                    "sieve": col_e,
                    "type_shape": col_d,
                    "dia_qlty": col_c,
                    "qty": col_h,
                    "each_dia_wt": col_g,
                    "setting": setting_val,
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

    wb.close()
    return products