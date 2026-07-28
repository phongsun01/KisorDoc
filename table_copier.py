import re
from copy import deepcopy
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree

from config import AppConfig

TABLE_PLACEHOLDER_RE = re.compile(r"\{\{(DanhMucKoGia|DanhMuc)\}\}")

EXCEL_BORDER_MAP = {
    "thin": "single",
    "medium": "single",
    "thick": "single",
    "double": "double",
    "dashed": "dashed",
    "dotted": "dotted",
    "hair": "single",
    None: "none",
}

EXCEL_BORDER_SIZE = {
    "thin": 4,
    "medium": 8,
    "thick": 12,
    "double": 6,
    "dashed": 4,
    "dotted": 4,
    "hair": 2,
}


def copy_tables_for_file(
    doc_path: Path,
    config: AppConfig,
    goi_thau_id: str,
    tables_data: list[dict],
    xlsx_path: Path,
):
    doc = Document(str(doc_path))

    doc_tables = _collect_table_placeholders(doc)
    file_stem = doc_path.stem
    if file_stem.endswith("-Template"):
        file_stem = file_stem[: -len("-Template")]

    matching_tables = [
        t for t in tables_data
        if t["GoiThau_ID"] == goi_thau_id and _match_word(t["Word"], file_stem)
    ]

    placeholder_occurrences: dict[str, list] = {}
    for t in matching_tables:
        key = t["Name"]
        if key not in placeholder_occurrences:
            placeholder_occurrences[key] = []
        placeholder_occurrences[key].append(t)

    for placeholder_key, occurrences in placeholder_occurrences.items():
        doc_occurrences = doc_tables.get(placeholder_key, [])
        if not doc_occurrences:
            continue

        for i, occ in enumerate(occurrences):
            if i >= len(doc_occurrences):
                break
            para_element = doc_occurrences[i]
            sheet = occ.get("Sheet", "")
            range_spec = occ.get("Range", "A1")
            hide_cols = occ.get("Hide", "")

            if not sheet:
                continue

            ws_data = _read_excel_range(xlsx_path, sheet, range_spec, hide_cols)
            if ws_data is None:
                continue

            _insert_table_at_paragraph(doc, para_element, ws_data)

    doc.save(str(doc_path))


def _match_word(tables_word: str, file_stem: str) -> bool:
    t = tables_word.strip().lower()
    f = file_stem.strip().lower()
    return t == f or f.startswith(t) or t in f


def _collect_table_placeholders(doc: Document) -> dict[str, list]:
    occurrences: dict[str, list] = {}

    for para in doc.paragraphs:
        m = TABLE_PLACEHOLDER_RE.search(para.text)
        if m:
            key = m.group(1)
            wrapped = "{{" + key + "}}"
            if wrapped not in occurrences:
                occurrences[wrapped] = []
            occurrences[wrapped].append(para._element)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    m = TABLE_PLACEHOLDER_RE.search(para.text)
                    if m:
                        key = m.group(1)
                        wrapped = "{{" + key + "}}"
                        if wrapped not in occurrences:
                            occurrences[wrapped] = []
                        occurrences[wrapped].append(para._element)

    return occurrences


def _read_excel_range(xlsx_path: Path, sheet_name: str, range_spec: str, hide_cols: str):
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=False, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return None
        ws = wb[sheet_name]

        parsed = _parse_range(range_spec, ws)
        if parsed is None:
            wb.close()
            return None
        min_row, max_row, min_col, max_col = parsed

        hidden_cols = _parse_hidden_cols(hide_cols)
        visible_cols = [c for c in range(min_col, max_col + 1) if c not in hidden_cols]
        col_map = {orig: new for new, orig in enumerate(visible_cols)}

        merged_map = _build_merged_map(ws, min_row, max_row, min_col, max_col)
        visible_merged = {}
        for (r, c), (mr, mc, rs, cs) in merged_map.items():
            if c in hidden_cols:
                continue
            if mc in hidden_cols:
                new_mc = col_map.get(mc)
                if new_mc is None:
                    new_c = col_map.get(c)
                    visible_merged[(r, new_c)] = (r, new_c, rs, 1)
                continue
            new_r = r - min_row
            new_mr = mr - min_row
            new_mc = col_map.get(mc)
            new_c = col_map.get(c, new_mc)
            if new_mc is not None and new_c is not None:
                new_cs = max(1, cs - len([h for h in hidden_cols if mc < h <= mc + cs - 1]))
                visible_merged[(new_r, new_c)] = (new_mr, new_mc, rs, new_cs)

        row_heights = {}
        for r in range(min_row, max_row + 1):
            h = ws.row_dimensions[r].height
            if h:
                row_heights[r - min_row] = h

        col_widths = {}
        for c in visible_cols:
            w = ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width
            if w:
                col_widths[col_map[c]] = w

        data = {
            "rows": [],
            "merged": visible_merged,
            "row_heights": row_heights,
            "col_widths": col_widths,
        }

        for r in range(min_row, max_row + 1):
            row_data = []
            for c in visible_cols:
                cell = ws.cell(row=r, column=c)
                cell_info = {
                    "value": cell.value,
                    "font": cell.font,
                    "fill": cell.fill,
                    "border": cell.border,
                    "alignment": cell.alignment,
                    "number_format": cell.number_format,
                }
                row_data.append(cell_info)
            data["rows"].append(row_data)

        wb.close()
        return data

    except Exception:
        return None


def _parse_range(range_spec: str, ws):
    range_spec = range_spec.strip().upper()
    m = re.match(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d*))?$", range_spec)
    if not m:
        return None
    min_col = openpyxl.utils.column_index_from_string(m.group(1))
    min_row = int(m.group(2))
    if m.group(3):
        max_col = openpyxl.utils.column_index_from_string(m.group(3))
        if m.group(4):
            max_row = int(m.group(4))
        else:
            max_row = ws.max_row
    else:
        max_col = ws.max_column
        max_row = ws.max_row
    return min_row, max_row, min_col, max_col


def _parse_hidden_cols(hide_str: str) -> set[int]:
    if not hide_str or hide_str.strip() == "":
        return set()
    cols = set()
    for part in hide_str.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            cols.add(openpyxl.utils.column_index_from_string(part))
        except Exception:
            pass
    return cols


def _build_merged_map(ws, min_row, max_row, min_col, max_col):
    merged_map = {}
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row > max_row or merged_range.max_row < min_row
            or merged_range.min_col > max_col or merged_range.max_col < min_col
        ):
            continue
        mr = max(merged_range.min_row, min_row)
        mc = max(merged_range.min_col, min_col)
        rs = merged_range.max_row - merged_range.min_row + 1
        cs = merged_range.max_col - merged_range.min_col + 1
        for r in range(merged_range.min_row, merged_range.max_row + 1):
            for c in range(merged_range.min_col, merged_range.max_col + 1):
                if min_row <= r <= max_row and min_col <= c <= max_col:
                    merged_map[(r, c)] = (mr, mc, rs, cs)
    return merged_map


def _insert_table_at_paragraph(doc: Document, para_element, ws_data):
    rows_data = ws_data["rows"]
    if not rows_data:
        return
    nrows = len(rows_data)
    ncols = len(rows_data[0])
    if ncols == 0:
        return

    merged = ws_data.get("merged", {})
    row_heights = ws_data.get("row_heights", {})
    col_widths = ws_data.get("col_widths", {})

    tbl = _create_word_table_xml(nrows, ncols, rows_data, merged, row_heights, col_widths)

    parent = para_element.getparent()
    if parent is not None:
        idx = list(parent).index(para_element)
        parent.remove(para_element)
        parent.insert(idx, tbl)


def _create_word_table_xml(nrows, ncols, rows_data, merged, row_heights, col_widths):
    NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    table_xml = etree.Element(f"{{{NS}}}tbl")

    tblPr = etree.SubElement(table_xml, f"{{{NS}}}tblPr")
    tblStyle = etree.SubElement(tblPr, f"{{{NS}}}tblStyle")
    tblStyle.set(f"{{{NS}}}val", "TableGrid")
    tblW = etree.SubElement(tblPr, f"{{{NS}}}tblW")
    tblW.set(f"{{{NS}}}w", "5000")
    tblW.set(f"{{{NS}}}type", "pct")

    tblGrid = etree.SubElement(table_xml, f"{{{NS}}}tblGrid")
    for ci in range(ncols):
        gridCol = etree.SubElement(tblGrid, f"{{{NS}}}gridCol")
        cw = col_widths.get(ci)
        if cw:
            twips = int(cw * 96 * 914400 / 72 / 10000)
            gridCol.set(f"{{{NS}}}w", str(twips))
        else:
            gridCol.set(f"{{{NS}}}w", "1440")

    for ri in range(nrows):
        tr = etree.SubElement(table_xml, f"{{{NS}}}tr")

        trPr = etree.SubElement(tr, f"{{{NS}}}trPr")
        rh = row_heights.get(ri)
        if rh:
            trHeight = etree.SubElement(trPr, f"{{{NS}}}trHeight")
            trHeight.set(f"{{{NS}}}val", str(int(rh * 20)))
            trHeight.set(f"{{{NS}}}hRule", "atLeast")

        for ci in range(ncols):
            merged_key = (ri, ci)
            if merged_key in merged or any(
                m[0] <= ri < m[0] + m[2] and m[1] <= ci < m[1] + m[3]
                for mk, m in merged.items()
                if mk != merged_key and not (mk[0] == ri and mk[1] == ci)
            ):
                # Check if this cell is part of a merged range but not the master
                is_master = merged_key in merged
                if not is_master:
                    continue

            cell_data = rows_data[ri][ci] if ri < len(rows_data) and ci < len(rows_data[ri]) else None
            tc = _create_cell_xml(NS, cell_data, merged.get(merged_key))

            tr.append(tc)

    return table_xml


def _create_cell_xml(NS, cell_data, merge_info):
    tc = etree.Element(f"{{{NS}}}tc")

    tcPr = etree.SubElement(tc, f"{{{NS}}}tcPr")
    tcW = etree.SubElement(tcPr, f"{{{NS}}}tcW")
    tcW.set(f"{{{NS}}}w", "1440")
    tcW.set(f"{{{NS}}}type", "dxa")

    if merge_info:
        mr, mc, rs, cs = merge_info
        if cs > 1:
            gridSpan = etree.SubElement(tcPr, f"{{{NS}}}gridSpan")
            gridSpan.set(f"{{{NS}}}val", str(cs))
        if rs > 1:
            vMerge = etree.SubElement(tcPr, f"{{{NS}}}vMerge")
            vMerge.set(f"{{{NS}}}val", "restart")
    else:
        vMerge = etree.SubElement(tcPr, f"{{{NS}}}vMerge")
        vMerge.set(f"{{{NS}}}val", "continue")

    if cell_data:
        _apply_cell_formatting(tcPr, NS, cell_data)

    p = etree.SubElement(tc, f"{{{NS}}}p")
    pPr = etree.SubElement(p, f"{{{NS}}}pPr")

    if cell_data and cell_data.get("alignment"):
        align = cell_data["alignment"]
        if align.horizontal:
            jc = etree.SubElement(pPr, f"{{{NS}}}jc")
            jc.set(f"{{{NS}}}val", align.horizontal)

    if cell_data and cell_data.get("value") is not None:
        text_val = _format_cell_value(cell_data)
        r = etree.SubElement(p, f"{{{NS}}}r")
        rPr = etree.SubElement(r, f"{{{NS}}}rPr")
        if cell_data.get("font"):
            font = cell_data["font"]
            if font.name:
                rFonts = etree.SubElement(rPr, f"{{{NS}}}rFonts")
                rFonts.set(f"{{{NS}}}ascii", font.name)
                rFonts.set(f"{{{NS}}}hAnsi", font.name)
            if font.size:
                sz = etree.SubElement(rPr, f"{{{NS}}}sz")
                sz.set(f"{{{NS}}}val", str(int(font.size * 2)))
                szCs = etree.SubElement(rPr, f"{{{NS}}}szCs")
                szCs.set(f"{{{NS}}}val", str(int(font.size * 2)))
            if font.bold:
                etree.SubElement(rPr, f"{{{NS}}}b")
            if font.italic:
                etree.SubElement(rPr, f"{{{NS}}}i")
            if font.underline and font.underline != "none":
                u = etree.SubElement(rPr, f"{{{NS}}}u")
                u.set(f"{{{NS}}}val", font.underline)
            if font.color and font.color.rgb:
                color = etree.SubElement(rPr, f"{{{NS}}}color")
                color.set(f"{{{NS}}}val", str(font.color.rgb))
        t = etree.SubElement(r, f"{{{NS}}}t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = str(text_val)

    return tc


def _apply_cell_formatting(tcPr, NS, cell_data):
    if not cell_data:
        return

    fill = cell_data.get("fill")
    if fill and fill.fgColor and fill.fgColor.rgb and fill.fgColor.rgb != "00000000":
        shd = etree.SubElement(tcPr, f"{{{NS}}}shd")
        shd.set(f"{{{NS}}}val", "clear")
        shd.set(f"{{{NS}}}color", "auto")
        shd.set(f"{{{NS}}}fill", str(fill.fgColor.rgb))

    tcBorders = etree.SubElement(tcPr, f"{{{NS}}}tcBorders")
    border = cell_data.get("border")
    if border:
        for edge, edge_name in [
            ("top", "top"), ("bottom", "bottom"), ("left", "left"), ("right", "right")
        ]:
            b = getattr(border, edge, None)
            if b and b.style:
                be = etree.SubElement(tcBorders, f"{{{NS}}}{edge_name}")
                be.set(f"{{{NS}}}val", EXCEL_BORDER_MAP.get(b.style, "single"))
                be.set(f"{{{NS}}}sz", str(EXCEL_BORDER_SIZE.get(b.style, 4)))
                be.set(f"{{{NS}}}space", "0")
                if b.color and b.color.rgb:
                    be.set(f"{{{NS}}}color", str(b.color.rgb))
                else:
                    be.set(f"{{{NS}}}color", "000000")


def _format_cell_value(cell_data) -> str:
    value = cell_data.get("value")
    if value is None:
        return ""
    nf = cell_data.get("number_format", "")
    nf = str(nf) if nf else ""

    if isinstance(value, (int, float)):
        if "DD" in nf.upper() or "MM" in nf.upper() or "YY" in nf.upper():
            from datetime import timedelta, datetime
            base = datetime(1899, 12, 30)
            try:
                dt = base + timedelta(days=float(value))
                fmt = nf.replace("hh:mm:ss", "").strip()
                if not fmt:
                    fmt = "%d/%m/%Y"
                return dt.strftime(fmt)
            except Exception:
                return str(value)
        if "%" in nf:
            return f"{float(value) * 100:.2f}%"
        if "#,##" in nf or "#.##" in nf:
            return f"{float(value):,.0f}".replace(",", ".")
        return str(value)

    return str(value)
